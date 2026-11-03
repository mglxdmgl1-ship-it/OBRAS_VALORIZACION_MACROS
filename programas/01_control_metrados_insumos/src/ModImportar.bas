Attribute VB_Name = "ModImportar"
Option Explicit

'==========================================================================================
'  CONTROL DE METRADOS DIARIOS - IMPORTAR PRESUPUESTO Y A.C.U.
'  Lee libros de Excel exportados de S10, Delphin, Kennedy u otros programas: busca los títulos
'  de las columnas, los encabezados de cada partida ("Partida 01.01.01" o "01.01.01 DESCRIPCIÓN"),
'  el rendimiento, las secciones (mano de obra, materiales, equipos, subcontratos, subpartidas)
'  y los insumos. Las subpartidas se reemplazan por sus propios insumos cuando su análisis está
'  en el mismo archivo.
'==========================================================================================

' ---- presupuesto leído
Private eN As Long
Private eCod() As String, eDesc() As String, eUnd() As String, eMet() As Variant, ePU() As Variant, ePar() As Variant
Private eHoja As String, eDupl As Long
' ---- A.C.U. leído: bloques (partidas) e insumos
Private bN As Long
Private bCod() As String, bDesc() As String, bSub() As Boolean, bRend() As Double, bIni() As Long, bFin() As Long
Private bDescN() As String, bUnd() As String
Private tN As Long
Private tIns() As String, tUnd() As String, tTipo() As String, tCant() As Double, tPre() As Double, tCodI() As String
Private tPend() As Boolean
Private sHojas As String
' ---- columnas del encabezado del A.C.U.
Private cIns As Long, cUnd As Long, cCua As Long, cCan As Long, cPre As Long, cPar As Long, cCodI As Long


'========================================================================== BOTONES
Public Sub ImportarPresupuesto()
    Dim f As Variant, msg As String
    f = Application.GetOpenFilename("Libros de Excel (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", , _
                                    "Seleccione el archivo del PRESUPUESTO (S10, Delphin u otro)")
    If VarType(f) = vbBoolean Then Exit Sub
    If UltimaFila(ThisWorkbook.Worksheets("PRESUPUESTO"), 2, PP0, PPMAX) >= PP0 Then
        If MsgBox("La hoja PRESUPUESTO ya tiene datos. ¿Reemplazarlos por los del archivo elegido?" & vbCrLf & vbCrLf & _
                  "(los metrados digitados en el RESUMEN no se tocan; se conservan al presionar PREPARAR OBRA)", _
                  vbYesNo + vbQuestion, "Importar presupuesto") <> vbYes Then Exit Sub
    End If
    msg = ImportarPresupuestoDe(CStr(f))
    MsgBox msg, IIf(Left(msg, 8) = "No se pu", vbExclamation, vbInformation), "Importar presupuesto"
    If Left(msg, 8) <> "No se pu" Then ThisWorkbook.Worksheets("PRESUPUESTO").Activate
End Sub

Public Sub ImportarACU()
    Dim f As Variant, msg As String
    f = Application.GetOpenFilename("Libros de Excel (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb", , _
                                    "Seleccione el archivo de ANÁLISIS DE COSTOS UNITARIOS (S10, Delphin u otro)")
    If VarType(f) = vbBoolean Then Exit Sub
    If UltimaFila(ThisWorkbook.Worksheets("PRESUPUESTO"), 2, PP0, PPMAX) < PP0 Then
        If MsgBox("Todavía no hay PRESUPUESTO. Conviene importarlo primero para emparejar las partidas del A.C.U." & vbCrLf & _
                  vbCrLf & "¿Importar el A.C.U. de todas formas?", vbYesNo + vbQuestion, "Importar A.C.U.") <> vbYes Then Exit Sub
    ElseIf UltimaFila(ThisWorkbook.Worksheets("ACU"), 1, AC0, ACMAX) >= AC0 Then
        If MsgBox("La hoja ACU ya tiene datos. ¿Reemplazarlos por los del archivo elegido?", vbYesNo + vbQuestion, _
                  "Importar A.C.U.") <> vbYes Then Exit Sub
    End If
    msg = ImportarACUDe(CStr(f))
    MsgBox msg, IIf(Left(msg, 8) = "No se pu", vbExclamation, vbInformation), "Importar A.C.U."
    If Left(msg, 8) <> "No se pu" Then ThisWorkbook.Worksheets("ACU").Activate
End Sub

' pruebas automáticas: la ruta del archivo va en la celda H1 de la hoja DATOS
Public Sub ImportarPresupuestoSilencioso()
    gSilencio = True
    ThisWorkbook.Worksheets("DATOS").Range(D_MSGPREP).Value = ImportarPresupuestoDe(Txt(ThisWorkbook.Worksheets("DATOS").Range("H1").Value))
    gSilencio = False
End Sub

Public Sub ImportarACUSilencioso()
    gSilencio = True
    ThisWorkbook.Worksheets("DATOS").Range(D_MSGPREP).Value = ImportarACUDe(Txt(ThisWorkbook.Worksheets("DATOS").Range("H2").Value))
    gSilencio = False
End Sub


'========================================================================== ABRIR EL ARCHIVO
Private Function AbrirLibro(ByVal ruta As String, ByRef yaAbierto As Boolean) As Workbook
    Dim w As Workbook, sec As Long
    yaAbierto = False
    For Each w In Workbooks
        If LCase(w.FullName) = LCase(ruta) Then
            Set AbrirLibro = w
            yaAbierto = True
            Exit Function
        End If
    Next w
    On Error Resume Next
    sec = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Application.DisplayAlerts = False
    On Error GoTo 0
    Set AbrirLibro = Workbooks.Open(ruta, 0, True)
    On Error Resume Next
    If sec <> 0 Then Application.AutomationSecurity = sec
    Application.DisplayAlerts = True
    On Error GoTo 0
End Function

Private Function NombreArchivo(ByVal ruta As String) As String
    Dim p As Long
    p = InStrRev(Replace(ruta, "/", "\"), "\")
    NombreArchivo = Mid(ruta, p + 1)
End Function


'========================================================================== PRESUPUESTO
Public Function ImportarPresupuestoDe(ByVal ruta As String) As String
    Dim calcMode As Long, wb As Workbook, ws As Worksheet, ya As Boolean, mejor As Long, n As Long
    Dim bHo As String, bDu As Long, i As Long, nT As Long, nP As Long, cD As Double, dst As Worksheet
    Dim o() As Variant, g() As Variant, colG As Collection, v As Variant, k As Long
    Inicio calcMode, "Leyendo el presupuesto..."
    On Error GoTo Fallo
    If ruta = "" Then Err.Raise 3001, , "No se indicó el archivo."
    If Dir(ruta) = "" Then Err.Raise 3002, , "No se encontró el archivo: " & ruta
    Set wb = AbrirLibro(ruta, ya)
    mejor = 0
    bHo = ""
    For Each ws In wb.Worksheets
        n = LeerHojaPresupuesto(ws)
        If n > mejor Then
            mejor = n
            bHo = ws.Name
        End If
    Next ws
    ' se vuelve a leer la mejor hoja (sus datos quedan en las listas e*)
    If mejor > 0 Then mejor = LeerHojaPresupuesto(wb.Worksheets(bHo))
    bDu = eDupl
    If Not ya Then wb.Close False
    Set wb = Nothing
    If mejor = 0 Then Err.Raise 3003, , "No se encontró un presupuesto en el archivo (se buscan columnas ÍTEM, DESCRIPCIÓN, " & _
                                      "UND, METRADO, PRECIO y PARCIAL, con ítems como 01.01.01)."
    ' rendimientos ya registrados (vienen del A.C.U.): se conservan por ítem
    Set dst = ThisWorkbook.Worksheets("PRESUPUESTO")
    Set colG = New Collection
    v = dst.Range(dst.Cells(PP0, 1), dst.Cells(PPMAX, 7)).Value
    For i = 1 To UBound(v, 1)
        If Txt(v(i, 1)) <> "" And Num(v(i, 7)) > 0 Then
            If Buscar(colG, "G" & Clave(Txt(v(i, 1)))) = 0 Then colG.Add i, "G" & Clave(Txt(v(i, 1)))
        End If
    Next i
    ReDim o(1 To mejor, 1 To 7)
    For i = 1 To mejor
        o(i, 1) = eCod(i)
        o(i, 2) = eDesc(i)
        o(i, 3) = eUnd(i)
        o(i, 4) = eMet(i)
        o(i, 5) = ePU(i)
        o(i, 6) = ePar(i)
        k = Buscar(colG, "G" & Clave(eCod(i)))
        If k > 0 Then o(i, 7) = v(k, 7)
        If eUnd(i) <> "" Then
            nP = nP + 1
            cD = cD + Num(ePar(i))
        Else
            nT = nT + 1
        End If
    Next i
    dst.Range(dst.Cells(PP0, 1), dst.Cells(PPMAX, 9)).ClearContents
    dst.Range(dst.Cells(PP0, 1), dst.Cells(PPMAX, 3)).NumberFormat = "@"
    dst.Range(dst.Cells(PP0, 1), dst.Cells(PP0 + mejor - 1, 7)).Value = o
    ImportarPresupuestoDe = "Presupuesto importado de «" & NombreArchivo(ruta) & "» (hoja «" & bHo & "»)." & vbCrLf & vbCrLf & _
                            "• " & nT & " títulos y " & nP & " partidas." & vbCrLf & _
                            "• Costo directo (suma de partidas): S/ " & Format(cD, "#,##0.00") & "."
    If bDu > 0 Then ImportarPresupuestoDe = ImportarPresupuestoDe & vbCrLf & "• Se omitieron " & bDu & " fila(s) con ítem repetido."
    ImportarPresupuestoDe = ImportarPresupuestoDe & vbCrLf & vbCrLf & _
                            "Revise la hoja PRESUPUESTO y siga con IMPORTAR A.C.U. (o péguelo en la hoja ACU) y PREPARAR OBRA."
Salir:
    Fin calcMode
    Exit Function
Fallo:
    ImportarPresupuestoDe = "No se pudo importar el presupuesto." & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then
        If Not ya Then wb.Close False
    End If
    Resume Salir
End Function

Private Function EncabezadoPres(ByVal h As String, ByVal tipo As String) As Boolean
    Dim t As String
    t = Replace(Replace(h, ".", ""), " ", "")
    EncabezadoPres = False
    Select Case tipo
        Case "ITEM"
            EncabezadoPres = (t = "ITEM" Or t = "ITEMS" Or t = "COD" Or t = "CODIGO" Or t = "N°ITEM" Or t = "NITEM" Or t = "PARTIDAN°")
        Case "DESC"
            EncabezadoPres = (Left(t, 7) = "DESCRIP" Or t = "PARTIDA" Or t = "PARTIDAS" Or t = "NOMBREDELAPARTIDA")
        Case "UND"
            EncabezadoPres = (t = "UND" Or t = "UNID" Or t = "UNIDAD" Or t = "UM" Or t = "UNIDADDEMEDIDA" Or t = "U")
        Case "MET"
            EncabezadoPres = (Left(t, 7) = "METRADO" Or Left(t, 4) = "CANT")
        Case "PU"
            EncabezadoPres = (Left(t, 6) = "PRECIO" Or t = "PU" Or t = "CU" Or Left(t, 4) = "PUNI" Or Left(t, 9) = "COSTOUNIT" Or _
                              t = "PUNITARIO" Or t = "PRECIOUNITARIO" Or t = "COSTO")
        Case "PAR"
            EncabezadoPres = (Left(t, 7) = "PARCIAL")
    End Select
End Function

Private Function LeerHojaPresupuesto(ByVal ws As Worksheet) As Long
    Dim v As Variant, nr As Long, nc As Long, r As Long, rr As Long, c As Long, h As String, hdr As Long
    Dim cI As Long, cD As Long, cU As Long, cM As Long, cP As Long, cA As Long, nOk As Long
    Dim cod As String, ds As String, un As String, num0 As Boolean, k As String, intento As Long
    Dim colE As Collection, r0 As Long, c0 As Long, rg As Range
    eN = 0
    eDupl = 0
    LeerHojaPresupuesto = 0
    On Error Resume Next
    Set rg = ws.UsedRange
    If rg Is Nothing Then Exit Function
    v = rg.Value
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    If Not IsArray(v) Then Exit Function
    r0 = rg.Row
    c0 = rg.Column
    nr = UBound(v, 1)
    nc = UBound(v, 2)
    If nc > 40 Then nc = 40
    ' 1) fila de títulos de columnas (pueden estar repartidos en dos filas)
    hdr = 0
    For r = 1 To IIf(nr > 120, 120, nr)
        cI = 0: cD = 0: cU = 0: cM = 0: cP = 0: cA = 0
        For rr = r To IIf(r + 1 > nr, nr, r + 1)
            For c = 1 To nc
                If VarType(v(rr, c)) = vbString Then
                    h = Normal(v(rr, c))
                    If h <> "" Then
                        If cI = 0 And EncabezadoPres(h, "ITEM") Then
                            cI = c
                        ElseIf cD = 0 And EncabezadoPres(h, "DESC") And c <> cI Then
                            cD = c
                        ElseIf cU = 0 And EncabezadoPres(h, "UND") Then
                            cU = c
                        ElseIf cM = 0 And EncabezadoPres(h, "MET") Then
                            cM = c
                        ElseIf cP = 0 And EncabezadoPres(h, "PU") Then
                            cP = c
                        ElseIf cA = 0 And EncabezadoPres(h, "PAR") Then
                            cA = c
                        End If
                    End If
                End If
            Next c
        Next rr
        nOk = 0
        If cU > 0 Then nOk = nOk + 1
        If cM > 0 Then nOk = nOk + 1
        If cP > 0 Then nOk = nOk + 1
        If cA > 0 Then nOk = nOk + 1
        If cI > 0 And cD > 0 And nOk >= 3 Then
            hdr = r
            If cU > 0 And cM > 0 Then
                ' si los títulos ocupan dos filas, los datos empiezan después de la segunda
                For c = 1 To nc
                    If VarType(v(r, c)) = vbString Then
                        If c = cU Or c = cM Or c = cP Or c = cA Then Exit For
                    End If
                Next c
                If c > nc And r < nr Then hdr = r + 1
            End If
            Exit For
        End If
    Next r
    If hdr = 0 Then
        ' sin títulos: se asume ÍTEM, DESCRIPCIÓN, UND, METRADO, P.U., PARCIAL desde la primera columna con ítems
        For c = 1 To IIf(nc > 5, 5, nc)
            nOk = 0
            For r = 1 To IIf(nr > 200, 200, nr)
                If EsCodigo(CodigoCelda(v(r, c))) Then nOk = nOk + 1
            Next r
            If nOk >= 3 Then Exit For
        Next c
        If c > IIf(nc > 5, 5, nc) Or c + 5 > nc Then Exit Function
        cI = c: cD = c + 1: cU = c + 2: cM = c + 3: cP = c + 4: cA = c + 5
        hdr = 0
    End If
    ' 2) filas con ítem
    ReDim eCod(1 To nr): ReDim eDesc(1 To nr): ReDim eUnd(1 To nr)
    ReDim eMet(1 To nr): ReDim ePU(1 To nr): ReDim ePar(1 To nr)
    Set colE = New Collection
    For r = hdr + 1 To nr
        cod = CodigoCelda(v(r, cI))
        num0 = (VarType(v(r, cI)) <> vbString)
        If num0 And cod <> "" Then
            ' ítem guardado como número: se usa el texto que muestra la celda (1.10 y no 1.1)
            On Error Resume Next
            h = Trim(Replace(ws.Cells(r0 + r - 1, c0 + cI - 1).Text, ",", "."))
            On Error GoTo 0
            If EsCodigo(h) Then cod = h
        End If
        If Right(cod, 2) = ".0" And InStr(cod, ".") = InStrRev(cod, ".") Then cod = Left(cod, Len(cod) - 2)
        If EsCodigo(cod) Then
            ds = ""
            If cD > 0 Then ds = Limpio(v(r, cD))
            If ds = "" Then
                For c = cI + 1 To nc
                    If c <> cU And VarType(v(r, c)) = vbString Then
                        If Limpio(v(r, c)) <> "" Then
                            ds = Limpio(v(r, c))
                            Exit For
                        End If
                    End If
                Next c
            End If
            un = ""
            If cU > 0 Then un = Limpio(v(r, cU))
            If ds <> "" Or un <> "" Then
                k = Clave(cod)
                If num0 And Buscar(colE, "E" & k) > 0 Then
                    ' 1.1 repetido después de 1.9: es 1.10
                    For intento = 1 To 2
                        cod = cod & "0"
                        k = Clave(cod)
                        If Buscar(colE, "E" & k) = 0 Then Exit For
                    Next intento
                End If
                If Buscar(colE, "E" & k) > 0 Then
                    eDupl = eDupl + 1
                Else
                    eN = eN + 1
                    colE.Add eN, "E" & k
                    eCod(eN) = cod
                    eDesc(eN) = ds
                    eUnd(eN) = un
                    If un <> "" Then
                        If cM > 0 Then eMet(eN) = Num(v(r, cM))
                        If cP > 0 Then ePU(eN) = Num(v(r, cP))
                        If cA > 0 Then ePar(eN) = Num(v(r, cA))
                        If Num(ePar(eN)) = 0 Then ePar(eN) = Redondear(Num(eMet(eN)) * Num(ePU(eN)), 2)
                    ElseIf cA > 0 Then
                        If EsNum(v(r, cA)) Then ePar(eN) = Num(v(r, cA))
                    End If
                End If
            End If
        End If
    Next r
    nOk = 0
    For r = 1 To eN
        If eUnd(r) <> "" Then nOk = nOk + 1
    Next r
    If nOk = 0 Then
        eN = 0
        Exit Function
    End If
    eHoja = ws.Name
    LeerHojaPresupuesto = eN
End Function

Private Function CodigoCelda(ByVal x As Variant) As String
    Dim s As String
    CodigoCelda = ""
    If IsEmpty(x) Or IsError(x) Then Exit Function
    If VarType(x) = vbString Then
        s = Limpio(x)
    ElseIf IsNumeric(x) Then
        s = Replace(Trim(CStr(x)), ",", ".")
    Else
        Exit Function
    End If
    If Right(s, 1) = "." Then s = Left(s, Len(s) - 1)
    CodigoCelda = s
End Function


'========================================================================== A.C.U.
Public Function ImportarACUDe(ByVal ruta As String) As String
    Dim calcMode As Long, wb As Workbook, ws As Worksheet, ya As Boolean, msg As String
    Dim colPres As Collection, colDesc As Collection, pres As Worksheet, v As Variant, i As Long, j As Long, k As Long
    Dim nPres As Long, codPres() As String, puPres() As Double, usado() As Boolean, enAcu() As Boolean
    Dim destino() As String, nCod As Long, nDes As Long, nFuera As Long, sFuera As String
    Dim nSpOk As Long, nSpNo As Long, sSpNo As String, nPU As Long, sPU As String, suma As Double
    Dim o() As Variant, nf As Long, nSinAcu As Long, sSinAcu As String, rendG() As Variant, cnt(1 To 6) As Long
    Dim hayPres As Boolean, kk As Long
    Inicio calcMode, "Leyendo el A.C.U...."
    On Error GoTo Fallo
    If ruta = "" Then Err.Raise 3001, , "No se indicó el archivo."
    If Dir(ruta) = "" Then Err.Raise 3002, , "No se encontró el archivo: " & ruta
    bN = 0
    tN = 0
    sHojas = ""
    ReDim bCod(1 To 64): ReDim bDesc(1 To 64): ReDim bSub(1 To 64): ReDim bRend(1 To 64)
    ReDim bIni(1 To 64): ReDim bFin(1 To 64): ReDim bDescN(1 To 64): ReDim bUnd(1 To 64)
    ReDim tIns(1 To 512): ReDim tUnd(1 To 512): ReDim tTipo(1 To 512): ReDim tCant(1 To 512)
    ReDim tPre(1 To 512): ReDim tCodI(1 To 512): ReDim tPend(1 To 512)
    Set wb = AbrirLibro(ruta, ya)
    For Each ws In wb.Worksheets
        k = bN
        LeerHojaAcu ws
        If bN > k Then sHojas = Agregar(sHojas, ws.Name)
    Next ws
    If Not ya Then wb.Close False
    Set wb = Nothing
    If bN = 0 Then Err.Raise 3004, , "No se encontraron análisis de costos unitarios en el archivo (se buscan encabezados " & _
                                   "«Partida 01.01.01» o «01.01.01 DESCRIPCIÓN» y columnas INSUMO / UNIDAD / CANTIDAD / PRECIO)."
    ExpandirSubpartidas nSpOk, nSpNo, sSpNo
    ' ---- partidas del presupuesto
    Set pres = ThisWorkbook.Worksheets("PRESUPUESTO")
    Set colPres = New Collection
    Set colDesc = New Collection
    v = pres.Range(pres.Cells(PP0, 1), pres.Cells(PPMAX, 7)).Value
    nPres = 0
    ReDim codPres(1 To UBound(v, 1)): ReDim puPres(1 To UBound(v, 1)): ReDim enAcu(1 To UBound(v, 1))
    ReDim rendG(1 To UBound(v, 1), 1 To 1)
    For i = 1 To UBound(v, 1)
        rendG(i, 1) = v(i, 7)
        If Txt(v(i, 1)) <> "" And Txt(v(i, 3)) <> "" Then
            nPres = nPres + 1
            codPres(i) = CodigoCelda(v(i, 1))
            puPres(i) = Num(v(i, 5))
            If Buscar(colPres, "P" & Clave(codPres(i))) = 0 Then colPres.Add i, "P" & Clave(codPres(i))
            If Buscar(colDesc, "D" & Normal(v(i, 2))) = 0 Then
                colDesc.Add i, "D" & Normal(v(i, 2))
            Else
                colDesc.Remove "D" & Normal(v(i, 2))
                colDesc.Add -1, "D" & Normal(v(i, 2))
            End If
        End If
    Next i
    hayPres = (nPres > 0)
    ' ---- emparejar cada análisis con su partida (por ítem y, si no, por descripción)
    ReDim destino(1 To bN): ReDim usado(1 To UBound(v, 1))
    For j = 1 To bN
        destino(j) = ""
        If Not bSub(j) Then
            k = 0
            If hayPres Then
                k = Buscar(colPres, "P" & Clave(bCod(j)))
                If k > 0 Then
                    If usado(k) Then k = 0
                End If
                If k > 0 Then
                    nCod = nCod + 1
                Else
                    k = Buscar(colDesc, "D" & bDescN(j))
                    If k > 0 Then
                        If usado(k) Then
                            k = 0
                        Else
                            nDes = nDes + 1
                        End If
                    Else
                        k = 0
                    End If
                End If
                If k > 0 Then
                    usado(k) = True
                    destino(j) = codPres(k)
                    If bRend(j) > 0 Then rendG(k, 1) = bRend(j)
                ElseIf bFin(j) >= bIni(j) Then
                    nFuera = nFuera + 1
                    sFuera = Agregar(sFuera, bCod(j))
                End If
            Else
                destino(j) = bCod(j)
            End If
        End If
    Next j
    ' ---- filas para la hoja ACU (insumos repetidos de una partida se suman)
    nf = 0
    For j = 1 To bN
        If destino(j) <> "" Then nf = nf + (bFin(j) - bIni(j) + 1)
    Next j
    If nf = 0 Then Err.Raise 3005, , "Ningún análisis del archivo corresponde a las partidas del PRESUPUESTO (" & bN & _
                                   " análisis leídos). Revise que el archivo sea de la misma obra."
    If nf > ACMAX - AC0 + 1 Then Err.Raise 3006, , "El A.C.U. tiene " & nf & " filas; el programa admite hasta " & (ACMAX - AC0 + 1) & "."
    ReDim o(1 To nf, 1 To 6)
    nf = 0
    For j = 1 To bN
        If destino(j) <> "" Then
            k = nf
            For i = bIni(j) To bFin(j)
                kk = 0
                For kk = k + 1 To nf
                    If UCase(o(kk, 2)) = UCase(tIns(i)) And o(kk, 4) = tTipo(i) And UCase(o(kk, 3)) = UCase(tUnd(i)) And _
                       o(kk, 6) = tPre(i) Then Exit For
                Next kk
                If kk <= nf And nf > k Then
                    o(kk, 5) = o(kk, 5) + tCant(i)
                Else
                    nf = nf + 1
                    o(nf, 1) = destino(j)
                    o(nf, 2) = tIns(i)
                    o(nf, 3) = tUnd(i)
                    o(nf, 4) = tTipo(i)
                    o(nf, 5) = tCant(i)
                    o(nf, 6) = tPre(i)
                End If
                Select Case tTipo(i)
                    Case "MO": cnt(1) = cnt(1) + 1
                    Case "MAT": cnt(2) = cnt(2) + 1
                    Case "EQ": cnt(3) = cnt(3) + 1
                    Case "SC": cnt(4) = cnt(4) + 1
                    Case "%MO": cnt(5) = cnt(5) + 1
                    Case Else: cnt(6) = cnt(6) + 1
                End Select
            Next i
            ' revisión del P.U.
            If hayPres Then
                k = Buscar(colPres, "P" & Clave(destino(j)))
                If k > 0 Then
                    suma = 0
                    For i = bIni(j) To bFin(j)
                        If tTipo(i) = "%MO" Then
                            suma = suma + tCant(i) * tPre(i) / 100
                        Else
                            suma = suma + tCant(i) * tPre(i)
                        End If
                    Next i
                    If Abs(suma - puPres(k)) > IIf(puPres(k) * 0.01 > 0.1, puPres(k) * 0.01, 0.1) Then
                        nPU = nPU + 1
                        sPU = Agregar(sPU, destino(j))
                    End If
                End If
            End If
        End If
    Next j
    ' ---- partidas del presupuesto sin análisis
    If hayPres Then
        For i = 1 To UBound(v, 1)
            If codPres(i) <> "" Then
                If Not usado(i) Then
                    nSinAcu = nSinAcu + 1
                    sSinAcu = Agregar(sSinAcu, codPres(i))
                End If
            End If
        Next i
    End If
    ' ---- escritura
    With ThisWorkbook.Worksheets("ACU")
        .Range(.Cells(AC0, 1), .Cells(ACMAX, 16)).ClearContents
        .Range(.Cells(AC0, 1), .Cells(ACMAX, 4)).NumberFormat = "@"
        .Range(.Cells(AC0, 1), .Cells(AC0 + nf - 1, 6)).Value = o
    End With
    If hayPres Then pres.Range(pres.Cells(PP0, 7), pres.Cells(PPMAX, 7)).Value = rendG
    msg = "A.C.U. importado de «" & NombreArchivo(ruta) & "» (hoja(s): " & Cortar(sHojas, 80) & ")." & vbCrLf & vbCrLf & _
          "• " & bN & " análisis leídos; " & nf & " filas cargadas en la hoja ACU." & vbCrLf & _
          "• Filas por tipo: MO " & cnt(1) & ", MAT " & cnt(2) & ", EQ " & cnt(3) & ", SC " & cnt(4) & ", %MO " & cnt(5) & _
          IIf(cnt(6) > 0, ", SP sin análisis " & cnt(6), "") & "."
    If hayPres Then
        msg = msg & vbCrLf & "• Partidas emparejadas: " & (nCod + nDes) & " de " & nPres & " (" & nCod & " por ítem" & _
              IIf(nDes > 0, ", " & nDes & " por descripción", "") & ")."
    Else
        msg = msg & vbCrLf & "• No hay PRESUPUESTO: las partidas se cargaron con los ítems del archivo."
    End If
    If nSpOk > 0 Then msg = msg & vbCrLf & "• " & nSpOk & " subpartida(s) reemplazada(s) por sus insumos."
    If nSpNo > 0 Or nSinAcu > 0 Or nFuera > 0 Or nPU > 0 Then msg = msg & vbCrLf & vbCrLf & "Revise:"
    If nSpNo > 0 Then msg = msg & vbCrLf & "• " & nSpNo & " subpartida(s) sin análisis en el archivo (quedan como SP, no se controlan): " & _
                            Cortar(sSpNo, 120)
    If nSinAcu > 0 Then msg = msg & vbCrLf & "• " & nSinAcu & " partida(s) del presupuesto sin análisis: " & Cortar(sSinAcu, 150)
    If nFuera > 0 Then msg = msg & vbCrLf & "• " & nFuera & " análisis de partidas que no están en el presupuesto (no se cargaron): " & _
                             Cortar(sFuera, 120)
    If nPU > 0 Then msg = msg & vbCrLf & "• " & nPU & " partida(s) cuyo análisis no suma el P.U. del presupuesto: " & Cortar(sPU, 120)
    ImportarACUDe = msg & vbCrLf & vbCrLf & "Siguiente paso: PREPARAR OBRA."
Salir:
    Fin calcMode
    Exit Function
Fallo:
    ImportarACUDe = "No se pudo importar el A.C.U." & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then
        If Not ya Then wb.Close False
    End If
    Resume Salir
End Function

Private Sub LeerHojaAcu(ByVal ws As Worksheet)
    Dim v As Variant, nr As Long, nc As Long, r As Long, c As Long, rg As Range, h As String, s As String
    Dim cod As String, ds As String, esSub As Boolean, lab As String, post As Boolean, i As Long, cur As Long
    Dim ins As String, un As String, curTipo As String, pend0 As Long, cant As Variant, x As Double, p As Long
    Dim pr As Double, ultFila As Long, modoP As Boolean
    On Error Resume Next
    Set rg = ws.UsedRange
    If rg Is Nothing Then Exit Sub
    v = rg.Value
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    If Not IsArray(v) Then Exit Sub
    nr = UBound(v, 1)
    nc = UBound(v, 2)
    If nc > 70 Then nc = 70
    cIns = 0: cUnd = 0: cCua = 0: cCan = 0: cPre = 0: cPar = 0: cCodI = 0
    cur = 0
    curTipo = ""
    pend0 = tN + 1
    ultFila = -9
    modoP = (ContarPartidas(v, nr, nc) >= 3)
    For r = 1 To nr
        ' ---- títulos de columnas (se repiten en cada partida)
        If EncabezadoAcu(v, r, nc) Then GoTo Siguiente
        ' ---- encabezado de partida
        If CabeceraPartida(v, r, nc, cod, ds, esSub, modoP) Then
            CerrarBloque cur
            bN = bN + 1
            If bN > UBound(bCod) Then
                ReDim Preserve bCod(1 To 2 * bN): ReDim Preserve bDesc(1 To 2 * bN): ReDim Preserve bSub(1 To 2 * bN)
                ReDim Preserve bRend(1 To 2 * bN): ReDim Preserve bIni(1 To 2 * bN): ReDim Preserve bFin(1 To 2 * bN)
                ReDim Preserve bDescN(1 To 2 * bN): ReDim Preserve bUnd(1 To 2 * bN)
            End If
            cur = bN
            bCod(cur) = cod
            bDesc(cur) = ds
            bDescN(cur) = Normal(ds)
            bSub(cur) = esSub
            bRend(cur) = RendimientoFila(v, r, nc, False)
            bIni(cur) = tN + 1
            bFin(cur) = tN
            curTipo = ""
            pend0 = tN + 1
            GoTo Siguiente
        End If
        If cur = 0 Then GoTo Siguiente
        ' ---- rendimiento (fila que empieza con "Rendimiento" / "Rend.")
        x = RendimientoFila(v, r, nc, True)
        If x > 0 Then
            If bRend(cur) = 0 Then bRend(cur) = x
            GoTo Siguiente
        End If
        ' ---- insumo (requiere los títulos de columnas)
        If cIns > 0 Then
            s = Txt(v(r, cIns))
            un = ""
            If cUnd > 0 Then un = Limpio(v(r, cUnd))
            ' nombre corrido a otra columna (archivos convertidos desde PDF)
            If s = "" And un <> "" Then
                For c = 1 To cUnd - 1
                    If VarType(v(r, c)) = vbString Then
                        If Limpio(v(r, c)) <> "" And Not SoloCodigo(Limpio(v(r, c))) Then
                            s = Txt(v(r, c))
                            Exit For
                        End If
                    End If
                Next c
            End If
            ' etiqueta de sección pegada al nombre ("Materiales" + salto de línea + insumo)
            If InStr(s, vbLf) > 0 Then
                h = Left(s, InStr(s, vbLf) - 1)
                lab = Etiqueta(h)
                If lab <> "" Then
                    curTipo = lab
                    pend0 = tN + 1
                    s = Mid(s, InStr(s, vbLf) + 1)
                ElseIf Trim(Replace(h, vbCr, "")) = "" Then
                    s = Mid(s, InStr(s, vbLf) + 1)
                End If
            End If
            ins = QuitarCodigo(Limpio(s))
            If ins <> "" And un <> "" And Etiqueta(ins) = "" Then
                If CantidadPrecio(v, r, nc, un, x, pr) Then
                    tN = tN + 1
                    If tN > UBound(tIns) Then
                        ReDim Preserve tIns(1 To 2 * tN): ReDim Preserve tUnd(1 To 2 * tN): ReDim Preserve tTipo(1 To 2 * tN)
                        ReDim Preserve tCant(1 To 2 * tN): ReDim Preserve tPre(1 To 2 * tN): ReDim Preserve tCodI(1 To 2 * tN)
                        ReDim Preserve tPend(1 To 2 * tN)
                    End If
                    tIns(tN) = ins
                    tUnd(tN) = un
                    tCant(tN) = x
                    tPre(tN) = pr
                    tCodI(tN) = ""
                    If cCodI > 0 And cCodI <> cIns Then tCodI(tN) = Limpio(v(r, cCodI))
                    tPend(tN) = False
                    If TipoPorUnidad(un) = "%MO" Then
                        tTipo(tN) = "%MO"
                    ElseIf TipoPorUnidad(un) <> "MAT" And curTipo <> "MO" And curTipo <> "EQ" Then
                        ' hh = mano de obra; hm / he = equipo (aunque falte o se haya perdido la etiqueta de la sección)
                        tTipo(tN) = TipoPorUnidad(un)
                    ElseIf curTipo <> "" Then
                        tTipo(tN) = curTipo
                    Else
                        tTipo(tN) = ""
                        tPend(tN) = True
                    End If
                    bFin(cur) = tN
                    ultFila = r
                    GoTo Siguiente
                End If
            ElseIf ins <> "" And un = "" And ultFila = r - 1 And tN >= bIni(cur) And Etiqueta(ins) = "" Then
                ' continuación del nombre del insumo en la fila siguiente
                If Not HayNumero(v, r, nc) Then
                    tIns(tN) = tIns(tN) & " " & ins
                    ultFila = r
                    GoTo Siguiente
                End If
            End If
        End If
        ' ---- etiqueta de sección ("Mano de obra" antes de sus insumos, o "Mano de obra:" con su subtotal después)
        lab = ""
        post = False
        For c = 1 To nc
            If VarType(v(r, c)) = vbString Then
                h = Trim(v(r, c))
                If h <> "" Then
                    lab = Etiqueta(h)
                    If lab <> "" Then
                        If Right(h, 1) = ":" Then
                            For p = c + 1 To nc
                                If EsNum(v(r, p)) Then
                                    post = True
                                    Exit For
                                End If
                            Next p
                        End If
                        Exit For
                    End If
                End If
            End If
        Next c
        If lab <> "" Then
            If post Then
                For i = pend0 To tN
                    If tPend(i) Then
                        tTipo(i) = lab
                        tPend(i) = False
                    End If
                Next i
                curTipo = ""
            Else
                curTipo = lab
            End If
            pend0 = tN + 1
        End If
Siguiente:
    Next r
    CerrarBloque cur
End Sub

Private Function SoloCodigo(ByVal s As String) As Boolean
    ' "0147010003", "47 00005", "SP 36276": código de recurso sin nombre
    Dim t As String
    t = UCase(Replace(s, " ", ""))
    If Left(t, 2) = "SP" Then t = Mid(t, 3)
    SoloCodigo = EsDigitos(t)
End Function

Private Function HayNumero(ByRef v As Variant, ByVal r As Long, ByVal nc As Long) As Boolean
    Dim c As Long
    HayNumero = False
    For c = 1 To nc
        If EsNum(v(r, c)) Then
            HayNumero = True
            Exit Function
        End If
    Next c
End Function

Private Function Cuadra(ByVal c As Double, ByVal p As Double, ByVal par As Double, ByVal pct As Boolean) As Boolean
    Dim x As Double, tol As Double
    If pct Then
        x = c * p / 100
    Else
        x = c * p
    End If
    tol = Abs(par) * 0.002
    If tol < 0.0101 Then tol = 0.0101
    Cuadra = (Abs(x - par) <= tol)
End Function

Private Function CantidadPrecio(ByRef v As Variant, ByVal r As Long, ByVal nc As Long, ByVal un As String, _
                                ByRef cant As Double, ByRef pre As Double) As Boolean
    ' cantidad y precio del insumo: los de las columnas del encabezado; si no cuadran con el parcial (archivos
    ' convertidos desde PDF con columnas corridas), el par de números de la fila cuyo producto da el parcial
    Dim nv(1 To 12) As Double, nn As Long, c As Long, i As Long, j As Long, hayC As Boolean, hayP As Boolean
    Dim pct As Boolean, ch As Double, ph As Double
    CantidadPrecio = False
    cant = 0
    pre = 0
    pct = (TipoPorUnidad(un) = "%MO")
    nn = 0
    For c = cUnd + 1 To nc
        If EsNum(v(r, c)) Then
            If nn < 12 Then
                nn = nn + 1
                nv(nn) = Num(v(r, c))
            End If
        End If
    Next c
    If cCan > 0 Then
        If EsNum(v(r, cCan)) Then
            hayC = True
            ch = Num(v(r, cCan))
        End If
    End If
    If cPre > 0 Then
        If EsNum(v(r, cPre)) Then
            hayP = True
            ph = Num(v(r, cPre))
        End If
    End If
    If nn >= 3 Then
        If hayC And hayP Then
            If Cuadra(ch, ph, nv(nn), pct) Then
                cant = ch
                pre = ph
                CantidadPrecio = True
                Exit Function
            End If
        End If
        For j = nn - 1 To 2 Step -1
            For i = j - 1 To 1 Step -1
                If Cuadra(nv(i), nv(j), nv(nn), pct) Then
                    cant = nv(i)
                    pre = nv(j)
                    CantidadPrecio = True
                    Exit Function
                End If
            Next i
        Next j
    End If
    If hayC Then
        cant = ch
        pre = ph
        CantidadPrecio = True
    ElseIf nn >= 3 Then
        cant = nv(nn - 2)
        pre = nv(nn - 1)
        CantidadPrecio = True
    ElseIf nn = 2 And cCan = 0 Then
        cant = nv(1)
        pre = nv(2)
        CantidadPrecio = True
    End If
End Function

Private Sub CerrarBloque(ByVal cur As Long)
    Dim i As Long
    If cur = 0 Then Exit Sub
    For i = bIni(cur) To bFin(cur)
        If tPend(i) Then
            tTipo(i) = TipoPorUnidad(tUnd(i))
            tPend(i) = False
        End If
    Next i
End Sub

Private Function Etiqueta(ByVal s As String) As String
    Dim t As String
    t = Normal(s)
    If Right(t, 1) = ":" Then t = Trim(Left(t, Len(t) - 1))
    Select Case t
        Case "MANO DE OBRA", "MANO OBRA", "M.O."
            Etiqueta = "MO"
        Case "MATERIALES", "MATERIAL"
            Etiqueta = "MAT"
        Case "EQUIPOS", "EQUIPO", "EQUIPOS Y HERRAMIENTAS", "HERRAMIENTAS Y EQUIPOS", "MAQUINARIA Y EQUIPOS", "EQUIPO Y HERRAMIENTAS"
            Etiqueta = "EQ"
        Case "SUBCONTRATOS", "SUBCONTRATO", "SUB CONTRATOS", "SERVICIOS"
            Etiqueta = "SC"
        Case "SUBPARTIDAS", "SUB PARTIDAS", "SUB PARTIDA", "SUBPARTIDA"
            Etiqueta = "SP"
        Case Else
            Etiqueta = ""
    End Select
End Function

Private Function EncabezadoAcu(ByRef v As Variant, ByVal r As Long, ByVal nc As Long) As Boolean
    Dim c As Long, h As String, t As String, i0 As Long, u0 As Long, k0 As Long
    Dim q0 As Long, p0 As Long, a0 As Long, d0 As Long
    EncabezadoAcu = False
    For c = 1 To nc
        If VarType(v(r, c)) = vbString Then
            h = Normal(v(r, c))
            If h <> "" And Len(h) < 60 Then
                t = Replace(Replace(h, ".", ""), " ", "")
                If i0 = 0 And (InStr(t, "INSUMO") > 0 Or InStr(t, "RECURSO") > 0 Or Left(t, 7) = "DESCRIP") Then
                    i0 = c
                    If Left(t, 3) = "COD" Then d0 = c
                ElseIf u0 = 0 And (t = "UNIDAD" Or t = "UND" Or t = "UNID" Or t = "UM") Then
                    u0 = c
                ElseIf q0 = 0 And Left(t, 4) = "CANT" Then
                    q0 = c
                ElseIf k0 = 0 And Left(t, 4) = "CUAD" Then
                    k0 = c
                ElseIf p0 = 0 And (Left(t, 6) = "PRECIO" Or t = "PU" Or t = "PUNIT" Or t = "COSTO" Or Left(t, 9) = "COSTOUNIT") Then
                    p0 = c
                ElseIf a0 = 0 And Left(t, 7) = "PARCIAL" Then
                    a0 = c
                ElseIf d0 = 0 And Left(t, 3) = "COD" Then
                    d0 = c
                End If
            End If
        End If
    Next c
    If i0 > 0 And u0 > 0 And q0 > 0 Then
        cIns = i0: cUnd = u0: cCan = q0: cCua = k0: cPre = p0: cPar = a0: cCodI = d0
        EncabezadoAcu = True
    End If
End Function

Private Function CabeceraPartida(ByRef v As Variant, ByVal r As Long, ByVal nc As Long, ByRef cod As String, _
                                 ByRef ds As String, ByRef esSub As Boolean, ByVal modoPartida As Boolean) As Boolean
    Dim c As Long, h As String, t As String, resto As String, cc As Long
    CabeceraPartida = False
    cod = ""
    ds = ""
    esSub = False
    ' una fila de insumo (con unidad y cantidad) no es un encabezado
    If cUnd > 0 And cCan > 0 Then
        If Txt(v(r, cUnd)) <> "" And EsNum(v(r, cCan)) Then Exit Function
    End If
    ' primera celda con dato
    For c = 1 To nc
        If Txt(v(r, c)) <> "" Then Exit For
    Next c
    If c > nc Then Exit Function
    If VarType(v(r, c)) <> vbString Then Exit Function
    h = Limpio(v(r, c))
    t = Normal(h)
    If Etiqueta(h) <> "" Then Exit Function
    If Left(t, 7) = "PARTIDA" Or Left(t, 11) = "SUB PARTIDA" Or Left(t, 10) = "SUBPARTIDA" Then
        ' "Partida 01.01.01 ..." / "Sub Partida 36276 ..." / "Partida" y el ítem en otra celda
        If Left(t, 8) = "PARTIDAS" Or Left(t, 12) = "SUB PARTIDAS" Or Left(t, 11) = "SUBPARTIDAS" Then Exit Function
        If Left(t, 7) = "PARTIDA" Then
            resto = Trim(Mid(h, 8))
        ElseIf Left(t, 11) = "SUB PARTIDA" Then
            resto = Trim(Mid(h, 12))
            esSub = True
        Else
            resto = Trim(Mid(h, 11))
            esSub = True
        End If
        If Left(resto, 1) = ":" Or Left(resto, 1) = "." Or Left(resto, 1) = "-" Then resto = Trim(Mid(resto, 2))
        If resto <> "" Then SepararCodigo resto, cod, ds
        If cod = "" Then
            ' el ítem está en la celda siguiente (como texto)
            For cc = c + 1 To nc
                If VarType(v(r, cc)) = vbString Then
                    If Limpio(v(r, cc)) <> "" Then
                        SepararCodigo Limpio(v(r, cc)), cod, ds
                        c = cc
                        Exit For
                    End If
                ElseIf Txt(v(r, cc)) <> "" Then
                    Exit For
                End If
            Next cc
        End If
        If cod = "" Then Exit Function
        If ds = "" Then
            For cc = c + 1 To nc
                If VarType(v(r, cc)) = vbString Then
                    If Limpio(v(r, cc)) <> "" And Not EsCodigo(Limpio(v(r, cc))) Then
                        If Left(Normal(v(r, cc)), 4) <> "REND" Then
                            ds = Limpio(v(r, cc))
                            Exit For
                        End If
                    End If
                End If
            Next cc
        End If
        CabeceraPartida = True
        Exit Function
    End If
    ' "1.1.1.1.1 LIMPIEZA ..." (ítem con puntos, un espacio y la descripción en la misma celda); solo en los
    ' archivos que no usan la palabra "Partida" y si la fila no tiene unidad
    If modoPartida Then Exit Function
    If cUnd > 0 Then
        If Txt(v(r, cUnd)) <> "" Then Exit Function
    End If
    SepararCodigo h, cod, ds
    If cod <> "" Then
        If Mid(h, Len(cod) + 1, 1) <> " " Then cod = ""
    End If
    If cod <> "" Then
        If Val(Split(cod, ".")(0)) = 0 Then cod = ""
    End If
    If cod <> "" And InStr(cod, ".") > 0 And Len(ds) >= 2 Then
        t = Normal(ds)
        If (Left(t, 1) >= "A" And Left(t, 1) <= "Z" And Mid(t, 2, 1) >= "A" And Mid(t, 2, 1) <= "Z") Or Left(ds, 1) = """" _
           Or Left(ds, 1) = "(" Then
            CabeceraPartida = True
            Exit Function
        End If
    End If
    cod = ""
    ds = ""
End Function

Private Function ContarPartidas(ByRef v As Variant, ByVal nr As Long, ByVal nc As Long) As Long
    ' filas cuya primera celda dice "Partida" (formato S10 y similares)
    Dim r As Long, c As Long, t As String, n As Long
    n = 0
    For r = 1 To nr
        For c = 1 To nc
            If Txt(v(r, c)) <> "" Then Exit For
        Next c
        If c <= nc Then
            If VarType(v(r, c)) = vbString Then
                t = Normal(v(r, c))
                If (Left(t, 8) = "PARTIDA " Or t = "PARTIDA") Then n = n + 1
            End If
        End If
        If n >= 3 Then Exit For
    Next r
    ContarPartidas = n
End Function

Private Sub SepararCodigo(ByVal s As String, ByRef cod As String, ByRef ds As String)
    ' "01.01.02   DESCRIPCIÓN" -> cod = "01.01.02", ds = "DESCRIPCIÓN"
    Dim i As Long, ch As String
    cod = ""
    ds = ""
    s = Trim(s)
    If s = "" Then Exit Sub
    If Left(s, 1) < "0" Or Left(s, 1) > "9" Then Exit Sub
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or ch = ".") Then Exit Do
        i = i + 1
    Loop
    cod = Left(s, i - 1)
    If Right(cod, 1) = "." Then cod = Left(cod, Len(cod) - 1)
    ds = Trim(Mid(s, i))
    If ds <> "" Then
        If Left(ds, 1) = "-" Or Left(ds, 1) = ":" Then ds = Trim(Mid(ds, 2))
    End If
    If Not EsCodigo(cod) Then
        cod = ""
        ds = ""
    End If
End Sub

Private Function RendimientoFila(ByRef v As Variant, ByVal r As Long, ByVal nc As Long, ByVal soloPrimera As Boolean) As Double
    ' "Rendimiento: 300.0000 M2/DIA" / "Rend: 300" / "Rendimiento | und/DIA | 1.0000"
    Dim c As Long, t As String, p As Long, x As Double
    RendimientoFila = 0
    For c = 1 To nc
        If Txt(v(r, c)) <> "" Then
            If VarType(v(r, c)) = vbString Then
                t = Normal(v(r, c))
                If Left(t, 4) = "REND" Then
                    x = PrimerNumero(Mid(t, 5))
                    If x <= 0 Then
                        For p = c + 1 To nc
                            If EsNum(v(r, p)) Then
                                x = Num(v(r, p))
                                Exit For
                            End If
                        Next p
                    End If
                    RendimientoFila = x
                    Exit Function
                End If
            End If
            If soloPrimera Then Exit Function
        End If
    Next c
End Function

Private Function PrimerNumero(ByVal s As String) As Double
    Dim i As Long, j As Long, ch As String, t As String, ok As Boolean, x As Double
    PrimerNumero = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            j = i
            Do While j <= Len(s)
                ch = Mid(s, j, 1)
                If Not ((ch >= "0" And ch <= "9") Or ch = "." Or ch = ",") Then Exit Do
                j = j + 1
            Loop
            t = Mid(s, i, j - i)
            x = ParseNum(t, ok)
            If ok Then PrimerNumero = x
            Exit Function
        End If
    Next i
End Function

Private Function QuitarCodigo(ByVal s As String) As String
    ' "47 00005 OPERARIO" o "0147010003 Operario" -> "OPERARIO" (código de recurso al inicio del nombre)
    Dim i As Long, ch As String, dig As Long
    QuitarCodigo = s
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            dig = dig + 1
        ElseIf ch <> " " Then
            Exit Do
        End If
        i = i + 1
    Loop
    If dig >= 6 And i > 1 And i <= Len(s) Then
        If Mid(s, i - 1, 1) = " " Then QuitarCodigo = Trim(Mid(s, i))
    End If
End Function

Private Sub ExpandirSubpartidas(ByRef nOk As Long, ByRef nNo As Long, ByRef sNo As String)
    ' reemplaza cada fila SP por los insumos de su análisis (multiplicados por la cantidad de la subpartida)
    Dim pasada As Long, j As Long, i As Long, k As Long, m As Long, hubo As Boolean, f As Double
    Dim nIns() As String, nUnd() As String, nTip() As String, nCan() As Double, nPre() As Double, nCod() As String
    Dim nn As Long, nIni() As Long, nFin() As Long, codSp As String
    nOk = 0
    nNo = 0
    sNo = ""
    For pasada = 1 To 4
        hubo = False
        nn = 0
        ReDim nIns(1 To tN * 2 + 16): ReDim nUnd(1 To tN * 2 + 16): ReDim nTip(1 To tN * 2 + 16)
        ReDim nCan(1 To tN * 2 + 16): ReDim nPre(1 To tN * 2 + 16): ReDim nCod(1 To tN * 2 + 16)
        ReDim nIni(1 To bN): ReDim nFin(1 To bN)
        For j = 1 To bN
            nIni(j) = nn + 1
            For i = bIni(j) To bFin(j)
                k = 0
                If tTipo(i) = "SP" Then k = BuscarSubpartida(i, j)
                If k > 0 And pasada < 4 Then
                    f = tCant(i)
                    For m = bIni(k) To bFin(k)
                        nn = nn + 1
                        If nn > UBound(nIns) Then
                            ReDim Preserve nIns(1 To 2 * nn): ReDim Preserve nUnd(1 To 2 * nn): ReDim Preserve nTip(1 To 2 * nn)
                            ReDim Preserve nCan(1 To 2 * nn): ReDim Preserve nPre(1 To 2 * nn): ReDim Preserve nCod(1 To 2 * nn)
                        End If
                        nIns(nn) = tIns(m)
                        nUnd(nn) = tUnd(m)
                        nTip(nn) = tTipo(m)
                        If tTipo(m) = "%MO" Then
                            nCan(nn) = tCant(m)
                            nPre(nn) = tPre(m) * f
                        Else
                            nCan(nn) = tCant(m) * f
                            nPre(nn) = tPre(m)
                        End If
                        nCod(nn) = tCodI(m)
                    Next m
                    If pasada = 1 Then nOk = nOk + 1
                    hubo = True
                Else
                    nn = nn + 1
                    If nn > UBound(nIns) Then
                        ReDim Preserve nIns(1 To 2 * nn): ReDim Preserve nUnd(1 To 2 * nn): ReDim Preserve nTip(1 To 2 * nn)
                        ReDim Preserve nCan(1 To 2 * nn): ReDim Preserve nPre(1 To 2 * nn): ReDim Preserve nCod(1 To 2 * nn)
                    End If
                    nIns(nn) = tIns(i)
                    nUnd(nn) = tUnd(i)
                    nTip(nn) = tTipo(i)
                    nCan(nn) = tCant(i)
                    nPre(nn) = tPre(i)
                    nCod(nn) = tCodI(i)
                End If
            Next i
            nFin(j) = nn
        Next j
        tN = nn
        ReDim tIns(1 To nn + 1): ReDim tUnd(1 To nn + 1): ReDim tTipo(1 To nn + 1): ReDim tCant(1 To nn + 1)
        ReDim tPre(1 To nn + 1): ReDim tCodI(1 To nn + 1): ReDim tPend(1 To nn + 1)
        For i = 1 To nn
            tIns(i) = nIns(i): tUnd(i) = nUnd(i): tTipo(i) = nTip(i): tCant(i) = nCan(i): tPre(i) = nPre(i): tCodI(i) = nCod(i)
        Next i
        For j = 1 To bN
            bIni(j) = nIni(j)
            bFin(j) = nFin(j)
        Next j
        If Not hubo Then Exit For
    Next pasada
    For j = 1 To bN
        If Not bSub(j) Then
            For i = bIni(j) To bFin(j)
                If tTipo(i) = "SP" Then
                    nNo = nNo + 1
                    sNo = Agregar(sNo, tIns(i))
                End If
            Next i
        End If
    Next j
End Sub

Private Function BuscarSubpartida(ByVal i As Long, ByVal j As Long) As Long
    ' análisis de la subpartida de la fila i (bloque j): primero por código ("SP 36276" / "Sub Partida 36276"),
    ' luego por descripción (se prefieren los bloques de subpartidas)
    Dim k As Long, codSp As String, dn As String, mejor As Long
    BuscarSubpartida = 0
    codSp = Replace(UCase(tCodI(i)), "SP", "")
    codSp = Replace(codSp, " ", "")
    If codSp <> "" Then
        For k = 1 To bN
            If k <> j And bSub(k) Then
                If Replace(bCod(k), " ", "") = codSp And bFin(k) >= bIni(k) Then
                    BuscarSubpartida = k
                    Exit Function
                End If
            End If
        Next k
    End If
    dn = Normal(tIns(i))
    mejor = 0
    For k = 1 To bN
        If k <> j And bDescN(k) = dn And bFin(k) >= bIni(k) Then
            If bSub(k) Then
                BuscarSubpartida = k
                Exit Function
            End If
            If mejor = 0 Then mejor = k
        End If
    Next k
    BuscarSubpartida = mejor
End Function
