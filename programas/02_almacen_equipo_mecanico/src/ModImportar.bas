Attribute VB_Name = "ModImportar"
Option Explicit

'==========================================================================================
'  IMPORTAR DEL CONTROL DE METRADOS DIARIOS
'  Ingresos registrados, salidas según A.C.U. x metrado del día y partes de los equipos.
'==========================================================================================

' ---- opciones (hoja DATOS)
Public Const O_CAT As String = "C33"
Public Const O_ING As String = "C34"
Public Const O_SAL As String = "C35"
Public Const O_PAR As String = "C36"
Public Const O_EXC As String = "C37"
' ---- hojas del CONTROL DE METRADOS
Private Const X_RR0 As Long = 10
Private Const X_RRMAX As Long = 5009
Private Const X_D1 As Long = 9
Private Const X_AC0 As Long = 8
Private Const X_ACMAX As Long = 40007
Private Const X_IN0 As Long = 8
Private Const X_INMAX As Long = 3007
Private Const X_IG0 As Long = 8
Private Const X_IGMAX As Long = 1007
Private Const KW_COMB As String = "DIESEL;PETROLEO;GASOLINA;GASOHOL;KEROSENE;COMBUSTIBLE"

Private mapCat As Collection                  ' nombre normalizado -> índice del catálogo (-1 = ambiguo)
Private cFac() As Double
Private mapCtl As Collection                  ' insumo del control -> 1 controlar / 2 no controlar
Private mapEq As Collection                   ' nombre del A.C.U. -> índice del equipo
Private nCat0 As Long, nNuevos As Long
Private qCod() As String, qComb() As Long, nEqI As Long, nEq0 As Long
Private nuevosEq As Long
Private sinCat As String, nSinCat As Long
Private excl() As String
Private mIni As Double, mFin As Double


Public Sub ImportarControl()
    Dim r As Variant, ruta As String, msg As String
    ruta = txt(Hoja("DATOS").Range(D_CTRL).Value)
    If ruta <> "" Then
        On Error Resume Next
        If Dir(ruta) = "" Then ruta = ""
        On Error GoTo 0
    End If
    If ruta = "" Then
        If gSilencio Then Exit Sub
        r = Application.GetOpenFilename("Libros de Excel (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , _
                                        "Elija el archivo del CONTROL DE METRADOS DIARIOS")
        If VarType(r) = vbBoolean Then Exit Sub
        ruta = CStr(r)
    End If
    If Not Pregunta("Se leerá el CONTROL DE METRADOS:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
                    "Según las opciones de la hoja DATOS se agregan al catálogo los materiales nuevos, se importan los " & _
                    "ingresos, se calculan las salidas (A.C.U. x metrado del día) y los partes de los equipos." & vbCrLf & _
                    "Lo importado antes se reemplaza; lo digitado a mano se conserva." & vbCrLf & vbCrLf & "¿Continuar?", _
                    "IMPORTAR DEL CONTROL DE METRADOS", True) Then Exit Sub
    msg = ImportarControlDe(ruta)
    Aviso msg, "IMPORTAR DEL CONTROL DE METRADOS"
End Sub


Private Function Opcion(ByVal celda As String) As Boolean
    Opcion = (UCase(Left(txt(Hoja("DATOS").Range(celda).Value), 1)) <> "N")
End Function

Public Function TieneHoja(ByVal wb As Workbook, ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    TieneHoja = False
    For Each ws In wb.Worksheets
        If UCase(ws.Name) = UCase(nombre) Then
            TieneHoja = True
            Exit Function
        End If
    Next ws
End Function

Public Function AbrirLibroX(ByVal ruta As String, ByRef yaAbierto As Boolean, ByVal soloLectura As Boolean) As Workbook
    Dim w As Workbook, sec As Long, nom As String
    yaAbierto = False
    nom = Mid(Replace(ruta, "/", "\"), InStrRev(Replace(ruta, "/", "\"), "\") + 1)
    For Each w In Workbooks
        If LCase(w.FullName) = LCase(ruta) Or LCase(w.Name) = LCase(nom) Then
            Set AbrirLibroX = w
            yaAbierto = True
            Exit Function
        End If
    Next w
    On Error Resume Next
    sec = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Application.DisplayAlerts = False
    On Error GoTo 0
    Set AbrirLibroX = Workbooks.Open(ruta, 0, soloLectura)
    On Error Resume Next
    If sec <> 0 Then Application.AutomationSecurity = sec
    Application.DisplayAlerts = True
    On Error GoTo 0
End Function

Private Function EnMes(ByVal t As Double) As Boolean
    EnMes = (t >= mIni And t <= mFin And t > 0)
End Function

Private Function EsComb(ByVal nom As String) As Boolean
    Dim p() As String, i As Long, s As String
    EsComb = False
    s = Normal(nom)
    p = Split(KW_COMB, ";")
    For i = 0 To UBound(p)
        If p(i) <> "" Then
            If InStr(1, s, p(i), vbBinaryCompare) > 0 Then
                EsComb = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function EsGlobal(ByVal und As String) As Boolean
    Dim u As String
    u = Normal(Replace(und, ".", ""))
    EsGlobal = (u = "GLB" Or u = "GLOBAL" Or u = "EST" Or u = "%" Or u = "%MO" Or u = "SG")
End Function

Private Function Excluido(ByVal nom As String) As Boolean
    Dim i As Long, s As String
    Excluido = False
    s = Normal(nom)
    For i = 0 To UBound(excl)
        If excl(i) <> "" Then
            If InStr(1, s, excl(i), vbBinaryCompare) > 0 Then
                Excluido = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub Mapear(ByVal c As Collection, ByVal llave As String, ByVal k As Long, ByVal soloSiFalta As Boolean)
    Dim e As Long
    If llave = "" Then Exit Sub
    e = Buscar(c, llave)
    If e = 0 Then
        c.Add k, llave
    ElseIf e <> k And Not soloSiFalta Then
        c.Remove llave
        c.Add -1, llave
    End If
End Sub

Private Function Controlar(ByVal nomNormal As String) As Boolean
    Controlar = (Buscar(mapCtl, nomNormal) <> 2)
End Function

Private Function SiguienteCodigo(ByVal prefijo As String) As String
    Dim k As Long, mx As Long, ancho As Long, t As String
    mx = 0
    ancho = 2
    For k = 1 To anCat
        If UCase(Left(aCod(k), Len(prefijo))) = UCase(prefijo) Then
            t = Mid(aCod(k), Len(prefijo) + 1)
            If EsDigitos(t) Then
                If Val(t) > mx Then mx = Val(t)
                If Len(t) > ancho Then ancho = Len(t)
            End If
        End If
    Next k
    SiguienteCodigo = prefijo & Format(mx + 1, String(ancho, "0"))
End Function

Private Function AgregarCatalogo(ByVal nom As String, ByVal und As String, ByVal pu As Double) As Long
    ' agrega un insumo al final del CATALOGO y lo deja mapeado
    Dim ws As Worksheet, r As Long, cod As String, gru As String
    AgregarCatalogo = 0
    If anCat >= UBound(aCod) Then Exit Function
    If EsComb(nom) Then
        cod = SiguienteCodigo("C-")
        gru = "COMBUSTIBLE"
    Else
        cod = SiguienteCodigo("M-")
        gru = "MATERIALES"
    End If
    Set ws = Hoja("CATALOGO")
    r = UltimaFila(ws, CC_COD, CT0, CTMAX) + 1
    If r > CTMAX Then Exit Function
    anCat = anCat + 1
    aCod(anCat) = cod
    aDes(anCat) = nom
    aUnd(anCat) = und
    aGru(anCat) = gru
    aPU(anCat) = pu
    cFac(anCat) = 1
    ws.Cells(r, CC_COD).Value = cod
    ws.Cells(r, CC_DES).Value = nom
    ws.Cells(r, CC_UND).Value = und
    ws.Cells(r, CC_GRU).Value = gru
    If pu > 0 Then ws.Cells(r, CC_PU).Value = pu
    ws.Cells(r, CC_FUE).Value = "A.C.U. DEL EXPEDIENTE"
    ws.Cells(r, CC_SANT).Value = 0
    ws.Cells(r, CC_ACU).Value = nom
    ws.Cells(r, CC_FAC).Value = 1
    ws.Cells(r, CC_ULT).Value = "Agregado al importar del control: revise unidad, grupo y P.U."
    aCol.Add anCat, UCase(cod)
    Mapear mapCat, Normal(nom), anCat, False
    nNuevos = nNuevos + 1
    AgregarCatalogo = anCat
End Function

Private Function UnidadEquipo(ByVal und As String) As String
    Dim u As String
    u = Normal(Replace(und, ".", ""))
    Select Case u
        Case "HM", "H-M", "HRS", "HR", "HORA", "HORAS"
            UnidadEquipo = "H.M."
        Case "HE", "H-E"
            UnidadEquipo = "H.E."
        Case "DIA", "DIAS", "D"
            UnidadEquipo = "DÍA"
        Case Else
            If u = "" Then
                UnidadEquipo = "H.M."
            Else
                UnidadEquipo = UCase(und)
            End If
    End Select
End Function

Private Function AgregarEquipo(ByVal nom As String, ByVal und As String) As Long
    Dim ws As Worksheet, r As Long, k As Long, mx As Long, t As String, cod As String
    AgregarEquipo = 0
    If nEqI >= UBound(qCod) Then Exit Function
    Set ws = Hoja("EQUIPOS")
    r = UltimaFila(ws, EC_COD, EQ0, EQMAX) + 1
    If r > EQMAX Then Exit Function
    mx = 0
    For k = 1 To nEqI
        If UCase(Left(qCod(k), 2)) = "E-" Then
            t = Mid(qCod(k), 3)
            If EsDigitos(t) Then
                If Val(t) > mx Then mx = Val(t)
            End If
        End If
    Next k
    cod = "E-" & Format(mx + 1, "00")
    nEqI = nEqI + 1
    qCod(nEqI) = cod
    qComb(nEqI) = 0
    ws.Cells(r, EC_COD).Value = cod
    ws.Cells(r, EC_DES).Value = nom
    ws.Cells(r, EC_UND).Value = UnidadEquipo(und)
    ws.Cells(r, EC_ACU).Value = nom
    ws.Cells(r, EC_ULT).Value = "Agregado al importar del control: complete proveedor, O/S, cantidad contratada y costo."
    Mapear mapEq, Normal(nom), nEqI, False
    nuevosEq = nuevosEq + 1
    AgregarEquipo = nEqI
End Function

Private Sub SinCatalogo(ByVal nom As String)
    If InStr(1, vbLf & sinCat & vbLf, vbLf & "   • " & nom & vbLf, vbTextCompare) > 0 Then Exit Sub
    nSinCat = nSinCat + 1
    If nSinCat <= 12 Then
        If sinCat = "" Then
            sinCat = "   • " & nom
        Else
            sinCat = sinCat & vbLf & "   • " & nom
        End If
    End If
End Sub


Public Function ImportarControlDe(ByVal ruta As String) As String
    Dim cm As Long, wb As Workbook, ya As Boolean, d As Long, i As Long, a As Long, k As Long, j As Long, e As Long, c As Long
    Dim vR As Variant, vA As Variant, vI As Variant, vG As Variant, v As Variant, fechas(1 To 31) As Double
    Dim mesC As Double, nR As Long, nA As Long, nI As Long, nG As Long, last As Long, nm As String, s As String
    Dim opCat As Boolean, opIng As Boolean, opSal As Boolean, opPar As Boolean, colP As Collection, nxtA() As Long
    Dim kp As String, m As Double, q As Double, cod As String, des As String, und As String, f0 As Long
    Dim ws As Worksheet, vM As Variant, oM() As Variant, cntM As Long, nRowsM As Long, orig As String, vacia As Boolean
    Dim vP As Variant, oP() As Variant, nP As Long, nRowsP As Long, tk() As Long, tq() As Double, nt As Long
    Dim ej() As Long, ea() As Double, nE As Long, fk() As Long, fq() As Double, nF As Long, best As Long
    Dim nIng As Long, nSal As Long, nPar As Long, nIngX As Long, hm As Double, msg As String, nombres() As String
    Dim c1 As Variant, c2 As Variant, blk() As Variant, cc As Long, ya2 As Collection, llave As String, nDup As Long

    ImportarControlDe = ""
    opCat = Opcion(O_CAT)
    opIng = Opcion(O_ING)
    opSal = Opcion(O_SAL)
    opPar = Opcion(O_PAR)
    s = Normal(Hoja("DATOS").Range(O_EXC).Value)
    excl = Split(s & ";", ";")
    For i = 0 To UBound(excl)
        excl(i) = Trim(excl(i))
    Next i
    mIni = MesInicio()
    mFin = MesFin()
    nNuevos = 0
    nuevosEq = 0
    sinCat = ""
    nSinCat = 0

    Inicio cm, "Leyendo el control de metrados..."
    On Error GoTo Falla
    Set wb = AbrirLibroX(ruta, ya, True)
    If wb Is Nothing Then Err.Raise 3101, , "No se pudo abrir el archivo."
    If Not (TieneHoja(wb, "RESUMEN") And TieneHoja(wb, "ACU") And TieneHoja(wb, "INSUMOS") And TieneHoja(wb, "INGRESOS") _
            And TieneHoja(wb, "DATOS")) Then
        Err.Raise 3102, , "El archivo no es un CONTROL DE METRADOS (faltan las hojas RESUMEN, ACU, INSUMOS, INGRESOS o DATOS)."
    End If
    mesC = Fecha(wb.Worksheets("DATOS").Range("C19").Value)
    For d = 1 To 31
        fechas(d) = Fecha(wb.Worksheets("RESUMEN").Cells(8, X_D1 + d - 1).Value)
        If fechas(d) = 0 And mesC > 0 Then
            If Month(DateSerial(Year(CDate(mesC)), Month(CDate(mesC)), d)) = Month(CDate(mesC)) Then
                fechas(d) = CDbl(DateSerial(Year(CDate(mesC)), Month(CDate(mesC)), d))
            End If
        End If
    Next d
    last = UltimaFila(wb.Worksheets("RESUMEN"), 2, X_RR0, X_RRMAX)
    nR = last - X_RR0 + 1
    If nR > 0 Then vR = wb.Worksheets("RESUMEN").Range("A" & X_RR0 & ":AM" & last).Value
    last = UltimaFila(wb.Worksheets("ACU"), 2, X_AC0, X_ACMAX)
    nA = last - X_AC0 + 1
    If nA > 0 Then vA = wb.Worksheets("ACU").Range("A" & X_AC0 & ":F" & last).Value
    last = UltimaFila(wb.Worksheets("INSUMOS"), 2, X_IN0, X_INMAX)
    nI = last - X_IN0 + 1
    If nI > 0 Then vI = wb.Worksheets("INSUMOS").Range("A" & X_IN0 & ":F" & last).Value
    last = UltimaFilaVarias(wb.Worksheets("INGRESOS"), 2, 9, X_IG0, X_IGMAX)
    nG = last - X_IG0 + 1
    If nG > 0 Then vG = wb.Worksheets("INGRESOS").Range("A" & X_IG0 & ":P" & last).Value
    If Not ya Then wb.Close False
    Set wb = Nothing
    On Error GoTo 0

    If mesC > 0 Then
        If Format(CDate(mesC), "yyyymm") <> Format(CDate(MesInicio()), "yyyymm") Then
            If Not Pregunta("El CONTROL DE METRADOS está en " & NombreMes(mesC) & " " & Year(CDate(mesC)) & _
                            " y este programa en " & MesTexto() & "." & vbCrLf & _
                            "Solo se importan los días del mes de este programa. ¿Continuar?", "IMPORTAR", False) Then
                Fin cm
                ImportarControlDe = "Importación cancelada."
                Exit Function
            End If
        End If
    End If

    ' ---- catálogo y mapa de nombres del A.C.U.
    Estado "Vinculando insumos..."
    LeerAlmacen
    nCat0 = anCat
    ReDim Preserve aCod(1 To anCat + 2000)
    ReDim Preserve aDes(1 To anCat + 2000)
    ReDim Preserve aUnd(1 To anCat + 2000)
    ReDim Preserve aGru(1 To anCat + 2000)
    ReDim Preserve aPU(1 To anCat + 2000)
    ReDim Preserve aAnt(1 To anCat + 2000)
    ReDim cFac(1 To anCat + 2000)
    Set mapCat = New Collection
    Set ws = Hoja("CATALOGO")
    last = UltimaFila(ws, CC_COD, CT0, CTMAX)
    If last >= CT0 Then
        v = ws.Range(ws.Cells(CT0, 1), ws.Cells(last, CC_ULT)).Value
        For i = 1 To UBound(v, 1)
            k = Buscar(aCol, UCase(txt(v(i, CC_COD))))
            If k > 0 Then
                cFac(k) = Num(v(i, CC_FAC))
                If cFac(k) <= 0 Then cFac(k) = 1
                nombres = Split(txt(v(i, CC_ACU)), ";")
                For j = 0 To UBound(nombres)
                    Mapear mapCat, Normal(nombres(j)), k, False
                Next j
            End If
        Next i
        For i = 1 To UBound(v, 1)
            k = Buscar(aCol, UCase(txt(v(i, CC_COD))))
            If k > 0 Then Mapear mapCat, Normal(v(i, CC_DES)), k, True
        Next i
    End If
    Set mapCtl = New Collection
    For i = 1 To nI
        nm = Normal(vI(i, 2))
        If nm <> "" Then
            If UCase(txt(vI(i, 6))) = "NO" Then
                Mapear mapCtl, nm, 2, True
            Else
                Mapear mapCtl, nm, 1, True
            End If
        End If
    Next i
    ' A.C.U. por partida (lista enlazada en el orden del archivo)
    Set colP = New Collection
    If nA > 0 Then
        ReDim nxtA(1 To nA)
        For a = nA To 1 Step -1
            kp = Clave(txt(vA(a, 1)))
            If kp <> "" Then
                f0 = Buscar(colP, kp)
                nxtA(a) = f0
                If f0 > 0 Then colP.Remove kp
                colP.Add a, kp
            End If
        Next a
    End If

    ' ---- materiales nuevos para el catálogo (con consumo en el mes o con ingreso en el mes)
    If opCat Then
        For i = 1 To nR
            cod = txt(vR(i, 2))
            If cod <> "" Then
                m = 0
                For d = 1 To 31
                    If EnMes(fechas(d)) Then m = m + Num(vR(i, X_D1 + d - 1))
                Next d
                If m > 0 Then
                    a = Buscar(colP, Clave(cod))
                    Do While a > 0
                        If UCase(txt(vA(a, 4))) = "MAT" And Not EsGlobal(txt(vA(a, 3))) Then
                            nm = Normal(vA(a, 2))
                            If Buscar(mapCat, nm) = 0 And Controlar(nm) And Not Excluido(nm) Then
                                k = AgregarCatalogo(Limpio(vA(a, 2)), txt(vA(a, 3)), Num(vA(a, 6)))
                            End If
                        End If
                        a = nxtA(a)
                    Loop
                End If
            End If
        Next i
        If opIng Then
            For i = 1 To nG
                nm = Normal(vG(i, 9))
                If nm <> "" And EnMes(Fecha(vG(i, 2))) Then
                    If Buscar(mapCat, nm) = 0 And Controlar(nm) And Not Excluido(nm) Then
                        k = 0
                        For a = 1 To nI
                            If Normal(vI(a, 2)) = nm Then
                                If UCase(txt(vI(a, 4))) = "MAT" Then k = AgregarCatalogo(Limpio(vI(a, 2)), txt(vI(a, 3)), Num(vI(a, 5)))
                                Exit For
                            End If
                        Next a
                    End If
                End If
            Next i
        End If
    End If

    ' ---- MOVIMIENTOS: se conserva lo manual y se reemplaza lo importado antes
    Estado "Armando los movimientos..."
    Set ws = Hoja("MOVIMIENTOS")
    nRowsM = MVMAX - MV0 + 1
    ReDim oM(1 To nRowsM, 1 To MC_ULT)
    cntM = 0
    Set ya2 = New Collection
    last = UltimaFilaVarias(ws, MC_FEC, MC_ORIG, MV0, MVMAX)
    If last >= MV0 Then
        vM = ws.Range(ws.Cells(MV0, 1), ws.Cells(last, MC_ULT)).Value
        For i = 1 To UBound(vM, 1)
            orig = UCase(txt(vM(i, MC_ORIG)))
            vacia = True
            For c = MC_FEC To MC_ORIG
                If c <> 10 And c <> 11 Then
                    If txt(vM(i, c)) <> "" Then vacia = False
                End If
            Next c
            If Not vacia Then
                If Not ((opIng And orig = UCase(OR_ING)) Or (opSal And orig = UCase(OR_ACU))) Then
                    cntM = cntM + 1
                    For c = MC_FEC To MC_ORIG
                        oM(cntM, c) = vM(i, c)
                    Next c
                    If UCase(txt(vM(i, MC_TIPO))) = T_ING Then
                        llave = "I|" & Fecha(vM(i, MC_FEC)) & "|" & UCase(txt(vM(i, MC_COD))) & "|" & Format(Num(vM(i, MC_CANT)), "0.00")
                    Else
                        llave = "S|" & Fecha(vM(i, MC_FEC)) & "|" & UCase(txt(vM(i, MC_COD))) & "|" & Clave(txt(vM(i, MC_PART)))
                    End If
                    If Buscar(ya2, llave) = 0 Then ya2.Add 1, llave
                End If
            End If
        Next i
    End If
    If opIng Then
        For i = 1 To nG
            If EnMes(Fecha(vG(i, 2))) And txt(vG(i, 9)) <> "" And Num(vG(i, 12)) > 0 Then
                k = Buscar(mapCat, Normal(vG(i, 9)))
                If k > 0 Then
                    llave = "I|" & Fecha(vG(i, 2)) & "|" & UCase(aCod(k)) & "|" & Format(Redondear(Num(vG(i, 12)) * cFac(k), 4), "0.00")
                    If Buscar(ya2, llave) > 0 Then
                        k = -2
                        nDup = nDup + 1
                    End If
                End If
                If k > 0 And cntM < nRowsM Then
                    cntM = cntM + 1
                    nIng = nIng + 1
                    oM(cntM, MC_FEC) = CDate(Fecha(vG(i, 2)))
                    oM(cntM, MC_TIPO) = T_ING
                    oM(cntM, MC_DOC) = Limpio(vG(i, 3))
                    oM(cntM, MC_NDOC) = Limpio(vG(i, 4))
                    oM(cntM, MC_PROC) = Limpio(vG(i, 5))
                    oM(cntM, MC_PART) = txt(vG(i, 13))
                    oM(cntM, MC_COD) = aCod(k)
                    oM(cntM, MC_CANT) = Redondear(Num(vG(i, 12)) * cFac(k), 4)
                    s = Limpio(vG(i, 16))
                    If UCase(txt(vG(i, 15))) = "POR CONFIRMAR" Then s = Trim("POR CONFIRMAR. " & s)
                    oM(cntM, MC_OBS) = s
                    oM(cntM, MC_ORIG) = OR_ING
                ElseIf k = 0 And Not Excluido(Normal(vG(i, 9))) Then
                    nIngX = nIngX + 1
                End If
            End If
        Next i
    End If
    If opSal And nR > 0 And nA > 0 Then
        For d = 1 To 31
            If EnMes(fechas(d)) Then
                For i = 1 To nR
                    m = Num(vR(i, X_D1 + d - 1))
                    cod = txt(vR(i, 2))
                    If m > 0 And cod <> "" Then
                        des = Limpio(vR(i, 3))
                        nt = 0
                        ReDim tk(1 To 60)
                        ReDim tq(1 To 60)
                        a = Buscar(colP, Clave(cod))
                        Do While a > 0
                            If UCase(txt(vA(a, 4))) = "MAT" Then
                                nm = Normal(vA(a, 2))
                                k = Buscar(mapCat, nm)
                                If k > 0 Then
                                    q = m * Num(vA(a, 5)) * cFac(k)
                                    e = 0
                                    For j = 1 To nt
                                        If tk(j) = k Then e = j
                                    Next j
                                    If e = 0 And nt < 60 Then
                                        nt = nt + 1
                                        tk(nt) = k
                                        tq(nt) = q
                                    ElseIf e > 0 Then
                                        tq(e) = tq(e) + q
                                    End If
                                ElseIf Controlar(nm) And Not Excluido(nm) And Not EsGlobal(txt(vA(a, 3))) Then
                                    SinCatalogo Limpio(vA(a, 2))
                                End If
                            End If
                            a = nxtA(a)
                        Loop
                        For j = 1 To nt
                            llave = "S|" & fechas(d) & "|" & UCase(aCod(tk(j))) & "|" & Clave(cod)
                            If Buscar(ya2, llave) > 0 Then
                                tq(j) = 0
                                nDup = nDup + 1
                            End If
                            If tq(j) >= 0.005 And cntM < nRowsM Then
                                cntM = cntM + 1
                                nSal = nSal + 1
                                k = tk(j)
                                oM(cntM, MC_FEC) = CDate(fechas(d))
                                oM(cntM, MC_TIPO) = T_SAL
                                If aGru(k) = "COMBUSTIBLE" Or EsComb(aDes(k)) Then
                                    oM(cntM, MC_DOC) = "PARTE DIARIO DE MAQUINARIA"
                                Else
                                    oM(cntM, MC_DOC) = "VALE DE SALIDA"
                                End If
                                oM(cntM, MC_PROC) = cod & " " & des
                                oM(cntM, MC_PART) = cod
                                oM(cntM, MC_COD) = aCod(k)
                                oM(cntM, MC_CANT) = Redondear(tq(j), 2)
                                oM(cntM, MC_OBS) = "Según A.C.U. x metrado del día; ajustar con el vale"
                                oM(cntM, MC_ORIG) = OR_ACU
                            End If
                        Next j
                    End If
                Next i
            End If
        Next d
    End If
    c1 = Array(MC_FEC, MC_CANT)
    c2 = Array(MC_COD, MC_ORIG)
    For j = 0 To 1
        ws.Range(ws.Cells(MV0, c1(j)), ws.Cells(MVMAX, c2(j))).ClearContents
        If cntM > 0 Then
            ReDim blk(1 To cntM, 1 To c2(j) - c1(j) + 1)
            For i = 1 To cntM
                For cc = c1(j) To c2(j)
                    blk(i, cc - c1(j) + 1) = oM(i, cc)
                Next cc
            Next i
            ws.Range(ws.Cells(MV0, c1(j)), ws.Cells(MV0 + cntM - 1, c2(j))).Value = blk
        End If
    Next j
    ws.Range(ws.Cells(MV0, MC_SALDO), ws.Cells(MVMAX, MC_SALDO)).ClearContents

    ' ---- PARTES de los equipos según A.C.U.
    If opPar And nR > 0 And nA > 0 Then
        Estado "Armando los partes de los equipos..."
        Set mapEq = New Collection
        Set ws = Hoja("EQUIPOS")
        last = UltimaFila(ws, EC_COD, EQ0, EQMAX)
        ReDim qCod(1 To EQMAX - EQ0 + 1)
        ReDim qComb(1 To EQMAX - EQ0 + 1)
        nEqI = 0
        If last >= EQ0 Then
            v = ws.Range(ws.Cells(EQ0, 1), ws.Cells(last, EC_ULT)).Value
            For i = 1 To UBound(v, 1)
                If txt(v(i, EC_COD)) <> "" Then
                    nEqI = nEqI + 1
                    qCod(nEqI) = txt(v(i, EC_COD))
                    qComb(nEqI) = Buscar(aCol, UCase(txt(v(i, EC_COMB))))
                    nombres = Split(txt(v(i, EC_ACU)), ";")
                    For j = 0 To UBound(nombres)
                        Mapear mapEq, Normal(nombres(j)), nEqI, False
                    Next j
                End If
            Next i
            j = 0
            For i = 1 To UBound(v, 1)
                If txt(v(i, EC_COD)) <> "" Then
                    j = j + 1
                    Mapear mapEq, Normal(v(i, EC_DES)), j, True
                End If
            Next i
        End If
        nEq0 = nEqI
        Set ws = Hoja("PARTES")
        nRowsP = PTMAX - PT0 + 1
        ReDim oP(1 To nRowsP, 1 To PC_ULT)
        nP = 0
        last = UltimaFilaVarias(ws, PC_FEC, PC_ORIG, PT0, PTMAX)
        If last >= PT0 Then
            vP = ws.Range(ws.Cells(PT0, 1), ws.Cells(last, PC_ULT)).Value
            For i = 1 To UBound(vP, 1)
                vacia = True
                For c = PC_FEC To PC_ORIG
                    If c <> 4 And c <> PC_HACU And c <> PC_HVAL Then
                        If txt(vP(i, c)) <> "" Then vacia = False
                    End If
                Next c
                If Not vacia And UCase(txt(vP(i, PC_ORIG))) <> UCase(OR_ACU) Then
                    nP = nP + 1
                    For c = PC_FEC To PC_ORIG
                        oP(nP, c) = vP(i, c)
                    Next c
                    llave = "P|" & Fecha(vP(i, PC_FEC)) & "|" & UCase(txt(vP(i, PC_EQ))) & "|" & Clave(txt(vP(i, PC_PART)))
                    If Buscar(ya2, llave) = 0 Then ya2.Add 1, llave
                End If
            Next i
        End If
        For d = 1 To 31
            If EnMes(fechas(d)) Then
                For i = 1 To nR
                    m = Num(vR(i, X_D1 + d - 1))
                    cod = txt(vR(i, 2))
                    If m > 0 And cod <> "" Then
                        des = Limpio(vR(i, 3))
                        und = txt(vR(i, 4))
                        nE = 0
                        ReDim ej(1 To 30)
                        ReDim ea(1 To 30)
                        nF = 0
                        ReDim fk(1 To 10)
                        ReDim fq(1 To 10)
                        a = Buscar(colP, Clave(cod))
                        Do While a > 0
                            nm = Normal(vA(a, 2))
                            If UCase(txt(vA(a, 4))) = "EQ" And InStr(1, nm, "HERRAMIENTA", vbBinaryCompare) = 0 _
                               And Not Excluido(nm) And Num(vA(a, 5)) > 0 Then
                                j = Buscar(mapEq, nm)
                                If j = 0 Then j = AgregarEquipo(Limpio(vA(a, 2)), txt(vA(a, 3)))
                                If j > 0 Then
                                    e = 0
                                    For c = 1 To nE
                                        If ej(c) = j Then e = c
                                    Next c
                                    If e = 0 And nE < 30 Then
                                        nE = nE + 1
                                        ej(nE) = j
                                        ea(nE) = Num(vA(a, 5))
                                    ElseIf e > 0 Then
                                        If Num(vA(a, 5)) > ea(e) Then ea(e) = Num(vA(a, 5))
                                    End If
                                End If
                            ElseIf UCase(txt(vA(a, 4))) = "MAT" Then
                                k = Buscar(mapCat, nm)
                                If k > 0 Then
                                    If aGru(k) = "COMBUSTIBLE" Or EsComb(aDes(k)) Then
                                        e = 0
                                        For c = 1 To nF
                                            If fk(c) = k Then e = c
                                        Next c
                                        If e = 0 And nF < 10 Then
                                            nF = nF + 1
                                            fk(nF) = k
                                            fq(nF) = m * Num(vA(a, 5)) * cFac(k)
                                        ElseIf e > 0 Then
                                            fq(e) = fq(e) + m * Num(vA(a, 5)) * cFac(k)
                                        End If
                                    End If
                                End If
                            End If
                            a = nxtA(a)
                        Loop
                        For e = 1 To nE
                            llave = "P|" & fechas(d) & "|" & UCase(qCod(ej(e))) & "|" & Clave(cod)
                            If Buscar(ya2, llave) > 0 Then
                                nDup = nDup + 1
                            ElseIf nP < nRowsP Then
                                nP = nP + 1
                                nPar = nPar + 1
                                j = ej(e)
                                oP(nP, PC_FEC) = CDate(fechas(d))
                                oP(nP, PC_EQ) = qCod(j)
                                oP(nP, PC_PART) = cod
                                oP(nP, PC_TRAB) = des
                                oP(nP, PC_MET) = m
                                oP(nP, PC_UND) = und
                                oP(nP, PC_REND) = ea(e)
                                oP(nP, PC_ORIG) = OR_ACU
                                ' combustible del día: al equipo de más horas que usa ese combustible
                                For c = 1 To nF
                                    If qComb(j) = fk(c) And fq(c) >= 0.005 Then
                                        best = 0
                                        hm = -1
                                        For cc = 1 To nE
                                            If qComb(ej(cc)) = fk(c) Then
                                                If ea(cc) > hm Then
                                                    hm = ea(cc)
                                                    best = cc
                                                End If
                                            End If
                                        Next cc
                                        If best = e Then oP(nP, PC_COMB) = Redondear(fq(c), 2)
                                    End If
                                Next c
                            End If
                        Next e
                    End If
                Next i
            End If
        Next d
        c1 = Array(PC_FEC, PC_NPARTE, PC_HREAL, PC_COMB)
        c2 = Array(PC_EQ, PC_REND, PC_HREAL, PC_ORIG)
        For j = 0 To 3
            ws.Range(ws.Cells(PT0, c1(j)), ws.Cells(PTMAX, c2(j))).ClearContents
            If nP > 0 Then
                ReDim blk(1 To nP, 1 To c2(j) - c1(j) + 1)
                For i = 1 To nP
                    For cc = c1(j) To c2(j)
                        blk(i, cc - c1(j) + 1) = oP(i, cc)
                    Next cc
                Next i
                ws.Range(ws.Cells(PT0, c1(j)), ws.Cells(PT0 + nP - 1, c2(j))).Value = blk
            End If
        Next j
        ws.Range(ws.Cells(PT0, PC_CTRL), ws.Cells(PTMAX, PC_CTRL)).ClearContents
    End If
    Fin cm

    msg = "Del CONTROL DE METRADOS (" & MesTexto() & "):" & vbCrLf
    If opCat Then msg = msg & "   • Insumos nuevos en el CATÁLOGO: " & nNuevos & IIf(nNuevos > 0, _
                         " (revise su unidad, grupo y P.U.)", "") & vbCrLf
    If opIng Then
        msg = msg & "   • INGRESOS importados: " & nIng
        If nIngX > 0 Then msg = msg & "  (" & nIngX & " sin vínculo con el catálogo)"
        msg = msg & vbCrLf
    End If
    If opSal Then msg = msg & "   • SALIDAS según A.C.U. x metrado: " & nSal & " filas" & vbCrLf
    If opPar Then msg = msg & "   • PARTES según A.C.U. x metrado: " & nPar & " filas" & IIf(nuevosEq > 0, _
                         " (equipos nuevos: " & nuevosEq & ")", "") & vbCrLf
    msg = msg & "   • Lo digitado a mano se conservó."
    If nDup > 0 Then msg = msg & vbCrLf & "   • Se omitieron " & nDup & " filas que ya estaban registradas a mano (mismo día, insumo o equipo y partida)."
    If nSinCat > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Materiales con consumo que no están en el CATÁLOGO (sus salidas no se calcularon; " & _
              "agréguelos o anote su nombre del A.C.U. en la columna NOMBRE EN EL A.C.U.):" & vbCrLf & sinCat
        If nSinCat > 12 Then msg = msg & vbLf & "   ... y " & (nSinCat - 12) & " más."
    End If
    msg = msg & vbCrLf & vbCrLf & "Siguiente paso: PROCESAR Y VALIDAR."
    Mensaje D_MSG, "Importado del control el " & Format(Now, "dd/mm/yyyy hh:mm") & ": " & nIng & " ingresos, " & nSal & _
            " salidas y " & nPar & " partes (A.C.U.)."
    Hoja("DATOS").Range(D_CTRL).Value = ruta
    ImportarControlDe = msg
    Exit Function
Falla:
    msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then
        If Not ya Then wb.Close False
    End If
    Fin cm
    ImportarControlDe = "No se pudo importar: " & msg
End Function
