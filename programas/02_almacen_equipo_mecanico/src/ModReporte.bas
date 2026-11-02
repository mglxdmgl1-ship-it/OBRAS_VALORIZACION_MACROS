Attribute VB_Name = "ModReporte"
Option Explicit

'==========================================================================================
'  CUADROS DEL INFORME DENTRO DEL PROGRAMA (como el informe de ejemplo, todo en un solo archivo)
'   ALMACÉN : INV. 2026  -> se digita: un renglón por ingreso (O/C, P.U., cantidad, fecha); los
'                           renglones sin unidad ni cantidad son los RUBROS (1, 2, 3...)
'             4.10 MOV   -> se digita: saldo anterior y las cantidades de cada día (ENTRADA/SALIDA)
'             4.11 y 4.12 -> se calculan solos con fórmulas
'   EQUIPOS : EQUIPOS    -> se digita: datos de cada máquina
'             4.15       -> se digita: combustible (GAL) y horas de cada día
'  ACTUALIZAR CUADROS arma las filas de 4.10, 4.11, 4.12 y 4.15 con lo que hay en INV. y EQUIPOS,
'  sin perder lo ya digitado. Kardex, bincard y partes diarios se arman desde estos cuadros.
'==========================================================================================

Private Const GRIS_DOMINGO As Long = 14277081      ' RGB(217,217,217)
Private Const F0 As Long = 13                       ' primera fila de datos de 4.10, 4.11 y 4.12
Private Const F0Q As Long = 14                      ' primera fila de datos de 4.15
Private Const D1 As Long = 8                        ' 4.10 / 4.11: columna del día 1 (H)
Private Const D1Q As Long = 10                      ' 4.15: columna del día 1 (J)
Private Const C_RAR As Long = 44                    ' 4.15: código en la fila de horas (AR, oculta)
Private Const C_RAS As Long = 45                    ' 4.15: código en la fila de combustible (AS, oculta)
Private Const C_RAT As Long = 46                    ' 4.15: código del combustible (AT, oculta)

' ---- inventario leído (INV.)
Public cN As Long, rN As Long
Public cItem() As String, cDes() As String, cUnd() As String, cTipo() As String, cReq() As String, cProv() As String
Public cPU() As Double, cCant() As Double, cFec() As Double, cRub() As Long
Public rNom() As String, rNum() As String
Private nDup As Long
' ---- lo digitado en 4.10 (por ítem)
Private tN As Long, tCol As Collection
Private tItem() As String, tG() As Variant, tE() As Variant, tS() As Variant, tUsado() As Boolean
' ---- lo digitado en 4.15 (por equipo)
Private qN As Long, qCol As Collection
Private qCod() As String, qG() As Variant, qH() As Variant


'========================================================================== AUXILIARES
Public Function HojaPref(ByVal pref As String) As Worksheet
    ' hoja cuyo nombre empieza con pref ("INV.", "4.10", "4.11", "4.12", "4.15")
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, Len(pref))) = UCase(pref) Then
            Set HojaPref = ws
            Exit Function
        End If
    Next ws
    Err.Raise 9, , "No se encontró la hoja " & pref
End Function

Private Function MesNombre(ByVal t As Double) As String
    Dim s As String
    s = LCase(NombreMes(t))
    MesNombre = UCase(Left(s, 1)) & Mid(s, 2)
End Function

Private Function MesAnterior() As Double
    MesAnterior = CDbl(DateSerial(Year(CDate(MesInicio())), Month(CDate(MesInicio())) - 1, 1))
End Function

Private Function DiasMes() As Long
    DiasMes = Day(CDate(MesFin()))
End Function

Private Function DiaFecha(ByVal d As Long) As Double
    DiaFecha = CDbl(DateSerial(Year(CDate(MesInicio())), Month(CDate(MesInicio())), d))
End Function

Private Function LetraDia(ByVal t As Double) As String
    LetraDia = Choose(Weekday(CDate(t), vbMonday), "L", "M", "M", "J", "V", "S", "D")
End Function

Private Function Dato(ByVal celda As String) As String
    Dato = Limpio(Hoja("DATOS").Range(celda).Value)
End Function

Private Function ConDos(ByVal s As String) As String
    ConDos = ": " & s
End Function

Public Function Doc(ByVal s As String) As String
    s = Replace(s, "O/C N° ", "O/C ")
    s = Replace(s, "O/S N° ", "O/S ")
    s = Replace(s, "O/C N°", "O/C ")
    Doc = Trim(s)
End Function

Private Function Dir1(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dir1 = ws.Cells(r, c).Address(False, False)
End Function

Private Function NumF(ByVal x As Double) As String
    NumF = Trim(Str(x))
End Function

Private Sub CopiarDe(ByVal wsT As Worksheet, ByVal a As Long, ByVal b As Long, ByVal ws As Worksheet, ByVal r As Long)
    ' copia las filas modelo a..b de la plantilla FMT_ a la fila r
    Dim i As Long
    wsT.Rows(a & ":" & b).Copy Destination:=ws.Rows(r)
    For i = 0 To b - a
        ws.Rows(r + i).RowHeight = wsT.Rows(a + i).RowHeight
    Next i
End Sub

Private Sub LimpiarDatos(ByVal ws As Worksheet, ByVal desde As Long)
    ' borra las filas de datos (desde la fila indicada hasta la última usada)
    Dim last As Long
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 5
    If last < desde + 10 Then last = desde + 10
    ' se limpian (no se eliminan) para que las fórmulas que apuntan a este cuadro (INICIO, EQUIPOS) no se acorten
    With ws.Range(ws.Rows(desde), ws.Rows(last))
        .Clear
        .RowHeight = ws.StandardHeight
        .EntireRow.Hidden = False
    End With
End Sub

Private Sub Domingos(ByVal ws As Worksheet, ByVal col1 As Long, ByVal f1 As Long, ByVal f2 As Long)
    Dim d As Long
    If f2 < f1 Then Exit Sub
    For d = 1 To DiasMes()
        If Weekday(CDate(DiaFecha(d)), vbMonday) = 7 Then
            ws.Range(ws.Cells(f1, col1 + d - 1), ws.Cells(f2, col1 + d - 1)).Interior.Color = GRIS_DOMINGO
        End If
    Next d
End Sub

Private Sub CabeceraDias(ByVal ws As Worksheet, ByVal fLetra As Long, ByVal fNum As Long, ByVal col1 As Long)
    Dim d As Long
    For d = 1 To 31
        If d <= DiasMes() Then
            ws.Cells(fLetra, col1 + d - 1).Value = LetraDia(DiaFecha(d))
            ws.Cells(fNum, col1 + d - 1).Value = d
            ws.Columns(col1 + d - 1).Hidden = False
        Else
            ws.Cells(fLetra, col1 + d - 1).Value = ""
            ws.Cells(fNum, col1 + d - 1).Value = ""
            ws.Columns(col1 + d - 1).Hidden = True
        End If
    Next d
End Sub

Private Function SumaFila(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long) As String
    SumaFila = "=SUM(" & Dir1(ws, r, c1) & ":" & Dir1(ws, r, c2) & ")"
End Function

Private Function EsRubro(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    EsRubro = (Txt(ws.Cells(r, 3).Value) = "" And Txt(ws.Cells(r, 5).Value) = "" And Txt(ws.Cells(r, 6).Value) = "")
End Function


'========================================================================== LECTURA DE INV. Y 4.10
Public Sub LeerInv()
    Dim ws As Worksheet, last As Long, r As Long, v As Variant, a As String, b As String, cur As Long, n As Long
    Dim cnt() As Long, col As New Collection
    Set ws = HojaPref("INV.")
    last = UltimaFilaVarias(ws, 1, 2, 2, 6000)
    n = last
    If n < 2 Then n = 2
    ReDim cItem(1 To n): ReDim cDes(1 To n): ReDim cUnd(1 To n): ReDim cTipo(1 To n): ReDim cReq(1 To n)
    ReDim cProv(1 To n): ReDim cPU(1 To n): ReDim cCant(1 To n): ReDim cFec(1 To n): ReDim cRub(1 To n)
    ReDim rNom(1 To n): ReDim rNum(1 To n): ReDim cnt(1 To n)
    cN = 0: rN = 0: nDup = 0: cur = 0
    If last < 2 Then Exit Sub
    v = ws.Range("A2:J" & last).Value
    For r = 1 To UBound(v, 1)
        a = Txt(v(r, 1))
        b = Limpio(v(r, 2))
        If a <> "" Or b <> "" Then
            If Txt(v(r, 3)) = "" And Txt(v(r, 5)) = "" And Txt(v(r, 6)) = "" Then
                rN = rN + 1
                rNom(rN) = UCase(b)
                rNum(rN) = IIf(a = "", CStr(rN), a)
                cur = rN
            Else
                If cur = 0 Then
                    rN = rN + 1
                    rNom(rN) = "SIN RUBRO"
                    rNum(rN) = CStr(rN)
                    cur = rN
                End If
                cnt(cur) = cnt(cur) + 1
                If a = "" Then a = rNum(cur) & "." & cnt(cur)
                If Buscar(col, UCase(a)) > 0 Then
                    nDup = nDup + 1
                    a = a & "-" & nDup
                End If
                cN = cN + 1
                col.Add cN, UCase(a)
                cItem(cN) = a
                cDes(cN) = b
                cUnd(cN) = Txt(v(r, 3))
                cTipo(cN) = Doc(Limpio(v(r, 4)))
                cPU(cN) = Num(v(r, 5))
                cCant(cN) = Num(v(r, 6))
                cFec(cN) = Fecha(v(r, 8))
                cReq(cN) = Limpio(v(r, 9))
                cProv(cN) = Limpio(v(r, 10))
                cRub(cN) = cur
            End If
        End If
    Next r
End Sub

Private Sub Leer410()
    ' cantidades ya digitadas en 4.10, por ítem
    Dim ws As Worksheet, last As Long, r As Long, d As Long, key As String, n As Long
    Set ws = HojaPref("4.10")
    Set tCol = New Collection
    tN = 0
    last = UltimaFila(ws, 6, F0, 8000)
    n = (last - F0 + 2) \ 2
    If n < 1 Then n = 1
    ReDim tItem(1 To n): ReDim tG(1 To n): ReDim tE(1 To n, 1 To 31): ReDim tS(1 To n, 1 To 31): ReDim tUsado(1 To n)
    If last < F0 Then Exit Sub
    For r = F0 To last
        If UCase(Txt(ws.Cells(r, 6).Value)) = "ENTRADA" Then
            key = Txt(ws.Cells(r, 2).Value)
            If key <> "" And Buscar(tCol, UCase(key)) = 0 Then
                tN = tN + 1
                tItem(tN) = key
                tG(tN) = ws.Cells(r, 7).Value
                For d = 1 To 31
                    tE(tN, d) = ws.Cells(r, D1 + d - 1).Value
                    tS(tN, d) = ws.Cells(r + 1, D1 + d - 1).Value
                Next d
                tCol.Add tN, UCase(key)
            End If
        End If
    Next r
End Sub

Private Sub Leer415()
    ' combustible y horas ya digitados en 4.15, por código de equipo
    Dim ws As Worksheet, last As Long, r As Long, d As Long, key As String, n As Long
    Set ws = HojaPref("4.15")
    Set qCol = New Collection
    qN = 0
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    n = last + EQMAX
    If n < 1 Then n = 1
    ReDim qCod(1 To n): ReDim qG(1 To n, 1 To 31): ReDim qH(1 To n, 1 To 31)
    For r = F0Q To last
        key = Txt(ws.Cells(r, C_RAS).Value)
        If key <> "" And Buscar(qCol, UCase(key)) = 0 Then
            qN = qN + 1
            qCod(qN) = key
            For d = 1 To 31
                qG(qN, d) = ws.Cells(r, D1Q + d - 1).Value
                qH(qN, d) = ws.Cells(r + 1, D1Q + d - 1).Value
            Next d
            qCol.Add qN, UCase(key)
        End If
    Next r
    ' equipos con filas en MOV. EQUIPOS: sus horas y galones del 4.15 salen de ahí
    Dim x As Long, i As Long, k As Long, cods As New Collection
    LeerMovEq
    For x = 1 To mqN
        If Buscar(cods, UCase(mqCod(x))) = 0 Then
            cods.Add x, UCase(mqCod(x))
            i = Buscar(qCol, UCase(mqCod(x)))
            If i = 0 Then
                qN = qN + 1
                i = qN
                qCod(i) = mqCod(x)
                qCol.Add i, UCase(mqCod(x))
            End If
            For d = 1 To 31
                qG(i, d) = Empty
                qH(i, d) = Empty
            Next d
        End If
        i = Buscar(qCol, UCase(mqCod(x)))
        d = Day(CDate(mqFec(x)))
        k = Buscar(eCol, UCase(mqCod(x)))
        If mqG(x) <> 0 Then qG(i, d) = Num(qG(i, d)) + mqG(x)
        If mqH(x) <> 0 Then
            If k > 0 Then
                If EsUnidadHora(eUnd(k)) Then
                    qH(i, d) = Num(qH(i, d)) + mqH(x) / 24     ' horas como tiempo ([h]:mm)
                Else
                    qH(i, d) = Num(qH(i, d)) + mqH(x)
                End If
            Else
                qH(i, d) = Num(qH(i, d)) + mqH(x)
            End If
        End If
    Next x
End Sub


'========================================================================== ACTUALIZAR CUADROS
Public Sub ActualizarCuadrosBoton()
    Dim s As String
    s = ActualizarCuadros(False)
End Sub

Public Function ActualizarCuadros(ByVal silencio As Boolean) As String
    ' arma 4.10, 4.11, 4.12 y 4.15 con INV. y EQUIPOS, conservando lo digitado; devuelve un resumen
    Dim cm As Long, msg As String, neg As String, nNeg As Long, perd As String, nPerd As Long, k As Long, i As Long
    Dim fila410() As Long, d As Long, s As Double
    ActualizarCuadros = ""
    Inicio cm, "Actualizando los cuadros 4.10, 4.11, 4.12 y 4.15..."
    On Error GoTo Falla
    LeerInv
    Leer410
    ReDim fila410(1 To cN + 1)
    Armar410 fila410
    Armar411
    Armar412 fila410
    Armar415
    ' saldos negativos y datos de ítems que ya no están en INV.
    For k = 1 To cN
        i = Buscar(tCol, UCase(cItem(k)))
        If i > 0 Then
            tUsado(i) = True
            s = Num(tG(i))
            For d = 1 To 31
                s = s + Num(tE(i, d)) - Num(tS(i, d))
            Next d
            If s < -0.00001 Then
                nNeg = nNeg + 1
                If nNeg <= 12 Then neg = neg & vbCrLf & "   " & cItem(k) & "  " & Left(cDes(k), 45) & ":  " & Format(s, "0.00")
            End If
        End If
    Next k
    For i = 1 To tN
        If Not tUsado(i) Then
            For d = 1 To 31
                If Num(tE(i, d)) <> 0 Or Num(tS(i, d)) <> 0 Then
                    nPerd = nPerd + 1
                    If nPerd <= 8 Then perd = perd & " " & tItem(i)
                    Exit For
                End If
            Next d
        End If
    Next i
    Fin cm
    ' página de cada cuadro: todo el ancho en una hoja, títulos repetidos, sin cortes
    AjustarInforme HojaPref("INV.")
    AjustarInforme HojaPref("4.10")
    AjustarInforme HojaPref("4.11")
    AjustarInforme HojaPref("4.12")
    AjustarInforme HojaPref("4.15")
    SincronizarAuto                               ' equipos y rubros en AUTOLLENADO
    ListarSalidas                                 ' una fila por salida del 4.10 en DETALLE SALIDAS
    msg = "Cuadros actualizados (" & MesTexto() & "): " & cN & " ítems en " & rN & " rubros del INV.; " & enEq & " equipos en el 4.15."
    If nDup > 0 Then msg = msg & vbCrLf & "Atención: " & nDup & " ítem(s) con número repetido en INV. (se les agregó -1, -2...)."
    If nNeg > 0 Then msg = msg & vbCrLf & vbCrLf & "SALDO NEGATIVO en " & nNeg & " ítem(s) (revise las salidas del 4.10):" & neg
    If nPerd > 0 Then msg = msg & vbCrLf & vbCrLf & "Atención: había cantidades en el 4.10 de ítems que ya no están en INV.:" & perd & _
                            " (se quitaron; si fue un cambio de número, vuelva a digitarlas)."
    Mensaje D_MSG, "Cuadros actualizados el " & Format(Now, "dd/mm/yyyy hh:mm") & ": " & cN & " ítems, " & enEq & _
                   " equipos, " & nNeg & " con saldo negativo."
    ActualizarCuadros = msg
    If Not silencio Then Aviso msg, "ACTUALIZAR CUADROS"
    Exit Function
Falla:
    Fin cm
    ActualizarCuadros = "No se pudieron actualizar los cuadros: " & Err.Description
    Aviso ActualizarCuadros, "ACTUALIZAR CUADROS", 16
End Function

Private Sub Encabezado(ByVal ws As Worksheet, ByVal tit As String, ByVal cP As Long, ByVal cAno As Long, _
                       ByVal fP As Long, ByVal fMod As Long, ByVal fRes As Long, ByVal fSup As Long, ByVal fAno As Long, _
                       ByVal fMes As Long)
    ws.Cells(1, 2).Value = Dato(D_ENTIDAD) & vbLf & Dato(D_GEREN) & vbLf
    ws.Cells(3, 2).Value = tit & " - " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    ws.Cells(fP, cP).Value = ConDos("“" & Dato(D_PROY) & "”.")
    ws.Cells(fMod, cP).Value = ConDos(Dato(D_MODAL))
    ws.Cells(fRes, cP).Value = ConDos(Dato(D_RESID))
    ws.Cells(fSup, cP).Value = ConDos(Dato(D_SUPER))
    ws.Cells(fSup, 2).Value = Dato(D_F4)
    ws.Cells(fAno, cAno).Value = Year(CDate(MesInicio()))
    ws.Cells(fMes, cAno).Value = MesNombre(MesInicio())
End Sub

Private Sub Armar410(ByRef fila410() As Long)
    Dim ws As Worksheet, wsT As Worksheet, r As Long, g As Long, k As Long, i As Long, d As Long, v() As Variant
    Dim hayE As Boolean
    Set ws = HojaPref("4.10")
    Set wsT = Hoja("FMT_410")
    Encabezado ws, "4.10 CUADRO DE MOVIMIENTO DIARIO DE ALMACEN", 4, 40, 5, 6, 7, 8, 5, 7
    ws.Cells(10, 7).Value = "SALDO" & vbLf & NombreMes(MesAnterior()) & " " & Year(CDate(MesAnterior()))
    ws.Cells(10, 39).Value = "TOTAL"
    ws.Cells(10, 40).Value = "SALDO" & vbLf & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    CabeceraDias ws, 10, 11, D1
    LimpiarDatos ws, F0
    r = F0
    For g = 1 To rN
        CopiarDe wsT, 13, 14, ws, r
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = rNum(g)
        ws.Cells(r, 3).Value = rNom(g)
        r = r + 2
        For k = 1 To cN
            If cRub(k) = g Then
                CopiarDe wsT, 15, 16, ws, r
                fila410(k) = r
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 2).Value = cItem(k)
                ws.Cells(r, 3).Value = cDes(k)
                ws.Cells(r, 4).Value = cUnd(k)
                ws.Cells(r, 5).Value = cTipo(k)
                ws.Cells(r, 6).Value = "ENTRADA"
                ws.Cells(r + 1, 6).Value = "SALIDA"
                ReDim v(1 To 2, 1 To 31)
                i = Buscar(tCol, UCase(cItem(k)))
                hayE = False
                If i > 0 Then
                    ws.Cells(r, 7).Value = tG(i)
                    For d = 1 To 31
                        v(1, d) = tE(i, d)
                        v(2, d) = tS(i, d)
                        If Num(tE(i, d)) <> 0 Then hayE = True
                    Next d
                End If
                ' el ingreso del INV. se pone solo en su día (si la fila ENTRADA está vacía)
                If Not hayE And cCant(k) <> 0 And cFec(k) >= MesInicio() And cFec(k) <= MesFin() Then
                    v(1, Day(CDate(cFec(k)))) = cCant(k)
                End If
                If i = 0 And IsEmpty(ws.Cells(r, 7).Value) Then ws.Cells(r, 7).Value = 0
                ws.Range(ws.Cells(r, D1), ws.Cells(r + 1, D1 + 30)).Value = v
                ws.Cells(r, 39).Formula = SumaFila(ws, r, D1, D1 + 30)
                ws.Cells(r + 1, 39).Formula = SumaFila(ws, r + 1, D1, D1 + 30)
                ws.Cells(r, 40).Formula = "=" & Dir1(ws, r, 7) & "+" & Dir1(ws, r, 39) & "-" & Dir1(ws, r + 1, 39)
                r = r + 2
            End If
        Next k
    Next g
    Domingos ws, D1, F0, r - 1
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$AN$" & IIf(r - 1 < F0, F0, r - 1)
    ws.PageSetup.PrintTitleRows = "$10:$12"
    On Error GoTo 0
End Sub

Private Sub Armar411()
    ' 4.11: mismas filas que el 4.10, todo con fórmulas que leen el 4.10
    Dim ws As Worksheet, wsT As Worksheet, w10 As Worksheet, r As Long, last As Long, h As String, c As Long
    Set ws = HojaPref("4.11")
    Set wsT = Hoja("FMT_411")
    Set w10 = HojaPref("4.10")
    h = "'" & w10.Name & "'!"
    Encabezado ws, "4.11 CUADRO RESUMEN DE ALMACÉN (INGRESOS – EGRESOS)", 3, 42, 5, 6, 7, 8, 5, 7
    ws.Cells(10, 7).Value = "SALDO " & NombreMes(MesAnterior()) & " " & Year(CDate(MesAnterior()))
    ws.Cells(10, 39).Value = "TOTAL"
    ws.Cells(10, 40).Value = "INGRESOS"
    ws.Cells(10, 41).Value = "SALIDAS"
    ws.Cells(10, 42).Value = "SALDO " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    CabeceraDias ws, 10, 11, D1
    LimpiarDatos ws, F0
    last = UltimaFila(w10, 6, F0, 8000)
    r = F0
    Do While r <= last
        If Txt(w10.Cells(r, 6).Value) = "ENTRADA" Then
            CopiarDe wsT, 15, 16, ws, r
            For c = 2 To 7
                ws.Cells(r, c).Formula = "=" & h & Dir1(w10, r, c) & "&"""""
            Next c
            ws.Cells(r, 7).Formula = "=N(" & h & Dir1(w10, r, 7) & ")"
            ws.Cells(r + 1, 6).Formula = "=" & h & Dir1(w10, r + 1, 6) & "&"""""
            ws.Range(ws.Cells(r, D1), ws.Cells(r + 1, D1 + 30)).FormulaR1C1 = "=IF(" & h & "RC="""",""""," & h & "RC)"
            ws.Cells(r, 39).Formula = SumaFila(ws, r, D1, D1 + 30)
            ws.Cells(r + 1, 39).Formula = SumaFila(ws, r + 1, D1, D1 + 30)
            ws.Cells(r, 40).Formula = "=" & Dir1(ws, r, 39)
            ws.Cells(r, 41).Formula = "=" & Dir1(ws, r + 1, 39)
            ws.Cells(r, 42).Formula = "=" & Dir1(ws, r, 7) & "+" & Dir1(ws, r, 39) & "-" & Dir1(ws, r + 1, 39)
        Else
            CopiarDe wsT, 13, 14, ws, r
            ws.Cells(r, 2).Formula = "=" & h & Dir1(w10, r, 2) & "&"""""
            ws.Cells(r, 3).Formula = "=" & h & Dir1(w10, r, 3) & "&"""""
        End If
        r = r + 2
    Loop
    Domingos ws, D1, F0, r - 1
    On Error Resume Next
    ws.PageSetup.PrintArea = "$B$1:$AP$" & IIf(r - 1 < F0, F0, r - 1)
    ws.PageSetup.PrintTitleRows = "$10:$12"
    On Error GoTo 0
End Sub

Private Sub Armar412(ByRef fila410() As Long)
    ' 4.12: saldo anterior, ingresos, salidas y saldo, valorizados con el P.U. del INV.
    Dim ws As Worksheet, wsT As Worksheet, w10 As Worksheet, r As Long, g As Long, k As Long, f1 As Long, f2 As Long
    Dim h As String, c As Long, a As Long
    Set ws = HojaPref("4.12")
    Set wsT = Hoja("FMT_412")
    Set w10 = HojaPref("4.10")
    h = "'" & w10.Name & "'!"
    Encabezado ws, "4.12 SALDO VALORIZADO DE MATERIALES", 4, 14, 5, 7, 8, 9, 7, 9
    ws.Cells(11, 7).Value = "SALDO " & NombreMes(MesAnterior()) & " " & Year(CDate(MesAnterior()))
    ws.Cells(11, 9).Value = "INGRESOS " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    ws.Cells(11, 11).Value = "SALIDAS " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    ws.Cells(11, 13).Value = "SALDO " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    LimpiarDatos ws, F0
    r = F0
    f1 = r
    For g = 1 To rN
        CopiarDe wsT, 13, 13, ws, r
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = rNum(g)
        ws.Cells(r, 3).Value = rNom(g)
        r = r + 1
        For k = 1 To cN
            If cRub(k) = g Then
                a = fila410(k)
                CopiarDe wsT, 14, 14, ws, r
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 2).Value = cItem(k)
                ws.Cells(r, 3).Value = cDes(k)
                ws.Cells(r, 4).Value = cUnd(k)
                ws.Cells(r, 5).Value = cTipo(k)
                ws.Cells(r, 6).Value = cPU(k)
                ws.Cells(r, 7).Formula = "=N(" & h & Dir1(w10, a, 7) & ")"
                ws.Cells(r, 9).Formula = "=N(" & h & Dir1(w10, a, 39) & ")"
                ws.Cells(r, 11).Formula = "=N(" & h & Dir1(w10, a + 1, 39) & ")"
                ws.Cells(r, 13).Formula = "=G" & r & "+I" & r & "-K" & r
                ws.Cells(r, 8).Formula = "=ROUND($F" & r & "*G" & r & ",2)"
                ws.Cells(r, 10).Formula = "=ROUND($F" & r & "*I" & r & ",2)"
                ws.Cells(r, 12).Formula = "=ROUND($F" & r & "*K" & r & ",2)"
                ws.Cells(r, 14).Formula = "=ROUND($F" & r & "*M" & r & ",2)"
                r = r + 1
            End If
        Next k
    Next g
    f2 = r - 1
    CopiarDe wsT, 15, 20, ws, r
    ws.Cells(r, 2).Value = "Total"
    For c = 8 To 14 Step 2
        If f2 >= f1 Then
            ws.Cells(r, c).Formula = "=SUM(" & Dir1(ws, f1, c) & ":" & Dir1(ws, f2, c) & ")"
        Else
            ws.Cells(r, c).Value = 0
        End If
    Next c
    For c = 0 To 3
        ws.Cells(r + 2 + c, 12).Value = ws.Cells(11, 7 + 2 * c).Value
        ws.Cells(r + 2 + c, 14).Formula = "=" & Dir1(ws, r, 8 + 2 * c)
    Next c
    On Error Resume Next
    ws.PageSetup.PrintArea = "$B$1:$N$" & (r + 5)
    ws.PageSetup.PrintTitleRows = "$11:$12"
    On Error GoTo 0
End Sub

Public Function UnidadCorta(ByVal u As String) As String
    Dim s As String
    s = Replace(Normal(u), " ", "")
    If s = "HM" Or s = "H.M." Or s = "H.M" Or s = "HORAMAQUINA" Or s = "" Then
        UnidadCorta = "H.M."
    ElseIf s = "HE" Or s = "H.E." Or s = "H.E" Or s = "HORAEQUIPO" Then
        UnidadCorta = "H.E."
    ElseIf Left(s, 3) = "DIA" Or s = "D" Then
        UnidadCorta = "DIA"
    ElseIf Left(s, 3) = "VIA" Or Left(s, 3) = "VJE" Then
        UnidadCorta = "VJE"
    Else
        UnidadCorta = UCase(Trim(u))
    End If
End Function

Private Function FilaEquipo(ByVal cod As String) As Long
    Dim ws As Worksheet, f As Variant
    Set ws = Hoja("EQUIPOS")
    FilaEquipo = 0
    f = Application.Match(cod, ws.Range(ws.Cells(EQ0, EC_COD), ws.Cells(EQMAX, EC_COD)), 0)
    If Not IsError(f) Then FilaEquipo = EQ0 + CLng(f) - 1
End Function

Private Function DescCatalogoR(ByVal cod As String) As String
    ' nombre del combustible: del INV. (por ítem), del catálogo anterior o el mismo código
    Dim ws As Worksheet, f As Variant, k As Long
    DescCatalogoR = cod
    For k = 1 To cN
        If UCase(cItem(k)) = UCase(cod) Then
            DescCatalogoR = UCase(cDes(k))
            Exit Function
        End If
    Next k
    Set ws = Hoja("CATALOGO")
    f = Application.Match(cod, ws.Range(ws.Cells(CT0, CC_COD), ws.Cells(CTMAX, CC_COD)), 0)
    If Not IsError(f) Then DescCatalogoR = UCase(Txt(ws.Cells(CT0 + CLng(f) - 1, CC_DES).Value))
End Function

Private Sub Armar415()
    Dim ws As Worksheet, wsT As Worksheet, r As Long, k As Long, i As Long, d As Long, sec As Long, n As Long, f1 As Long
    Dim unid() As String, bloques As Collection, u As Variant, v() As Variant, hayEq As Boolean, secNom As String
    Dim combCod() As String, nComb As Long, x As Long, cK As Long, fe As Long, fuelF1 As Long, fuelF2 As Long, ent As String
    Dim usados() As Boolean, fmtH As String
    LeerEquipos
    Leer415
    Set ws = HojaPref("4.15")
    Set wsT = Hoja("FMT_415")
    ent = Dato(D_ENTIDAD)
    ws.Cells(1, 2).Value = ent
    ws.Cells(5, 2).Value = "CONTROL DE EQUIPO MECANICO Y COMBUSTIBLE CORRESPONDIENTE AL MES DE " & NombreMes(MesInicio()) & _
                           " " & Year(CDate(MesInicio()))
    ws.Cells(6, 6).Value = ConDos("“" & Dato(D_PROY) & "”")
    ws.Cells(8, 6).Value = ConDos(ent)
    ws.Cells(9, 6).Value = ConDos(Dato(D_MODAL))
    ws.Cells(10, 6).Value = ConDos(Dato(D_RESID))
    ws.Cells(11, 2).Value = Dato(D_F4)
    ws.Cells(11, 6).Value = ConDos(Dato(D_SUPER))
    LimpiarDatos ws, F0Q
    For d = 1 To 31
        ws.Columns(D1Q + d - 1).Hidden = (d > DiasMes())
    Next d
    ws.Range(ws.Columns(C_RAR), ws.Columns(C_RAT)).Hidden = True
    ReDim unid(1 To enEq + 1)
    For k = 1 To enEq
        unid(k) = UnidadCorta(eUnd(k))
    Next k
    r = F0Q
    fuelF1 = 0
    For sec = 1 To 2
        secNom = IIf(sec = 1, "MAQUINARIAS Y EQUIPOS ALQUILADOS POR SERVICIO (O/S)", "MAQUINARIAS Y EQUIPOS PROPIOS O PRESTADOS")
        Set bloques = New Collection
        hayEq = False
        For k = 1 To enEq
            If (sec = 1) = (eOS(k) <> "") Then
                hayEq = True
                On Error Resume Next
                bloques.Add unid(k), unid(k)
                On Error GoTo 0
            End If
        Next k
        If hayEq Then
            CopiarDe wsT, 14, 15, ws, r
            ws.Cells(r, 2).Value = secNom
            r = r + 2
            For Each u In bloques
                CopiarDe wsT, 16, 18, ws, r
                ws.Cells(r, 6).Value = "COMBUS./ " & u
                If u = "DIA" Then
                    ws.Cells(r, 8).Value = "TOTAL DIAS": ws.Cells(r, 9).Value = "DIAS ACUM."
                ElseIf u = "VJE" Then
                    ws.Cells(r, 8).Value = "TOTAL VIAJES": ws.Cells(r, 9).Value = "VIAJES ACUM."
                Else
                    ws.Cells(r, 8).Value = "TOTAL HRS": ws.Cells(r, 9).Value = "HORAS ACUM."
                End If
                ws.Cells(r, 10).Value = "MES DE " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
                ws.Cells(r, 41).Value = "TOTAL"
                ws.Cells(r + 1, 41).Value = "COMB."
                ws.Cells(r + 1, 42).Value = IIf(u = "DIA", "DIAS", u)
                CabeceraDias ws, r + 1, r + 2, D1Q
                r = r + 3
                f1 = r
                n = 0
                For k = 1 To enEq
                    If (sec = 1) = (eOS(k) <> "") And unid(k) = u Then
                        n = n + 1
                        CopiarDe wsT, 19, 20, ws, r
                        If fuelF1 = 0 Then fuelF1 = r
                        fuelF2 = r + 1
                        ws.Cells(r, 2).Value = n
                        ws.Cells(r, 3).Value = eDes(k) & IIf(eSerie(k) <> "", " – " & eSerie(k), "")
                        ws.Cells(r, 6).Value = "GAL"
                        ws.Cells(r + 1, 6).Value = IIf(u = "DIA", "DIA", u)
                        ReDim v(1 To 2, 1 To 31)
                        i = Buscar(qCol, UCase(eCod(k)))
                        If i > 0 Then
                            For d = 1 To 31
                                v(1, d) = qG(i, d)
                                v(2, d) = qH(i, d)
                            Next d
                        End If
                        ws.Range(ws.Cells(r, D1Q), ws.Cells(r + 1, D1Q + 30)).Value = v
                        ReDim usados(1 To 31)
                        For d = 1 To DiasMes()
                            usados(d) = (Num(v(1, d)) <> 0 Or Num(v(2, d)) <> 0)
                        Next d
                        ResumenUso eCod(k), usados
                        ws.Cells(r, 41).Formula = SumaFila(ws, r, D1Q, D1Q + 30)
                        ws.Cells(r + 1, 42).Formula = SumaFila(ws, r + 1, D1Q, D1Q + 30)
                        ' horas como tiempo [h]:mm (igual que el informe); días y viajes como números
                        If EsUnidadHora(eUnd(k)) Then fmtH = "[h]:mm" Else fmtH = "0"
                        ws.Range(ws.Cells(r + 1, D1Q), ws.Cells(r + 1, D1Q + 30)).NumberFormat = fmtH
                        ws.Cells(r + 1, 42).NumberFormat = fmtH
                        ws.Cells(r, 9).NumberFormat = fmtH
                        fe = FilaEquipo(eCod(k))
                        If fe > 0 And fmtH = "[h]:mm" Then
                            ws.Cells(r, 9).Formula = "=N(EQUIPOS!" & Hoja("EQUIPOS").Cells(fe, EC_ANT).Address & ")/24+" & Dir1(ws, r + 1, 42)
                        ElseIf fe > 0 Then
                            ws.Cells(r, 9).Formula = "=N(EQUIPOS!" & Hoja("EQUIPOS").Cells(fe, EC_ANT).Address & ")+" & Dir1(ws, r + 1, 42)
                        Else
                            ws.Cells(r, 9).Formula = "=" & Dir1(ws, r + 1, 42)
                        End If
                        ws.Cells(r, 43).Value = eOper(k)
                        ws.Cells(r, C_RAS).Value = eCod(k)
                        ws.Cells(r + 1, C_RAR).Value = eCod(k)
                        ws.Cells(r, C_RAT).Value = eComb(k)
                        r = r + 2
                    End If
                Next k
                CopiarDe wsT, 21, 22, ws, r
                ws.Cells(r, 31).Value = "TOTAL: COMBUSTIBLE/ " & IIf(u = "DIA", "DIAS", u)
                ws.Cells(r, 41).Formula = "=SUM(" & Dir1(ws, f1, 41) & ":" & Dir1(ws, r - 1, 41) & ")"
                ws.Cells(r, 42).Formula = "=SUM(" & Dir1(ws, f1, 42) & ":" & Dir1(ws, r - 1, 42) & ")"
                If EsUnidadHora(CStr(u)) Then ws.Cells(r, 42).NumberFormat = "[h]:mm" Else ws.Cells(r, 42).NumberFormat = "0"
                Domingos ws, D1Q, f1, r - 1
                r = r + 2
            Next u
        End If
    Next sec
    ' resumen de combustible por día y tipo (fórmulas)
    ReDim combCod(1 To enEq + 1)
    For k = 1 To enEq
        If eComb(k) <> "" Then
            cK = 0
            For x = 1 To nComb
                If UCase(combCod(x)) = UCase(eComb(k)) Then cK = x
            Next x
            If cK = 0 Then
                nComb = nComb + 1
                combCod(nComb) = eComb(k)
            End If
        End If
    Next k
    If nComb > 0 And fuelF1 > 0 Then
        CopiarDe wsT, 23, 24, ws, r
        CabeceraDias ws, r, r + 1, D1Q
        ws.Cells(r, 41).Value = "COMB."
        r = r + 2
        For x = 1 To nComb
            CopiarDe wsT, 25, 25, ws, r
            ws.Cells(r, 2).Value = DescCatalogoR(combCod(x))
            ws.Range(ws.Cells(r, D1Q), ws.Cells(r, D1Q + 30)).FormulaR1C1 = "=SUMIF(R" & fuelF1 & "C" & C_RAT & ":R" & fuelF2 & _
                "C" & C_RAT & ",""" & combCod(x) & """,R" & fuelF1 & "C:R" & fuelF2 & "C)"
            ws.Cells(r, 41).Formula = SumaFila(ws, r, D1Q, D1Q + 30)
            r = r + 1
        Next x
    End If
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$AP$" & IIf(r - 1 < F0Q, F0Q, r - 1)
    ws.PageSetup.PrintTitleRows = "$1:$12"
    On Error GoTo 0
End Sub


'========================================================================== DATOS PARA KARDEX, BINCARD Y PARTES
Public Sub LeerCuadrosAlmacen()
    ' llena las listas que usan el KARDEX y el BINCARD (ModKardex) a partir de INV. y 4.10
    Dim k As Long, i As Long, d As Long, m As Long, n As Long, q As Double, j As Long, p As Long
    Dim ped As String, obs As String, vale As String
    LeerInv
    Leer410
    LeerAuto
    Set aCol = New Collection
    anCat = cN
    anMov = 0
    anMal = 0
    n = cN
    If n < 1 Then n = 1
    ReDim aCod(1 To n): ReDim aDes(1 To n): ReDim aUnd(1 To n): ReDim aGru(1 To n)
    ReDim aPU(1 To n): ReDim aAnt(1 To n): ReDim aPrim(1 To n): ReDim aUlt(1 To n): ReDim aCntMov(1 To n)
    m = 0
    For k = 1 To cN
        aCod(k) = cItem(k)
        aDes(k) = cDes(k)
        aUnd(k) = cUnd(k)
        aGru(k) = rNom(cRub(k))
        aPU(k) = cPU(k)
        aCol.Add k, UCase(cItem(k))
        i = Buscar(tCol, UCase(cItem(k)))
        If i > 0 Then
            aAnt(k) = Num(tG(i))
            For d = 1 To 31
                If Num(tE(i, d)) <> 0 Then m = m + 1
                If Num(tS(i, d)) <> 0 Then m = m + 1
            Next d
        End If
    Next k
    If m < 1 Then m = 1
    ReDim vFec(1 To m): ReDim vTipo(1 To m): ReDim vDoc(1 To m): ReDim vNDoc(1 To m): ReDim vOC(1 To m)
    ReDim vProc(1 To m): ReDim vPart(1 To m): ReDim vPed(1 To m): ReDim vObs(1 To m)
    ReDim vCant(1 To m): ReDim vItem(1 To m): ReDim vFila(1 To m): ReDim vSig(1 To m): ReDim vOrden(1 To m)
    For k = 1 To cN
        i = Buscar(tCol, UCase(cItem(k)))
        If i > 0 Then
            For d = 1 To DiasMes()
                For p = 1 To 2
                    If p = 1 Then q = Num(tE(i, d)) Else q = Num(tS(i, d))
                    If q <> 0 Then
                        anMov = anMov + 1
                        vFec(anMov) = DiaFecha(d)
                        vItem(anMov) = k
                        vCant(anMov) = Abs(q)
                        If p = 1 Then
                            vTipo(anMov) = T_ING
                            vOC(anMov) = ""
                            If UCase(Left(cTipo(k), 3)) = "O/C" Or UCase(Left(cTipo(k), 3)) = "O/S" Then
                                vDoc(anMov) = Left(cTipo(k), 3)
                                vNDoc(anMov) = Trim(Mid(cTipo(k), 4))
                            Else
                                vDoc(anMov) = cTipo(k)
                                vNDoc(anMov) = cReq(k)
                            End If
                            vObs(anMov) = cProv(k)
                        Else
                            vTipo(anMov) = T_SAL
                            vDoc(anMov) = "SALIDA"
                            ped = ""
                            obs = ""
                            vale = ""
                            AutoSalida cItem(k), rNum(cRub(k)), DiaFecha(d), ped, obs, vale
                            vNDoc(anMov) = vale
                            vPed(anMov) = ped
                            vObs(anMov) = obs
                        End If
                        ' orden: por ítem, día, ingreso antes que salida
                        If aPrim(k) = 0 Then
                            aPrim(k) = anMov
                        Else
                            vSig(aUlt(k)) = anMov
                        End If
                        aUlt(k) = anMov
                        aCntMov(k) = aCntMov(k) + 1
                        vOrden(anMov) = anMov
                    End If
                Next p
            Next d
        End If
    Next k
End Sub

Public Sub LeerPartes415()
    ' llena las listas que usa ArmarPartes (ModPartes) con el combustible y las horas del 4.15
    Dim ws As Worksheet, last As Long, r As Long, d As Long, k As Long, n As Long, g As Double, h As Double
    Dim horo As Variant
    LeerEquipos
    LeerAuto
    pnPar = 0
    pnMal = 0
    Set ws = HojaPref("4.15")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    n = 0
    For r = F0Q To last
        If Txt(ws.Cells(r, C_RAS).Value) <> "" Then n = n + 31
    Next r
    LeerMovEq
    n = n + mqN
    If n < 1 Then n = 1
    ReDim pR(1 To n): ReDim pEq(1 To n): ReDim pFec(1 To n): ReDim pOrden(1 To n)
    ReDim pV(1 To n, 1 To PC_ULT + 2)            ' + 2: horas DE / A de cada línea (autollenado)
    For r = F0Q To last
        If Txt(ws.Cells(r, C_RAS).Value) <> "" Then
            k = Buscar(eCol, UCase(Txt(ws.Cells(r, C_RAS).Value)))
            If k > 0 And TieneMov(eCod(k)) Then
                PartesDesdeMov k                  ' horario y horómetro reales (MOV. EQUIPOS)
            ElseIf k > 0 Then
                horo = HorometroInicial(k)
                For d = 1 To DiasMes()
                    g = Num(ws.Cells(r, D1Q + d - 1).Value)
                    h = Num(ws.Cells(r + 1, D1Q + d - 1).Value)
                    If EsUnidadHora(eUnd(k)) Then h = Redondear(h * 24, 2)     ' [h]:mm -> horas
                    If (g <> 0 Or h <> 0) And auPartes Then
                        AutoParte k, DiaFecha(d), g, h, horo
                    ElseIf g <> 0 Or h <> 0 Then
                        pnPar = pnPar + 1
                        pR(pnPar) = pnPar
                        pEq(pnPar) = k
                        pFec(pnPar) = DiaFecha(d)
                        pOrden(pnPar) = pnPar
                        pV(pnPar, PC_FEC) = CDate(DiaFecha(d))
                        pV(pnPar, PC_EQ) = eCod(k)
                        pV(pnPar, PC_OPER) = eOper(k)
                        pV(pnPar, PC_TRAB) = ""
                        If h <> 0 Then pV(pnPar, PC_HREAL) = h Else pV(pnPar, PC_HREAL) = 0
                        pV(pnPar, PC_COMB) = g
                    End If
                Next d
            Else
                pnMal = pnMal + 1
            End If
        End If
    Next r
    OrdenarPartes
End Sub


'========================================================================== DETALLE SALIDAS
Public Sub ListarSalidas()
    ' una fila por salida (día e ítem) del 4.10, conservando el N° de vale, PEDIDO POR y destino ya escritos
    Dim wd As Worksheet, v As Variant, i As Long, n As Long, k As Long, d As Long, r As Long, t As Long, key As String
    Dim col As New Collection, kF() As Double, kI() As String, kVa() As Variant, kPe() As Variant, kDe() As Variant
    Dim usado() As Boolean, q As Double
    Set wd = HojaAuto(DS_HOJA)
    If wd Is Nothing Then Exit Sub
    v = wd.Range(wd.Cells(DS0, 1), wd.Cells(DS1, 9)).Value
    n = UBound(v, 1)
    ReDim kF(1 To n): ReDim kI(1 To n): ReDim kVa(1 To n): ReDim kPe(1 To n): ReDim kDe(1 To n): ReDim usado(1 To n)
    t = 0
    For i = 1 To n
        If Fecha(v(i, 2)) > 0 And Txt(v(i, 3)) <> "" Then
            key = Format(CDate(Fecha(v(i, 2))), "yyyymmdd") & "|" & UCase(Txt(v(i, 3)))
            If Buscar(col, key) = 0 Then
                t = t + 1
                col.Add t, key
                kF(t) = Fecha(v(i, 2))
                kI(t) = Txt(v(i, 3))
                kVa(t) = v(i, 7)
                kPe(t) = v(i, 8)
                kDe(t) = v(i, 9)
            End If
        End If
    Next i
    LeerInv
    Leer410
    wd.Range(wd.Cells(DS0, 2), wd.Cells(DS1, 3)).ClearContents
    wd.Range(wd.Cells(DS0, 6), wd.Cells(DS1, 9)).ClearContents
    r = DS0
    For d = 1 To DiasMes()
        For k = 1 To cN
            i = Buscar(tCol, UCase(cItem(k)))
            If i > 0 Then
                q = Num(tS(i, d))
                If q <> 0 And r <= DS1 Then
                    wd.Cells(r, 2).Value = CDate(DiaFecha(d))
                    wd.Cells(r, 3).NumberFormat = "@"
                    wd.Cells(r, 3).Value = cItem(k)
                    wd.Cells(r, 6).Value = q
                    i = Buscar(col, Format(CDate(DiaFecha(d)), "yyyymmdd") & "|" & UCase(cItem(k)))
                    If i > 0 Then
                        usado(i) = True
                        wd.Cells(r, 7).Value = kVa(i)
                        wd.Cells(r, 8).Value = kPe(i)
                        wd.Cells(r, 9).Value = kDe(i)
                    End If
                    r = r + 1
                End If
            End If
        Next k
    Next d
    ' lo escrito para salidas que ya no están en el 4.10 queda al final (la REVISIÓN lo avisa)
    For i = 1 To t
        If Not usado(i) And r <= DS1 Then
            If Txt(kVa(i)) <> "" Or Txt(kPe(i)) <> "" Or Txt(kDe(i)) <> "" Then
                wd.Cells(r, 2).Value = CDate(kF(i))
                wd.Cells(r, 3).NumberFormat = "@"
                wd.Cells(r, 3).Value = kI(i)
                wd.Cells(r, 7).Value = kVa(i)
                wd.Cells(r, 8).Value = kPe(i)
                wd.Cells(r, 9).Value = kDe(i)
                r = r + 1
            End If
        End If
    Next i
End Sub


'========================================================================== CIERRE DEL MES Y NUEVA OBRA
Public Sub CerrarCuadros()
    ' saldo del mes -> saldo anterior (4.10); horas del mes -> acumulado anterior (EQUIPOS); vacía los días
    Dim ws As Worksheet, last As Long, r As Long, s As Double, fe As Long, we As Worksheet
    Set ws = HojaPref("4.10")
    Application.Calculate
    last = UltimaFila(ws, 6, F0, 8000)
    For r = F0 To last
        If Txt(ws.Cells(r, 6).Value) = "ENTRADA" Then
            s = Num(ws.Cells(r, 7).Value) + Application.WorksheetFunction.Sum(ws.Range(ws.Cells(r, D1), ws.Cells(r, D1 + 30))) - _
                Application.WorksheetFunction.Sum(ws.Range(ws.Cells(r + 1, D1), ws.Cells(r + 1, D1 + 30)))
            ws.Cells(r, 7).Value = Redondear(s, 4)
            ws.Range(ws.Cells(r, D1), ws.Cells(r + 1, D1 + 30)).ClearContents
        End If
    Next r
    Set ws = HojaPref("4.15")
    Set we = Hoja("EQUIPOS")
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    For r = F0Q To last
        If Txt(ws.Cells(r, C_RAS).Value) <> "" Then
            fe = FilaEquipo(Txt(ws.Cells(r, C_RAS).Value))
            If fe > 0 Then
                s = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(r + 1, D1Q), ws.Cells(r + 1, D1Q + 30)))
                If EsUnidadHora(Txt(we.Cells(fe, EC_UND).Value)) Then s = s * 24
                we.Cells(fe, EC_ANT).Value = Redondear(Num(we.Cells(fe, EC_ANT).Value) + s, 4)
            End If
            ws.Range(ws.Cells(r, D1Q), ws.Cells(r + 1, D1Q + 30)).ClearContents
        End If
    Next r
End Sub

Public Sub VaciarCuadros()
    ' NUEVA OBRA: inventario y cuadros vacíos
    Dim ws As Worksheet, last As Long
    Set ws = HojaPref("INV.")
    last = UltimaFilaVarias(ws, 1, 10, 2, 6000)
    If last >= 2 Then ws.Range("A2:J" & last).ClearContents
    On Error Resume Next
    ws.Range("G2:G" & Application.WorksheetFunction.Max(last, 301)).Formula = "=IF(F2="""","""",E2*F2)"
    On Error GoTo 0
    LimpiarDatos HojaPref("4.10"), F0
    LimpiarDatos HojaPref("4.11"), F0
    LimpiarDatos HojaPref("4.12"), F0
    LimpiarDatos HojaPref("4.15"), F0Q
End Sub


'========================================================================== TRAER LOS DATOS DEL FORMATO ANTERIOR
Public Sub TraerDatosAnteriores()
    ' pasa CATALOGO + MOVIMIENTOS (y PARTES) del formato anterior al INV., al 4.10 y al 4.15
    Dim cm As Long, ws As Worksheet, w10 As Worksheet, g As Long, k As Long, r As Long, m As Long, d As Long
    Dim grupos As New Collection, gNom() As String, nG As Long, cnt() As Long, itNum() As String, i As Long, j As Long
    Dim ing As Double, fec As Double, tipo As String, prov As String, last As Long, v() As Variant, lo As Object
    Dim nItems As Long, nEq As Long, x As Long, hh() As Double, cc() As Double, w15 As Worksheet
    If Not Pregunta("Se copian los datos del formato anterior (hojas CATALOGO, MOVIMIENTOS y PARTES) al INV. 2026, " & _
                    "al 4.10 y al 4.15." & vbCrLf & vbCrLf & "Lo que hoy está en el INV. y en los cuadros se reemplaza. ¿Continuar?", _
                    "TRAER DATOS ANTERIORES", False) Then Exit Sub
    Inicio cm, "Pasando los datos al INV. y a los cuadros..."
    On Error GoTo Falla
    LeerAlmacenCatalogo
    Set ws = HojaPref("INV.")
    last = UltimaFilaVarias(ws, 1, 10, 2, 6000)
    If last >= 2 Then ws.Range("A2:J" & last).ClearContents
    ReDim gNom(1 To anCat + 1): ReDim cnt(1 To anCat + 1): ReDim itNum(1 To anCat + 1)
    For k = 1 To anCat
        g = Buscar(grupos, Normal(IIf(aGru(k) = "", "OTROS", aGru(k))))
        If g = 0 Then
            nG = nG + 1
            gNom(nG) = UCase(IIf(aGru(k) = "", "OTROS", aGru(k)))
            grupos.Add nG, Normal(gNom(nG))
        End If
    Next k
    r = 2
    For g = 1 To nG
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = CStr(g)
        ws.Cells(r, 2).Value = gNom(g)
        r = r + 1
        For k = 1 To anCat
            If Normal(IIf(aGru(k) = "", "OTROS", aGru(k))) = Normal(gNom(g)) Then
                cnt(g) = cnt(g) + 1
                itNum(k) = g & "." & cnt(g)
                ing = 0: fec = 0: tipo = "": prov = ""
                m = aPrim(k)
                Do While m > 0
                    If vTipo(m) = T_ING Then
                        ing = ing + vCant(m)
                        If fec = 0 Then
                            fec = vFec(m)
                            If vOC(m) <> "" Then tipo = Doc(vOC(m)) Else tipo = vDoc(m)
                            prov = vProc(m)
                        End If
                    End If
                    m = vSig(m)
                Loop
                ws.Cells(r, 1).NumberFormat = "@"
                ws.Cells(r, 1).Value = itNum(k)
                ws.Cells(r, 2).Value = aDes(k)
                ws.Cells(r, 3).Value = aUnd(k)
                ws.Cells(r, 4).Value = tipo
                ws.Cells(r, 5).Value = aPU(k)
                If ing <> 0 Then ws.Cells(r, 6).Value = ing
                ws.Cells(r, 7).Formula = "=IF(F" & r & "="""","""",E" & r & "*F" & r & ")"
                If fec > 0 Then ws.Cells(r, 8).Value = CDate(fec)
                ws.Cells(r, 10).Value = prov
                nItems = nItems + 1
                r = r + 1
            End If
        Next k
    Next g
    On Error Resume Next
    Set lo = ws.ListObjects(1)
    If Not lo Is Nothing Then
        If lo.Range.Rows.Count < r - 1 Then lo.Resize ws.Range("A1:J" & (r - 1))
    End If
    On Error GoTo Falla
    ' arma los cuadros y luego copia las cantidades por día
    Fin cm
    ActualizarCuadros True
    Inicio cm, "Copiando las cantidades por día..."
    Set w10 = HojaPref("4.10")
    last = UltimaFila(w10, 6, F0, 8000)
    For r = F0 To last
        If Txt(w10.Cells(r, 6).Value) = "ENTRADA" Then
            k = 0
            For j = 1 To anCat
                If itNum(j) = Txt(w10.Cells(r, 2).Value) Then k = j: Exit For
            Next j
            If k > 0 Then
                ReDim v(1 To 2, 1 To 31)
                m = aPrim(k)
                Do While m > 0
                    If vFec(m) >= MesInicio() And vFec(m) <= MesFin() Then
                        d = Day(CDate(vFec(m)))
                        If vTipo(m) = T_ING Then
                            v(1, d) = Num(v(1, d)) + vCant(m)
                        Else
                            v(2, d) = Num(v(2, d)) + vCant(m)
                        End If
                    End If
                    m = vSig(m)
                Loop
                w10.Cells(r, 7).Value = Redondear(aAnt(k), 4)
                w10.Range(w10.Cells(r, D1), w10.Cells(r + 1, D1 + 30)).Value = v
            End If
        End If
    Next r
    TraerSalidasAnteriores itNum                  ' N° de documento, PEDIDO POR y destino -> DETALLE SALIDAS
    ' equipos: horas y combustible de la hoja PARTES
    LeerPartesHoja
    TraerMovAnterior                              ' partes -> MOV. EQUIPOS (cuadro de movimiento diario)
    Set w15 = HojaPref("4.15")
    last = w15.UsedRange.Row + w15.UsedRange.Rows.Count
    For r = F0Q To last
        If Txt(w15.Cells(r, C_RAS).Value) <> "" Then
            k = Buscar(eCol, UCase(Txt(w15.Cells(r, C_RAS).Value)))
            If k > 0 Then
                ReDim v(1 To 2, 1 To 31)
                For x = 1 To pnPar
                    If pEq(x) = k And pFec(x) >= MesInicio() And pFec(x) <= MesFin() Then
                        d = Day(CDate(pFec(x)))
                        v(1, d) = Num(v(1, d)) + Num(pV(pR(x), PC_COMB))
                        v(2, d) = Num(v(2, d)) + HorasFila(pR(x))
                    End If
                Next x
                For d = 1 To 31
                    If Num(v(1, d)) = 0 Then v(1, d) = Empty
                    If Num(v(2, d)) = 0 Then
                        v(2, d) = Empty
                    ElseIf EsUnidadHora(eUnd(k)) Then
                        v(2, d) = Num(v(2, d)) / 24
                    End If
                Next d
                w15.Range(w15.Cells(r, D1Q), w15.Cells(r + 1, D1Q + 30)).Value = v
                nEq = nEq + 1
            End If
        End If
    Next r
    Fin cm
    ActualizarCuadros True
    Inicio cm, "Terminando..."
    Fin cm
    Aviso "Listo: " & nItems & " ítems pasaron al INV. 2026 (con su saldo anterior y las entradas y salidas de cada día en el " & _
          "4.10) y " & nEq & " equipo(s) al 4.15. Los destinos de las salidas pasaron a DETALLE SALIDAS " & _
          "y los partes a MOV. EQUIPOS: complete ahí lo que falte (horario, horómetro, operador) y pulse ACTUALIZAR Y REVISAR." & vbCrLf & vbCrLf & "Desde ahora registre directamente en el INV., el 4.10 " & _
          "y el 4.15.", "TRAER DATOS ANTERIORES"
    HojaPref("4.10").Activate
    Exit Sub
Falla:
    Fin cm
    Aviso "No se pudieron traer los datos: " & Err.Description, "TRAER DATOS ANTERIORES", 16
End Sub


'========================================================================== EXPORTAR EL INFORME (Excel + PDF, «Guardar como»)
Private Function NombreArchivoInforme(ByVal tema As String) As String
    Dim n As String
    n = Dato(D_NINF)
    If n <> "" Then n = n & ". "
    NombreArchivoInforme = n & tema & "_" & Format(CDate(FechaCorte()), "dd") & " DE " & NombreMes(MesInicio()) & " " & _
                           Year(CDate(MesInicio())) & ".xlsx"
End Function

Private Function GuardarLibro(ByVal wb As Workbook, ByVal tema As String, ByRef rutaPdf As String, _
                              ByVal preguntarPdf As Boolean) As String
    Dim ruta As String, conPdf As Boolean
    GuardarLibro = ""
    rutaPdf = ""
    ruta = GuardarComo(RutaSugerida("INFORME", NombreArchivoInforme(tema)), "Guardar el " & tema & " como", "XLSX")
    If ruta = "" Then Exit Function
    If LCase(Right(ruta, 5)) <> ".xlsx" Then ruta = ruta & ".xlsx"
    If Dir(ruta) <> "" And gCarpetaFija = "" And Not gSilencio Then
        If Not Pregunta("Ya existe:" & vbCrLf & ruta & vbCrLf & vbCrLf & "¿Reemplazarlo?", "GUARDAR", True) Then Exit Function
    End If
    On Error GoTo Falla
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=ruta, FileFormat:=51
    Application.DisplayAlerts = True
    GuardarLibro = ruta
    conPdf = True
    If preguntarPdf Then conPdf = Pregunta("Se guardó el Excel. ¿Exportar también a PDF (mismo nombre y carpeta)?", tema, True)
    If conPdf And Not gSinPDF Then
        rutaPdf = Left(ruta, Len(ruta) - 5) & ".pdf"
        wb.ExportAsFixedFormat Type:=0, Filename:=rutaPdf, Quality:=0, IncludeDocProperties:=False, _
                               IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If
    Exit Function
Falla:
    Application.DisplayAlerts = True
    Aviso "No se pudo guardar en:" & vbCrLf & ruta & vbCrLf & "(" & Err.Description & ")" & vbCrLf & vbCrLf & _
          "¿Hay un archivo con ese nombre abierto en Excel?", tema, 48
End Function

Private Function Exportar(ByVal hojas As Variant, ByVal tema As String, ByVal preguntarPdf As Boolean) As String
    Dim wb As Workbook, ruta As String, rutaPdf As String, msg As String, l As Variant, wsX As Worksheet
    Exportar = ""
    ThisWorkbook.Worksheets(hojas).Copy
    Set wb = ActiveWorkbook
    On Error Resume Next
    l = wb.LinkSources(1)
    If Not IsEmpty(l) Then wb.BreakLink Name:=ThisWorkbook.FullName, Type:=1
    On Error GoTo 0
    For Each wsX In wb.Worksheets
        AjustarInforme wsX
    Next wsX
    ruta = GuardarLibro(wb, tema, rutaPdf, preguntarPdf)
    If ruta = "" Then
        msg = tema & ": el libro quedó abierto sin guardar (se canceló «Guardar como»)."
    Else
        msg = tema & " de " & MesTexto() & vbCrLf & "Excel: " & ruta
        If rutaPdf <> "" Then msg = msg & vbCrLf & "PDF: " & rutaPdf
        Mensaje D_MSGINF, tema & " (" & Format(Now, "dd/mm/yyyy hh:mm") & "): " & ruta
    End If
    Exportar = ruta
    Aviso msg, tema
End Function

Public Sub ExportarInformeAlmacenBoton()
    Dim s As String
    s = ExportarInformeAlmacen(True)
End Sub

Public Function ExportarInformeAlmacen(ByVal preguntarPdf As Boolean) As String
    Dim s As String
    s = ActualizarCuadros(True)
    ExportarInformeAlmacen = Exportar(Array(HojaPref("INV.").Name, HojaPref("4.10").Name, HojaPref("4.11").Name, _
                                            HojaPref("4.12").Name), "ALMACEN DE OBRA", preguntarPdf)
End Function

Public Sub ExportarInformeEquiposBoton()
    Dim s As String
    s = ExportarInformeEquipos(True)
End Sub

Public Function ExportarInformeEquipos(ByVal preguntarPdf As Boolean) As String
    Dim s As String
    s = ActualizarCuadros(True)
    ExportarInformeEquipos = Exportar(Array(HojaPref("4.15").Name), "CONTROL DE EQUIPO MECANICO", preguntarPdf)
End Function
