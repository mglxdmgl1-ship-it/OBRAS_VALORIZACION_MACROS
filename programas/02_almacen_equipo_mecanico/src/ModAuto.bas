Attribute VB_Name = "ModAuto"
Option Explicit

'==========================================================================================
'  AUTOLLENADO (opcional) de los espacios que el 4.10 y el 4.15 no tienen:
'   - parte diario: operador, horario de mañana y tarde, horas DE / A, horómetro y trabajo realizado
'   - kardex: PEDIDO POR y destino (observación) de cada salida
'  Se configura en las hojas AUTOLLENADO (opciones, equipos, rubros y personal de la planilla)
'  y DETALLE DIARIO (trabajos de cada equipo por día y destino de cada salida).
'==========================================================================================

Public Const AU_HOJA As String = "AUTOLLENADO"
Public Const DD_HOJA As String = "DETALLE DIARIO"
' ---- opciones (columna D)
Public Const AU_PARTES As String = "D5"
Public Const AU_M1 As String = "D6"
Public Const AU_M2 As String = "D7"
Public Const AU_T1 As String = "D8"
Public Const AU_T2 As String = "D9"
Public Const AU_TRAB As String = "D10"
Public Const AU_PMODO As String = "D11"
Public Const AU_PNOM As String = "D12"
Public Const AU_POBS As String = "D13"
' ---- tablas de AUTOLLENADO
Public Const AE0 As Long = 17                 ' equipos: filas 17..36 (B código ... J trabajo)
Public Const AE1 As Long = 36
Public Const AR0 As Long = 40                 ' rubros: filas 40..59 (B rubro, C descripción, D pedido por, E observación)
Public Const AR1 As Long = 59
Public Const AP0 As Long = 63                 ' personal: filas 63..122 (C nombre, D cargo, E pide material, K..AO días)
Public Const AP1 As Long = 122
Public Const AP_DIA1 As Long = 11             ' columna K = día 1 del tareo
' ---- tablas de DETALLE DIARIO
Public Const DD0 As Long = 6                  ' primera fila de datos
Public Const DD1T As Long = 505               ' trabajos: B fecha, C equipo, D trabajo, E horas
Public Const DD1S As Long = 1505              ' salidas: H fecha, I ítem, J descripción, K destino, L pedido por

' ---- lo leído
Public auHay As Boolean                       ' existe la hoja AUTOLLENADO
Public auPartes As Boolean
Public auM1 As Double, auM2 As Double, auT1 As Double, auT2 As Double
Public auTrab As String, auModo As String, auNom As String, auObs As String
Private eqCol As Collection
Private eqOper() As String, eqM1() As Double, eqM2() As Double, eqT1() As Double, eqT2() As Double
Private eqHoro() As Variant, eqTrab() As String, eqFila() As Long
Private rbCol As Collection
Private rbNom() As String, rbObs() As String
Private psN As Long
Private psNom() As String, psPide() As Boolean, psTar() As String, psSinTareo() As Boolean
Private trCol As Collection                   ' "aaaammdd|COD" -> lista "i1,i2,..."
Private trTxt() As String, trH() As Double, trN As Long
Private saCol As Collection                   ' "aaaammdd|ITEM" -> lista
Private saDest() As String, saPed() As String, saN As Long


'========================================================================== UTILIDADES
Public Function HojaAuto(ByVal nombre As String) As Worksheet
    Set HojaAuto = Nothing
    On Error Resume Next
    Set HojaAuto = ThisWorkbook.Worksheets(nombre)
End Function

Private Function HoraDe(ByVal v As Variant, ByVal def As Double) As Double
    ' hora como fracción del día (acepta 07:30, 7:30 o 0.3125)
    Dim s As String
    HoraDe = def
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim(Replace(CStr(v), ".", ":"))
        If s = "" Then Exit Function
        On Error Resume Next
        HoraDe = CDbl(TimeValue(s))
        If Err.Number <> 0 Then HoraDe = def
        On Error GoTo 0
    ElseIf EsNum(v) Then
        HoraDe = Num(v) - Int(Num(v))
        If HoraDe = 0 And Num(v) <> 0 Then HoraDe = def
    End If
End Function

Private Function Minuto(ByVal t As Double) As Double
    Minuto = Round(t * 1440, 0) / 1440
End Function

Private Function ClaveDia(ByVal f As Double) As String
    ClaveDia = Format(CDate(f), "yyyymmdd")
End Function

Private Sub Agregar(ByVal c As Collection, ByVal clave As String, ByVal i As Long)
    ' guarda varias filas por clave: "i1,i2,..."
    Dim s As String
    On Error Resume Next
    s = c(clave)
    If Err.Number = 0 Then
        c.Remove clave
        c.Add s & "," & i, clave
    Else
        Err.Clear
        c.Add CStr(i), clave
    End If
End Sub

Private Function Lista(ByVal c As Collection, ByVal clave As String) As String
    Lista = ""
    On Error Resume Next
    Lista = c(clave)
End Function

Private Function Reloj(ByVal t0 As Double, ByVal horas As Double, ByVal m2 As Double, ByVal t1 As Double) As Double
    ' hora de término trabajando "horas" desde t0 (salta el refrigerio entre m2 y t1)
    Dim resto As Double, t As Double
    resto = horas / 24
    t = t0
    If t < m2 Then
        If resto <= m2 - t + 0.0000001 Then
            Reloj = Minuto(t + resto)
            Exit Function
        End If
        resto = resto - (m2 - t)
        t = t1
    ElseIf t < t1 Then
        t = t1
    End If
    Reloj = Minuto(t + resto)
End Function


'========================================================================== LEER LA CONFIGURACIÓN
Public Sub LeerAuto()
    Dim ws As Worksheet, wd As Worksheet, v As Variant, i As Long, n As Long, c As String, d As Long, s As String
    auHay = False
    auPartes = False
    auModo = "NO"
    Set eqCol = New Collection
    Set rbCol = New Collection
    Set trCol = New Collection
    Set saCol = New Collection
    psN = 0: trN = 0: saN = 0
    Set ws = HojaAuto(AU_HOJA)
    If ws Is Nothing Then Exit Sub
    auHay = True
    auPartes = (UCase(Left(Txt(ws.Range(AU_PARTES).Value), 1)) = "S")
    auM1 = HoraDe(ws.Range(AU_M1).Value, TimeSerial(7, 0, 0))
    auM2 = HoraDe(ws.Range(AU_M2).Value, TimeSerial(12, 0, 0))
    auT1 = HoraDe(ws.Range(AU_T1).Value, TimeSerial(13, 0, 0))
    auT2 = HoraDe(ws.Range(AU_T2).Value, TimeSerial(16, 0, 0))
    auTrab = Limpio(ws.Range(AU_TRAB).Value)
    s = UCase(Normal(ws.Range(AU_PMODO).Value))
    If InStr(s, "DIA") > 0 Then
        auModo = "DIA"
    ElseIf InStr(s, "RUBRO") > 0 Then
        auModo = "RUBRO"
    ElseIf InStr(s, "PERSONA") > 0 Then
        auModo = "UNA"
    Else
        auModo = "NO"
    End If
    auNom = Limpio(ws.Range(AU_PNOM).Value)
    auObs = Limpio(ws.Range(AU_POBS).Value)

    ' equipos
    n = AE1 - AE0 + 1
    ReDim eqOper(1 To n): ReDim eqM1(1 To n): ReDim eqM2(1 To n): ReDim eqT1(1 To n): ReDim eqT2(1 To n)
    ReDim eqHoro(1 To n): ReDim eqTrab(1 To n): ReDim eqFila(1 To n)
    v = ws.Range(ws.Cells(AE0, 1), ws.Cells(AE1, 10)).Value
    For i = 1 To n
        c = UCase(Txt(v(i, 2)))
        If c <> "" Then
            If Buscar(eqCol, c) = 0 Then
                eqCol.Add i, c
                eqFila(i) = AE0 + i - 1
                eqOper(i) = Limpio(v(i, 4))
                eqM1(i) = HoraDe(v(i, 5), auM1)
                eqM2(i) = HoraDe(v(i, 6), auM2)
                eqT1(i) = HoraDe(v(i, 7), auT1)
                eqT2(i) = HoraDe(v(i, 8), auT2)
                If EsNum(v(i, 9)) Then eqHoro(i) = Num(v(i, 9)) Else eqHoro(i) = Empty
                eqTrab(i) = Limpio(v(i, 10))
            End If
        End If
    Next i
    ' rubros
    n = AR1 - AR0 + 1
    ReDim rbNom(1 To n): ReDim rbObs(1 To n)
    v = ws.Range(ws.Cells(AR0, 1), ws.Cells(AR1, 5)).Value
    For i = 1 To n
        c = UCase(Txt(v(i, 2)))
        If c <> "" Then
            If Buscar(rbCol, c) = 0 Then
                rbCol.Add i, c
                rbNom(i) = Limpio(v(i, 4))
                rbObs(i) = Limpio(v(i, 5))
            End If
        End If
    Next i
    ' personal y tareo
    n = AP1 - AP0 + 1
    ReDim psNom(1 To n): ReDim psPide(1 To n): ReDim psTar(1 To n, 1 To 31): ReDim psSinTareo(1 To n)
    v = ws.Range(ws.Cells(AP0, 1), ws.Cells(AP1, AP_DIA1 + 30)).Value
    For i = 1 To n
        If Limpio(v(i, 3)) <> "" Then
            psN = psN + 1
            psNom(psN) = Limpio(v(i, 3))
            psPide(psN) = (UCase(Left(Txt(v(i, 5)), 1)) = "S")
            psSinTareo(psN) = True
            For d = 1 To 31
                psTar(psN, d) = UCase(Txt(v(i, AP_DIA1 + d - 1)))
                If psTar(psN, d) <> "" Then psSinTareo(psN) = False
            Next d
        End If
    Next i

    ' detalle diario
    Set wd = HojaAuto(DD_HOJA)
    If wd Is Nothing Then Exit Sub
    n = DD1T - DD0 + 1
    ReDim trTxt(1 To n): ReDim trH(1 To n)
    v = wd.Range(wd.Cells(DD0, 1), wd.Cells(DD1T, 5)).Value
    For i = 1 To n
        If Fecha(v(i, 2)) > 0 And Txt(v(i, 3)) <> "" Then
            trN = trN + 1
            trTxt(trN) = Limpio(v(i, 4))
            trH(trN) = Num(v(i, 5))
            Agregar trCol, ClaveDia(Fecha(v(i, 2))) & "|" & UCase(Txt(v(i, 3))), trN
        End If
    Next i
    n = DD1S - DD0 + 1
    ReDim saDest(1 To n): ReDim saPed(1 To n)
    v = wd.Range(wd.Cells(DD0, 8), wd.Cells(DD1S, 12)).Value
    For i = 1 To n
        If Fecha(v(i, 1)) > 0 Then
            saN = saN + 1
            saDest(saN) = Limpio(v(i, 4))
            saPed(saN) = Limpio(v(i, 5))
            Agregar saCol, ClaveDia(Fecha(v(i, 1))) & "|" & UCase(Txt(v(i, 2))), saN
        End If
    Next i
End Sub


'========================================================================== KARDEX: PEDIDO POR Y DESTINO
Public Sub AutoSalida(ByVal item As String, ByVal rubro As String, ByVal f As Double, ByRef ped As String, ByRef obs As String)
    ' completa el PEDIDO POR y la observación (destino) de una salida
    Dim l As String, p As Variant, j As Long, k As Long, d As Long, pres() As Long, np As Long
    If Not auHay Then Exit Sub
    ' 1) DETALLE DIARIO: fecha + ítem, o fecha con el ítem vacío (vale para todas las salidas del día)
    l = Lista(saCol, ClaveDia(f) & "|" & UCase(item))
    If l = "" Then l = Lista(saCol, ClaveDia(f) & "|")
    If l <> "" Then
        For Each p In Split(l, ",")
            j = CLng(p)
            If saDest(j) <> "" Then
                If InStr(1, obs, saDest(j), vbTextCompare) = 0 Then
                    If obs = "" Then obs = saDest(j) Else obs = obs & "; " & saDest(j)
                End If
            End If
            If ped = "" And saPed(j) <> "" Then ped = saPed(j)
        Next p
    End If
    ' 2) observación del rubro o la general
    k = Buscar(rbCol, UCase(rubro))
    If obs = "" Then
        If k > 0 Then obs = rbObs(k)
        If obs = "" Then obs = auObs
    End If
    ' 3) PEDIDO POR según la opción elegida
    If ped <> "" Or auModo = "NO" Then Exit Sub
    Select Case auModo
        Case "UNA"
            ped = auNom
        Case "RUBRO"
            If k > 0 Then ped = rbNom(k)
            If ped = "" Then ped = auNom
        Case "DIA"
            ' rota entre el personal que pide material y asistió ese día (tareo con A, o sin tareo)
            d = Day(CDate(f))
            ReDim pres(1 To psN + 1)
            np = 0
            For j = 1 To psN
                If psPide(j) Then
                    If psSinTareo(j) Or psTar(j, d) = "A" Then
                        np = np + 1
                        pres(np) = j
                    End If
                End If
            Next j
            If np > 0 Then
                ped = psNom(pres(((d - 1) Mod np) + 1))
            Else
                ped = auNom
            End If
    End Select
End Sub


'========================================================================== PARTE DIARIO
Public Function AutoLineas(ByVal cod As String, ByVal f As Double) As Long
    ' cuántas líneas de trabajo tiene el equipo ese día en DETALLE DIARIO (para dimensionar)
    Dim l As String
    AutoLineas = 1
    If Not auHay Then Exit Function
    l = Lista(trCol, ClaveDia(f) & "|" & UCase(cod))
    If l <> "" Then AutoLineas = UBound(Split(l, ",")) + 1
End Function

Public Function TotalLineasAuto() As Long
    TotalLineasAuto = trN
End Function

Public Sub AutoParte(ByVal k As Long, ByVal f As Double, ByVal g As Double, ByVal h As Double, ByRef horo As Variant)
    ' agrega a pV las líneas del parte de un equipo y un día, con horario, horas DE / A, horómetro y trabajo
    Dim e As Long, m1 As Double, m2 As Double, t1 As Double, t2 As Double, oper As String, l As String
    Dim idx() As String, nl As Long, i As Long, hi() As Double, tx() As String, tot As Double, acum As Double
    Dim reloj0 As Double, fin As Double, todos As Boolean, esDia As Boolean, hh As Double

    e = Buscar(eqCol, UCase(eCod(k)))
    If e > 0 Then
        m1 = eqM1(e): m2 = eqM2(e): t1 = eqT1(e): t2 = eqT2(e)
        oper = eqOper(e)
    Else
        m1 = auM1: m2 = auM2: t1 = auT1: t2 = auT2
    End If
    If oper = "" Then oper = eOper(k)
    esDia = (InStr(1, Normal(eUnd(k)), "DIA", vbTextCompare) > 0)

    ' líneas de trabajo del día
    l = Lista(trCol, ClaveDia(f) & "|" & UCase(eCod(k)))
    If l = "" Then
        nl = 1
        ReDim hi(1 To 1): ReDim tx(1 To 1)
        hi(1) = h
        If e > 0 Then tx(1) = eqTrab(e)
        If tx(1) = "" Then tx(1) = auTrab
    Else
        idx = Split(l, ",")
        nl = UBound(idx) + 1
        ReDim hi(1 To nl): ReDim tx(1 To nl)
        todos = True
        tot = 0
        For i = 1 To nl
            tx(i) = trTxt(CLng(idx(i - 1)))
            If trH(CLng(idx(i - 1))) <= 0 Then todos = False
            tot = tot + trH(CLng(idx(i - 1)))
        Next i
        ' reparte las horas del 4.15: según las horas de cada línea o en partes iguales
        acum = 0
        For i = 1 To nl
            If i = nl Then
                hi(i) = Redondear(h - acum, 2)
            ElseIf todos And tot > 0 Then
                hi(i) = Redondear(h * trH(CLng(idx(i - 1))) / tot, 2)
            Else
                hi(i) = Redondear(h / nl, 2)
            End If
            acum = acum + hi(i)
        Next i
    End If

    reloj0 = m1
    For i = 1 To nl
        pnPar = pnPar + 1
        pR(pnPar) = pnPar
        pEq(pnPar) = k
        pFec(pnPar) = f
        pOrden(pnPar) = pnPar
        pV(pnPar, PC_FEC) = CDate(f)
        pV(pnPar, PC_EQ) = eCod(k)
        pV(pnPar, PC_OPER) = oper
        pV(pnPar, PC_TRAB) = tx(i)
        pV(pnPar, PC_HREAL) = hi(i)
        If i = 1 Then pV(pnPar, PC_COMB) = g Else pV(pnPar, PC_COMB) = 0
        If h > 0 Then
            ' horas por línea (DE / A), una tras otra desde el inicio de la jornada
            If esDia Then
                hh = ((m2 - m1) + (t2 - t1)) * 24 * hi(i) / IIf(h = 0, 1, h)
            Else
                hh = hi(i)
            End If
            fin = Reloj(reloj0, hh, m2, t1)
            pV(pnPar, PC_ULT + 1) = reloj0
            pV(pnPar, PC_ULT + 2) = fin
            reloj0 = fin
            If reloj0 >= m2 - 0.0000001 And reloj0 < t1 Then reloj0 = t1
            ' horario de mañana y tarde del parte (en la primera línea)
            If i = 1 Then
                If esDia Then
                    fin = Reloj(m1, ((m2 - m1) + (t2 - t1)) * 24, m2, t1)
                Else
                    fin = Reloj(m1, h, m2, t1)
                End If
                pV(pnPar, PC_MDE) = m1
                If fin <= m2 + 0.0000001 Then
                    pV(pnPar, PC_MA) = fin
                Else
                    pV(pnPar, PC_MA) = m2
                    pV(pnPar, PC_TDE) = t1
                    pV(pnPar, PC_TA) = fin
                End If
            End If
        End If
    Next i
    ' horómetro: inicial = lectura acumulada; final = inicial + horas del día
    If Not IsEmpty(horo) And Not esDia Then
        pV(pnPar - nl + 1, PC_HINI) = Redondear(Num(horo), 2)
        pV(pnPar, PC_HFIN) = Redondear(Num(horo) + h, 2)
        horo = Num(horo) + h
    End If
End Sub

Public Function HorometroInicial(ByVal k As Long) As Variant
    Dim e As Long
    HorometroInicial = Empty
    If Not auHay Then Exit Function
    e = Buscar(eqCol, UCase(eCod(k)))
    If e > 0 Then HorometroInicial = eqHoro(e)
End Function


'========================================================================== SINCRONIZAR LISTAS (ACTUALIZAR CUADROS)
Public Sub SincronizarAuto()
    ' pone en AUTOLLENADO los equipos de EQUIPOS y los rubros del INV. que falten (sin borrar lo escrito)
    Dim ws As Worksheet, i As Long, r As Long, k As Long, c As String, libre As Long
    Set ws = HojaAuto(AU_HOJA)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    LeerEquipos
    For k = 1 To enEq
        r = 0
        For i = AE0 To AE1
            If UCase(Txt(ws.Cells(i, 2).Value)) = UCase(eCod(k)) Then r = i: Exit For
        Next i
        If r = 0 Then
            For i = AE0 To AE1
                If Txt(ws.Cells(i, 2).Value) = "" Then r = i: Exit For
            Next i
            If r > 0 Then ws.Cells(r, 2).Value = eCod(k)
        End If
        If r > 0 Then ws.Cells(r, 3).Value = eDes(k)
    Next k
    For k = 1 To rN
        r = 0
        For i = AR0 To AR1
            If Txt(ws.Cells(i, 2).Value) = rNum(k) Then r = i: Exit For
        Next i
        If r = 0 Then
            For i = AR0 To AR1
                If Txt(ws.Cells(i, 2).Value) = "" Then r = i: Exit For
            Next i
            If r > 0 Then
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 2).Value = rNum(k)
            End If
        End If
        If r > 0 Then ws.Cells(r, 3).Value = rNom(k)
    Next k
End Sub


'========================================================================== CIERRE, NUEVA OBRA Y DATOS ANTERIORES
Public Sub CerrarAuto()
    ' fin de mes: el horómetro avanza con las horas del mes; se vacían el tareo y el DETALLE DIARIO
    Dim ws As Worksheet, wd As Worksheet, x As Long, e As Long, k As Long, hm() As Double
    Set ws = HojaAuto(AU_HOJA)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    LeerAuto
    LeerPartes415
    ReDim hm(1 To AE1 - AE0 + 1)
    For x = 1 To pnPar
        e = Buscar(eqCol, UCase(eCod(pEq(x))))
        If e > 0 Then hm(e) = hm(e) + HorasFila(pR(x))
    Next x
    For e = 1 To AE1 - AE0 + 1
        If eqFila(e) > 0 And Not IsEmpty(eqHoro(e)) Then
            ws.Cells(eqFila(e), 9).Value = Redondear(Num(eqHoro(e)) + hm(e), 2)
        End If
    Next e
    ws.Range(ws.Cells(AP0, AP_DIA1), ws.Cells(AP1, AP_DIA1 + 30)).ClearContents
    Set wd = HojaAuto(DD_HOJA)
    If Not wd Is Nothing Then
        wd.Range(wd.Cells(DD0, 2), wd.Cells(DD1T, 5)).ClearContents
        wd.Range(wd.Cells(DD0, 8), wd.Cells(DD1S, 9)).ClearContents
        wd.Range(wd.Cells(DD0, 11), wd.Cells(DD1S, 12)).ClearContents
    End If
End Sub

Public Sub VaciarAuto()
    ' NUEVA OBRA: listas de equipos, rubros, personal y detalle diario
    Dim ws As Worksheet, wd As Worksheet
    Set ws = HojaAuto(AU_HOJA)
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Range(ws.Cells(AE0, 2), ws.Cells(AE1, 10)).ClearContents
        ws.Range(ws.Cells(AR0, 2), ws.Cells(AR1, 5)).ClearContents
        ws.Range(ws.Cells(AP0, 3), ws.Cells(AP1, 6)).ClearContents
        ws.Range(ws.Cells(AP0, AP_DIA1), ws.Cells(AP1, AP_DIA1 + 30)).ClearContents
    End If
    Set wd = HojaAuto(DD_HOJA)
    If Not wd Is Nothing Then
        wd.Range(wd.Cells(DD0, 2), wd.Cells(DD1T, 5)).ClearContents
        wd.Range(wd.Cells(DD0, 8), wd.Cells(DD1S, 9)).ClearContents
        wd.Range(wd.Cells(DD0, 11), wd.Cells(DD1S, 12)).ClearContents
    End If
End Sub

Public Sub TraerSalidasAnteriores(ByRef itNum() As String)
    ' TRAER DATOS ANTERIORES: destino (partida) y PEDIDO POR de cada salida de MOVIMIENTOS -> DETALLE DIARIO
    Dim wd As Worksheet, m As Long, r As Long, d As String
    Set wd = HojaAuto(DD_HOJA)
    If wd Is Nothing Then Exit Sub
    wd.Range(wd.Cells(DD0, 8), wd.Cells(DD1S, 9)).ClearContents
    wd.Range(wd.Cells(DD0, 11), wd.Cells(DD1S, 12)).ClearContents
    r = DD0
    For m = 1 To anMov
        If vTipo(m) = T_SAL And vFec(m) >= MesInicio() And vFec(m) <= MesFin() Then
            If r > DD1S Then Exit For
            d = vProc(m)
            If d = "" Then d = vPart(m)
            If d <> "" Or vPed(m) <> "" Then
                wd.Cells(r, 8).Value = CDate(vFec(m))
                wd.Cells(r, 9).NumberFormat = "@"
                wd.Cells(r, 9).Value = itNum(vItem(m))
                wd.Cells(r, 11).Value = d
                wd.Cells(r, 12).Value = vPed(m)
                r = r + 1
            End If
        End If
    Next m
End Sub

Public Sub TraerTrabajosAnteriores()
    ' TRAER DATOS ANTERIORES: trabajos de cada equipo por día de PARTES -> DETALLE DIARIO
    ' (LeerPartesHoja ya se ejecutó)
    Dim wd As Worksheet, x As Long, r As Long, i As Long, t As String
    Set wd = HojaAuto(DD_HOJA)
    If wd Is Nothing Then Exit Sub
    wd.Range(wd.Cells(DD0, 2), wd.Cells(DD1T, 5)).ClearContents
    r = DD0
    For x = 1 To pnPar
        i = pR(pOrden(x))
        If pFec(pOrden(x)) >= MesInicio() And pFec(pOrden(x)) <= MesFin() Then
            If r > DD1T Then Exit For
            t = Limpio(pV(i, PC_TRAB))
            If Txt(pV(i, PC_PART)) <> "" Then t = Txt(pV(i, PC_PART)) & " – " & t
            If t <> "" Then
                wd.Cells(r, 2).Value = CDate(pFec(pOrden(x)))
                wd.Cells(r, 3).Value = eCod(pEq(pOrden(x)))
                wd.Cells(r, 4).Value = t
                If HorasFila(i) > 0 Then wd.Cells(r, 5).Value = HorasFila(i)
                r = r + 1
            End If
        End If
    Next x
End Sub


'========================================================================== BOTONES DEL INICIO
Public Sub IrAutollenado()
    Dim ws As Worksheet
    Set ws = HojaAuto(AU_HOJA)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("D5").Select
End Sub

Public Sub IrDetalle()
    Dim ws As Worksheet
    Set ws = HojaAuto(DD_HOJA)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("B6").Select
End Sub
