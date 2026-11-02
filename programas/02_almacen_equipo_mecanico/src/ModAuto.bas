Attribute VB_Name = "ModAuto"
Option Explicit

'==========================================================================================
'  DATOS QUE COMPLETAN EL KARDEX Y EL PARTE DIARIO
'   - DETALLE SALIDAS: N° de vale, PEDIDO POR y destino de cada salida del MOV. ALMACÉN (kardex y bincard)
'   - AUTOLLENADO (opcional): jornada para equipos sin horario, trabajo por defecto, horómetro
'     inicial y PEDIDO POR cuando el DETALLE SALIDAS no lo tiene (maestro de obra o por rubro).
'  Los nombres de los obreros NO se llenan solos: PEDIDO POR = lo escrito en DETALLE SALIDAS
'  o, si está vacío, el MAESTRO DE OBRA de DATOS (o el de cada rubro).
'==========================================================================================

Public Const AU_HOJA As String = "AUTOLLENADO"
Public Const DS_HOJA As String = "DETALLE SALIDAS"
' ---- opciones (columna D)
Public Const AU_PARTES As String = "D5"
Public Const AU_M1 As String = "D6"
Public Const AU_M2 As String = "D7"
Public Const AU_T1 As String = "D8"
Public Const AU_T2 As String = "D9"
Public Const AU_TRAB As String = "D10"
Public Const AU_PMODO As String = "D11"
Public Const AU_POBS As String = "D13"
' ---- tablas de AUTOLLENADO
Public Const AE0 As Long = 17                 ' equipos: filas 17..36 (B código ... J trabajo)
Public Const AE1 As Long = 36
Public Const AR0 As Long = 40                 ' rubros: filas 40..59 (B rubro, C descripción, D pedido por, E observación)
Public Const AR1 As Long = 59
' ---- DETALLE SALIDAS: B fecha, C ítem, D descripción, E unidad, F cantidad, G vale, H pedido por, I destino, J verificación
Public Const DS0 As Long = 6
Public Const DS1 As Long = 2005
Public Const D_MAESTRO As String = "C42"      ' DATOS: maestro de obra

' ---- lo leído
Public auHay As Boolean
Public auPartes As Boolean
Public auM1 As Double, auM2 As Double, auT1 As Double, auT2 As Double
Public auTrab As String, auModo As String, auNom As String, auObs As String
Private eqCol As Collection
Private eqOper() As String, eqM1() As Double, eqM2() As Double, eqT1() As Double, eqT2() As Double
Private eqHoro() As Variant, eqTrab() As String, eqFila() As Long
Private rbCol As Collection
Private rbNom() As String, rbObs() As String
Private saCol As Collection                   ' "aaaammdd|ITEM" -> fila
Private saVale() As String, saDest() As String, saPed() As String, saN As Long


'========================================================================== UTILIDADES
Public Function HojaAuto(ByVal nombre As String) As Worksheet
    Set HojaAuto = Nothing
    On Error Resume Next
    Set HojaAuto = ThisWorkbook.Worksheets(nombre)
End Function

Public Function HoraDe(ByVal v As Variant, ByVal def As Double) As Double
    ' hora como fracción del día (acepta 07:30, 7.30 o 0.3125); def si está vacía
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

Public Function Minuto(ByVal t As Double) As Double
    Minuto = Round(t * 1440, 0) / 1440
End Function

Public Function ClaveDia(ByVal f As Double) As String
    ClaveDia = Format(CDate(f), "yyyymmdd")
End Function

Public Function Reloj(ByVal t0 As Double, ByVal horas As Double, ByVal m2 As Double, ByVal t1 As Double) As Double
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

Public Function Maestro() As String
    Maestro = Limpio(Hoja("DATOS").Range(D_MAESTRO).Value)
End Function


'========================================================================== LEER LA CONFIGURACIÓN
Public Sub LeerAuto()
    Dim ws As Worksheet, wd As Worksheet, v As Variant, i As Long, n As Long, c As String, s As String, k As String
    auHay = False
    auPartes = False
    auModo = "MAESTRO"
    Set eqCol = New Collection
    Set rbCol = New Collection
    Set saCol = New Collection
    saN = 0
    auNom = Maestro()
    auM1 = TimeSerial(7, 0, 0): auM2 = TimeSerial(12, 0, 0): auT1 = TimeSerial(13, 0, 0): auT2 = TimeSerial(16, 0, 0)
    Set ws = HojaAuto(AU_HOJA)
    If Not ws Is Nothing Then
        auHay = True
        auPartes = (UCase(Left(Txt(ws.Range(AU_PARTES).Value), 1)) = "S")
        auM1 = HoraDe(ws.Range(AU_M1).Value, auM1)
        auM2 = HoraDe(ws.Range(AU_M2).Value, auM2)
        auT1 = HoraDe(ws.Range(AU_T1).Value, auT1)
        auT2 = HoraDe(ws.Range(AU_T2).Value, auT2)
        auTrab = Limpio(ws.Range(AU_TRAB).Value)
        s = UCase(Normal(ws.Range(AU_PMODO).Value))
        If InStr(s, "RUBRO") > 0 Then
            auModo = "RUBRO"
        ElseIf s = "NO" Then
            auModo = "NO"
        Else
            auModo = "MAESTRO"
        End If
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
    End If
    ' detalle de las salidas
    Set wd = HojaAuto(DS_HOJA)
    If wd Is Nothing Then Exit Sub
    n = DS1 - DS0 + 1
    ReDim saVale(1 To n): ReDim saDest(1 To n): ReDim saPed(1 To n)
    v = wd.Range(wd.Cells(DS0, 1), wd.Cells(DS1, 9)).Value
    For i = 1 To n
        If Fecha(v(i, 2)) > 0 And Txt(v(i, 3)) <> "" Then
            k = ClaveDia(Fecha(v(i, 2))) & "|" & UCase(Txt(v(i, 3)))
            If Buscar(saCol, k) = 0 Then
                saN = saN + 1
                saVale(saN) = Txt(v(i, 7))
                saPed(saN) = Limpio(v(i, 8))
                saDest(saN) = Limpio(v(i, 9))
                saCol.Add saN, k
            End If
        End If
    Next i
End Sub


'========================================================================== KARDEX: N° DE VALE, PEDIDO POR Y DESTINO
Public Sub AutoSalida(ByVal item As String, ByVal rubro As String, ByVal f As Double, ByRef ped As String, _
                      ByRef obs As String, ByRef vale As String)
    Dim j As Long, k As Long
    j = Buscar(saCol, ClaveDia(f) & "|" & UCase(item))
    If j > 0 Then
        vale = saVale(j)
        ped = saPed(j)
        obs = saDest(j)
    End If
    k = Buscar(rbCol, UCase(rubro))
    If obs = "" Then
        If k > 0 Then obs = rbObs(k)
        If obs = "" Then obs = auObs
    End If
    If ped <> "" Or auModo = "NO" Then Exit Sub
    If auModo = "RUBRO" And k > 0 Then ped = rbNom(k)
    If ped = "" Then ped = auNom                  ' maestro de obra (DATOS)
End Sub


'========================================================================== PARTE DIARIO (equipos sin MOV. EQUIPOS)
Public Sub DatosEquipo(ByVal cod As String, ByRef oper As String, ByRef m1 As Double, ByRef m2 As Double, _
                       ByRef t1 As Double, ByRef t2 As Double, ByRef trab As String)
    ' operador, jornada y trabajo por defecto de un equipo (AUTOLLENADO)
    Dim e As Long
    e = 0
    If Not eqCol Is Nothing Then e = Buscar(eqCol, UCase(cod))
    If e > 0 Then
        m1 = eqM1(e): m2 = eqM2(e): t1 = eqT1(e): t2 = eqT2(e)
        If oper = "" Then oper = eqOper(e)
        trab = eqTrab(e)
    Else
        m1 = auM1: m2 = auM2: t1 = auT1: t2 = auT2
        trab = ""
    End If
    If trab = "" Then trab = auTrab
End Sub

Public Sub AutoParte(ByVal k As Long, ByVal f As Double, ByVal g As Double, ByVal h As Double, ByRef horo As Variant)
    ' una línea del parte desde el CONTROL DE EQUIPO (equipo sin filas en MOV. EQUIPOS): horario desde la jornada
    Dim m1 As Double, m2 As Double, t1 As Double, t2 As Double, oper As String, trab As String
    Dim fin As Double, esDia As Boolean
    oper = eOper(k)
    DatosEquipo eCod(k), oper, m1, m2, t1, t2, trab
    esDia = (InStr(1, Normal(eUnd(k)), "DIA", vbTextCompare) > 0)
    pnPar = pnPar + 1
    pR(pnPar) = pnPar
    pEq(pnPar) = k
    pFec(pnPar) = f
    pOrden(pnPar) = pnPar
    pV(pnPar, PC_FEC) = CDate(f)
    pV(pnPar, PC_EQ) = eCod(k)
    pV(pnPar, PC_OPER) = oper
    pV(pnPar, PC_TRAB) = trab
    pV(pnPar, PC_HREAL) = h
    pV(pnPar, PC_COMB) = g
    If h > 0 Then
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
        pV(pnPar, PC_ULT + 1) = m1
        pV(pnPar, PC_ULT + 2) = fin
    End If
    If Not IsEmpty(horo) And Not esDia Then
        pV(pnPar, PC_HINI) = Redondear(Num(horo), 2)
        pV(pnPar, PC_HFIN) = Redondear(Num(horo) + h, 2)
        horo = Num(horo) + h
    End If
End Sub

Public Function HorometroInicial(ByVal k As Long) As Variant
    Dim e As Long
    HorometroInicial = Empty
    If eqCol Is Nothing Then Exit Function
    e = Buscar(eqCol, UCase(eCod(k)))
    If e > 0 Then HorometroInicial = eqHoro(e)
End Function


'========================================================================== SINCRONIZAR LISTAS (ACTUALIZAR CUADROS)
Public Sub SincronizarAuto()
    ' pone en AUTOLLENADO los equipos de EQUIPOS y los rubros del INV. que falten (sin borrar lo escrito)
    Dim ws As Worksheet, i As Long, r As Long, k As Long
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
    ' fin de mes: horómetro inicial del mes siguiente; se vacían DETALLE SALIDAS y MOV. EQUIPOS
    Dim ws As Worksheet, wd As Worksheet, x As Long, e As Long, hm() As Double, hf() As Double
    On Error Resume Next
    LeerAuto
    LeerPartes415
    Set ws = HojaAuto(AU_HOJA)
    If Not ws Is Nothing Then
        ReDim hm(1 To AE1 - AE0 + 1)
        ReDim hf(1 To AE1 - AE0 + 1)
        For x = 1 To pnPar
            e = Buscar(eqCol, UCase(eCod(pEq(x))))
            If e > 0 Then
                hm(e) = hm(e) + HorasFila(pR(x))
                If EsNum(pV(pR(x), PC_HFIN)) Then
                    If Num(pV(pR(x), PC_HFIN)) > hf(e) Then hf(e) = Num(pV(pR(x), PC_HFIN))
                End If
            End If
        Next x
        For e = 1 To AE1 - AE0 + 1
            If eqFila(e) > 0 Then
                If hf(e) > 0 Then
                    ws.Cells(eqFila(e), 9).Value = Redondear(hf(e), 2)       ' última lectura del mes
                ElseIf Not IsEmpty(eqHoro(e)) Then
                    ws.Cells(eqFila(e), 9).Value = Redondear(Num(eqHoro(e)) + hm(e), 2)
                End If
            End If
        Next e
    End If
    VaciarDetalles
End Sub

Public Sub VaciarDetalles()
    ' DETALLE SALIDAS y MOV. EQUIPOS (solo lo digitado; las fórmulas quedan)
    Dim wd As Worksheet, wm As Worksheet
    On Error Resume Next
    Set wd = HojaAuto(DS_HOJA)
    If Not wd Is Nothing Then
        wd.Range(wd.Cells(DS0, 2), wd.Cells(DS1, 3)).ClearContents
        wd.Range(wd.Cells(DS0, 6), wd.Cells(DS1, 9)).ClearContents
    End If
    Set wm = HojaAuto(MV_HOJA)
    If Not wm Is Nothing Then
        wm.Range(wm.Cells(MQ0, MQ_FEC), wm.Cells(MQ1, MQ_COD)).ClearContents
        wm.Range(wm.Cells(MQ0, MQ_OPER), wm.Cells(MQ1, MQ_TA)).ClearContents
        wm.Range(wm.Cells(MQ0, MQ_HINI), wm.Cells(MQ1, MQ_HFIN)).ClearContents
        wm.Range(wm.Cells(MQ0, MQ_HDIG), wm.Cells(MQ1, MQ_HDIG)).ClearContents
        wm.Range(wm.Cells(MQ0, MQ_GAL), wm.Cells(MQ1, MQ_OBS)).ClearContents
    End If
End Sub

Public Sub VaciarAuto()
    ' NUEVA OBRA
    Dim ws As Worksheet
    Set ws = HojaAuto(AU_HOJA)
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Range(ws.Cells(AE0, 2), ws.Cells(AE1, 10)).ClearContents
        ws.Range(ws.Cells(AR0, 2), ws.Cells(AR1, 5)).ClearContents
    End If
    VaciarDetalles
End Sub

Public Sub TraerSalidasAnteriores(ByRef itNum() As String)
    ' TRAER DATOS ANTERIORES: N° de documento, PEDIDO POR y destino de cada salida de MOVIMIENTOS -> DETALLE SALIDAS
    Dim wd As Worksheet, m As Long, r As Long, d As String, k As String, j As Long, col As New Collection
    Dim fe() As Double, it() As String, q() As Double, va() As String, pe() As String, de() As String, n As Long
    Set wd = HojaAuto(DS_HOJA)
    If wd Is Nothing Then Exit Sub
    If anMov < 1 Then Exit Sub
    ReDim fe(1 To anMov): ReDim it(1 To anMov): ReDim q(1 To anMov): ReDim va(1 To anMov)
    ReDim pe(1 To anMov): ReDim de(1 To anMov)
    For m = 1 To anMov
        If vTipo(m) = T_SAL And vFec(m) >= MesInicio() And vFec(m) <= MesFin() Then
            k = ClaveDia(vFec(m)) & "|" & UCase(itNum(vItem(m)))
            j = Buscar(col, k)
            If j = 0 Then
                n = n + 1
                j = n
                col.Add n, k
                fe(n) = vFec(m)
                it(n) = itNum(vItem(m))
            End If
            q(j) = q(j) + vCant(m)
            If va(j) = "" Then va(j) = vNDoc(m)
            If pe(j) = "" Then pe(j) = vPed(m)
            d = vProc(m)
            If d = "" Then d = vPart(m)
            If d <> "" And InStr(1, de(j), d, vbTextCompare) = 0 Then
                If de(j) = "" Then de(j) = d Else de(j) = de(j) & "; " & d
            End If
        End If
    Next m
    wd.Range(wd.Cells(DS0, 2), wd.Cells(DS1, 3)).ClearContents
    wd.Range(wd.Cells(DS0, 6), wd.Cells(DS1, 9)).ClearContents
    r = DS0
    For j = 1 To n
        If r > DS1 Then Exit For
        wd.Cells(r, 2).Value = CDate(fe(j))
        wd.Cells(r, 3).NumberFormat = "@"
        wd.Cells(r, 3).Value = it(j)
        wd.Cells(r, 6).Value = q(j)
        wd.Cells(r, 7).Value = va(j)
        wd.Cells(r, 8).Value = pe(j)
        wd.Cells(r, 9).Value = de(j)
        r = r + 1
    Next j
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
    Set ws = HojaAuto(DS_HOJA)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("G6").Select
End Sub
