Attribute VB_Name = "ModEquipos"
Option Explicit

'==========================================================================================
'  CUADRO DE MOVIMIENTO DIARIO DE EQUIPOS Y MAQUINARIA (hoja MOV. EQUIPOS)
'  Una fila por equipo y día (o por cada trabajo del día): operador, horario de mañana y tarde,
'  horómetro inicial y final, combustible, partida y trabajo realizado.
'  De aquí salen el 4.15 (horas y galones por día), los partes diarios y el resumen de EQUIPOS.
'  También: REVISIÓN de datos antes de generar los formatos finales.
'==========================================================================================

Public Const MV_HOJA As String = "MOV. EQUIPOS"
Public Const RV_HOJA As String = "REVISION"
Public Const MQ0 As Long = 7                  ' primera fila de datos
Public Const MQ1 As Long = 3006
Public Const MQ_FEC As Long = 2
Public Const MQ_COD As Long = 3
Public Const MQ_DES As Long = 4
Public Const MQ_OPER As Long = 5
Public Const MQ_MDE As Long = 6
Public Const MQ_MA As Long = 7
Public Const MQ_TDE As Long = 8
Public Const MQ_TA As Long = 9
Public Const MQ_HHOR As Long = 10             ' horas según horario (fórmula)
Public Const MQ_HINI As Long = 11
Public Const MQ_HFIN As Long = 12
Public Const MQ_HHRM As Long = 13             ' horas según horómetro (fórmula)
Public Const MQ_HDIG As Long = 14             ' horas digitadas (si no hay horario ni horómetro)
Public Const MQ_HVAL As Long = 15             ' horas (o días) a valorizar (fórmula)
Public Const MQ_GAL As Long = 16
Public Const MQ_PART As Long = 17
Public Const MQ_TRAB As Long = 18
Public Const MQ_OBS As Long = 19
Public Const MQ_VER As Long = 20
' ---- EQUIPOS: resumen de uso por fechas
Public Const EC_PRIM As Long = 26             ' Z
Public Const EC_ULTD As Long = 27             ' AA
Public Const EC_DIAS As Long = 28             ' AB
Public Const EC_FECHAS As Long = 29           ' AC

' ---- filas del mes leídas
Public mqN As Long
Public mqFila() As Long, mqFec() As Double, mqCod() As String, mqOper() As String
Public mqMDE() As Variant, mqMA() As Variant, mqTDE() As Variant, mqTA() As Variant
Public mqHI() As Variant, mqHF() As Variant, mqH() As Double, mqG() As Double
Public mqPart() As String, mqTrab() As String, mqObs() As String, mqHorario() As Boolean
Public mqCol As Collection                    ' códigos con filas en el mes


Private Function HoraV(ByVal v As Variant) As Variant
    ' hora (fracción del día) o Empty
    Dim t As Double
    HoraV = Empty
    t = HoraDe(v, -1)
    If t >= 0 Then HoraV = t
End Function

Private Function Tr(ByVal a As Variant, ByVal b As Variant) As Double
    Tr = 0
    If Not IsEmpty(a) And Not IsEmpty(b) Then
        If b > a Then Tr = (b - a) * 24
    End If
End Function

Public Function EsUnidadDia(ByVal und As String) As Boolean
    EsUnidadDia = (InStr(1, Normal(und), "DIA", vbTextCompare) > 0)
End Function

Public Function EsUnidadHora(ByVal und As String) As Boolean
    ' H.M., H.E., HORA: en el 4.15 las horas van como tiempo (formato [h]:mm, igual que el informe)
    Dim s As String
    s = Normal(und)
    EsUnidadHora = (Not EsUnidadDia(und)) And (InStr(s, "H.") > 0 Or InStr(s, "HORA") > 0 Or s = "HM" Or s = "HE")
End Function


'========================================================================== LEER MOV. EQUIPOS
Public Sub LeerMovEq()
    ' filas del mes con fecha y código (LeerEquipos debe estar hecho para saber la unidad)
    Dim ws As Worksheet, v As Variant, i As Long, n As Long, f As Double, c As String, k As Long
    Dim hh As Double, hr As Double, hd As Double, und As String
    mqN = 0
    Set mqCol = New Collection
    Set ws = HojaAuto(MV_HOJA)
    If ws Is Nothing Then Exit Sub
    v = ws.Range(ws.Cells(MQ0, 1), ws.Cells(MQ1, MQ_OBS)).Value
    n = UBound(v, 1)
    ReDim mqFila(1 To n): ReDim mqFec(1 To n): ReDim mqCod(1 To n): ReDim mqOper(1 To n)
    ReDim mqMDE(1 To n): ReDim mqMA(1 To n): ReDim mqTDE(1 To n): ReDim mqTA(1 To n)
    ReDim mqHI(1 To n): ReDim mqHF(1 To n): ReDim mqH(1 To n): ReDim mqG(1 To n)
    ReDim mqPart(1 To n): ReDim mqTrab(1 To n): ReDim mqObs(1 To n): ReDim mqHorario(1 To n)
    For i = 1 To n
        f = Fecha(v(i, MQ_FEC))
        c = Txt(v(i, MQ_COD))
        If f >= MesInicio() And f <= MesFin() And c <> "" Then
            mqN = mqN + 1
            mqFila(mqN) = MQ0 + i - 1
            mqFec(mqN) = f
            mqCod(mqN) = c
            mqOper(mqN) = Limpio(v(i, MQ_OPER))
            mqMDE(mqN) = HoraV(v(i, MQ_MDE))
            mqMA(mqN) = HoraV(v(i, MQ_MA))
            mqTDE(mqN) = HoraV(v(i, MQ_TDE))
            mqTA(mqN) = HoraV(v(i, MQ_TA))
            If EsNum(v(i, MQ_HINI)) Then mqHI(mqN) = Num(v(i, MQ_HINI)) Else mqHI(mqN) = Empty
            If EsNum(v(i, MQ_HFIN)) Then mqHF(mqN) = Num(v(i, MQ_HFIN)) Else mqHF(mqN) = Empty
            hh = Tr(mqMDE(mqN), mqMA(mqN)) + Tr(mqTDE(mqN), mqTA(mqN))
            mqHorario(mqN) = (hh > 0)
            hr = 0
            If Not IsEmpty(mqHI(mqN)) And Not IsEmpty(mqHF(mqN)) Then
                If mqHF(mqN) > mqHI(mqN) Then hr = mqHF(mqN) - mqHI(mqN)
            End If
            hd = Num(v(i, MQ_HDIG))
            und = ""
            k = Buscar(eCol, UCase(c))
            If k > 0 Then und = eUnd(k)
            If EsUnidadDia(und) Then
                If hh > 0 Or hr > 0 Or hd > 0 Then mqH(mqN) = 1
            ElseIf hr > 0 Then
                mqH(mqN) = Redondear(hr, 2)
            ElseIf hh > 0 Then
                mqH(mqN) = Redondear(hh, 2)
            Else
                mqH(mqN) = Redondear(hd, 2)
            End If
            mqG(mqN) = Num(v(i, MQ_GAL))
            mqPart(mqN) = Txt(v(i, MQ_PART))
            mqTrab(mqN) = Limpio(v(i, MQ_TRAB))
            mqObs(mqN) = Limpio(v(i, MQ_OBS))
            If Buscar(mqCol, UCase(c)) = 0 Then mqCol.Add mqN, UCase(c)
        End If
    Next i
End Sub

Public Function TieneMov(ByVal cod As String) As Boolean
    TieneMov = False
    If mqCol Is Nothing Then Exit Function
    TieneMov = (Buscar(mqCol, UCase(cod)) > 0)
End Function


'========================================================================== PARTES DESDE MOV. EQUIPOS
Public Sub PartesDesdeMov(ByVal k As Long)
    ' agrega a pV una línea por cada fila de MOV. EQUIPOS del equipo k (horario y horómetro reales)
    Dim x As Long, m1 As Double, m2 As Double, t1 As Double, t2 As Double, oper As String, trab As String
    Dim reloj0 As Double, diaAnt As Double, fin As Double, de As Variant, a As Variant
    oper = ""
    DatosEquipo eCod(k), oper, m1, m2, t1, t2, trab
    diaAnt = 0
    For x = 1 To mqN
        If UCase(mqCod(x)) = UCase(eCod(k)) Then
            pnPar = pnPar + 1
            pR(pnPar) = pnPar
            pEq(pnPar) = k
            pFec(pnPar) = mqFec(x)
            pOrden(pnPar) = pnPar
            pV(pnPar, PC_FEC) = CDate(mqFec(x))
            pV(pnPar, PC_EQ) = eCod(k)
            If mqOper(x) <> "" Then
                pV(pnPar, PC_OPER) = mqOper(x)
            ElseIf eOper(k) <> "" Then
                pV(pnPar, PC_OPER) = eOper(k)
            Else
                pV(pnPar, PC_OPER) = oper
            End If
            pV(pnPar, PC_PART) = mqPart(x)
            If mqTrab(x) <> "" Then pV(pnPar, PC_TRAB) = mqTrab(x) Else pV(pnPar, PC_TRAB) = trab
            pV(pnPar, PC_HREAL) = mqH(x)
            pV(pnPar, PC_COMB) = mqG(x)
            pV(pnPar, PC_OBS) = mqObs(x)
            pV(pnPar, PC_HINI) = mqHI(x)
            pV(pnPar, PC_HFIN) = mqHF(x)
            If mqHorario(x) Then
                pV(pnPar, PC_MDE) = mqMDE(x)
                pV(pnPar, PC_MA) = mqMA(x)
                pV(pnPar, PC_TDE) = mqTDE(x)
                pV(pnPar, PC_TA) = mqTA(x)
                If Not IsEmpty(mqMDE(x)) Then de = mqMDE(x) Else de = mqTDE(x)
                If Not IsEmpty(mqTA(x)) Then a = mqTA(x) Else a = mqMA(x)
                pV(pnPar, PC_ULT + 1) = de
                pV(pnPar, PC_ULT + 2) = a
            ElseIf auPartes And mqH(x) > 0 Then
                ' sin horario: se arma desde la jornada (opción de AUTOLLENADO), una línea tras otra
                If mqFec(x) <> diaAnt Then reloj0 = m1
                diaAnt = mqFec(x)
                fin = Reloj(reloj0, mqH(x), m2, t1)
                If reloj0 < m2 Then
                    pV(pnPar, PC_MDE) = reloj0
                    If fin <= m2 + 0.0000001 Then
                        pV(pnPar, PC_MA) = fin
                    Else
                        pV(pnPar, PC_MA) = m2
                        pV(pnPar, PC_TDE) = t1
                        pV(pnPar, PC_TA) = fin
                    End If
                Else
                    pV(pnPar, PC_TDE) = reloj0
                    pV(pnPar, PC_TA) = fin
                End If
                pV(pnPar, PC_ULT + 1) = reloj0
                pV(pnPar, PC_ULT + 2) = fin
                reloj0 = fin
                If reloj0 >= m2 - 0.0000001 And reloj0 < t1 Then reloj0 = t1
            End If
        End If
    Next x
End Sub

Public Sub OrdenarPartes()
    ' por equipo, fecha y orden de registro (ArmarPartes junta las líneas del mismo equipo y día)
    Dim llave() As Double, x As Long
    If pnPar < 2 Then Exit Sub
    ReDim llave(1 To pnPar)
    For x = 1 To pnPar
        llave(x) = CDbl(pEq(x)) * 10000000000# + pFec(x) * 10000# + x
    Next x
    Ordenar llave, pOrden, pnPar
End Sub


'========================================================================== RESUMEN DE USO EN EQUIPOS
Public Sub ResumenUso(ByVal cod As String, ByRef dias() As Boolean)
    ' escribe en EQUIPOS: primer día, último día, días trabajados y fechas de uso del mes
    Dim we As Worksheet, fe As Long, d As Long, n As Long, prim As Long, ult As Long, s As String
    Set we = Hoja("EQUIPOS")
    fe = 0
    For d = EQ0 To EQMAX
        If UCase(Txt(we.Cells(d, EC_COD).Value)) = UCase(cod) Then fe = d: Exit For
    Next d
    If fe = 0 Then Exit Sub
    For d = 1 To 31
        If dias(d) Then
            n = n + 1
            If prim = 0 Then prim = d
            ult = d
            If s = "" Then s = CStr(d) Else s = s & ", " & d
        End If
    Next d
    On Error Resume Next
    If n > 0 Then
        we.Cells(fe, EC_PRIM).Value = CDate(DateSerial(Year(CDate(MesInicio())), Month(CDate(MesInicio())), prim))
        we.Cells(fe, EC_ULTD).Value = CDate(DateSerial(Year(CDate(MesInicio())), Month(CDate(MesInicio())), ult))
        we.Cells(fe, EC_DIAS).Value = n
        we.Cells(fe, EC_FECHAS).Value = s & " / " & Format(CDate(MesInicio()), "mm/yyyy")
    Else
        we.Range(we.Cells(fe, EC_PRIM), we.Cells(fe, EC_FECHAS)).ClearContents
    End If
End Sub


'========================================================================== TRAER DATOS ANTERIORES
Public Sub TraerMovAnterior()
    ' partes de la hoja PARTES (formato anterior) -> MOV. EQUIPOS (LeerPartesHoja ya se ejecutó)
    Dim ws As Worksheet, x As Long, r As Long, i As Long
    Set ws = HojaAuto(MV_HOJA)
    If ws Is Nothing Then Exit Sub
    VaciarMov ws
    r = MQ0
    For x = 1 To pnPar
        i = pR(pOrden(x))
        If pFec(pOrden(x)) >= MesInicio() And pFec(pOrden(x)) <= MesFin() Then
            If r > MQ1 Then Exit For
            ws.Cells(r, MQ_FEC).Value = CDate(pFec(pOrden(x)))
            ws.Cells(r, MQ_COD).Value = eCod(pEq(pOrden(x)))
            ws.Cells(r, MQ_OPER).Value = Limpio(pV(i, PC_OPER))
            If Not IsEmpty(pV(i, PC_MDE)) Then ws.Cells(r, MQ_MDE).Value = pV(i, PC_MDE)
            If Not IsEmpty(pV(i, PC_MA)) Then ws.Cells(r, MQ_MA).Value = pV(i, PC_MA)
            If Not IsEmpty(pV(i, PC_TDE)) Then ws.Cells(r, MQ_TDE).Value = pV(i, PC_TDE)
            If Not IsEmpty(pV(i, PC_TA)) Then ws.Cells(r, MQ_TA).Value = pV(i, PC_TA)
            If EsNum(pV(i, PC_HINI)) Then ws.Cells(r, MQ_HINI).Value = Num(pV(i, PC_HINI))
            If EsNum(pV(i, PC_HFIN)) Then ws.Cells(r, MQ_HFIN).Value = Num(pV(i, PC_HFIN))
            If HorasFila(i) > 0 Then ws.Cells(r, MQ_HDIG).Value = HorasFila(i)
            If Num(pV(i, PC_COMB)) <> 0 Then ws.Cells(r, MQ_GAL).Value = Num(pV(i, PC_COMB))
            ws.Cells(r, MQ_PART).NumberFormat = "@"
            ws.Cells(r, MQ_PART).Value = Txt(pV(i, PC_PART))
            ws.Cells(r, MQ_TRAB).Value = Limpio(pV(i, PC_TRAB))
            ws.Cells(r, MQ_OBS).Value = Limpio(pV(i, PC_OBS))
            r = r + 1
        End If
    Next x
End Sub

Public Sub VaciarMov(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(ws.Cells(MQ0, MQ_FEC), ws.Cells(MQ1, MQ_COD)).ClearContents
    ws.Range(ws.Cells(MQ0, MQ_OPER), ws.Cells(MQ1, MQ_TA)).ClearContents
    ws.Range(ws.Cells(MQ0, MQ_HINI), ws.Cells(MQ1, MQ_HFIN)).ClearContents
    ws.Range(ws.Cells(MQ0, MQ_HDIG), ws.Cells(MQ1, MQ_HDIG)).ClearContents
    ws.Range(ws.Cells(MQ0, MQ_GAL), ws.Cells(MQ1, MQ_OBS)).ClearContents
End Sub


'========================================================================== REVISIÓN DE DATOS
Private Sub Obs(ByVal ws As Worksheet, ByRef r As Long, ByVal hoja As String, ByVal fila As Long, ByVal dato As String, _
                ByVal texto As String)
    r = r + 1
    ws.Cells(r, 2).Value = r - 4
    ws.Cells(r, 3).Value = hoja
    If fila > 0 Then
        ws.Cells(r, 4).Value = fila
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 4), Address:="", SubAddress:="'" & hoja & "'!A" & fila, TextToDisplay:=CStr(fila)
        On Error GoTo 0
    End If
    ws.Cells(r, 5).Value = dato
    ws.Cells(r, 6).Value = texto
End Sub

Public Function RevisarDatos() As Long
    ' llena la hoja REVISION con lo que falta o no cuadra; devuelve el número de observaciones
    Dim wr As Worksheet, ws As Worksheet, r As Long, i As Long, x As Long, k As Long, last As Long, n As Long
    Dim d As Double, hh As Double, hr As Double, ult() As Variant, ultF() As Double, c As String, dato As String
    Dim operD As String, m1 As Double, m2 As Double, t1 As Double, t2 As Double, trabD As String, msg As String

    Set wr = HojaAuto(RV_HOJA)
    If wr Is Nothing Then Exit Function
    On Error Resume Next
    wr.Hyperlinks.Delete
    On Error GoTo 0
    wr.Range("B5:F5000").ClearContents
    wr.Range("B2").Value = "Revisión del " & Format(Now, "dd/mm/yyyy hh:mm") & " – mes de " & MesTexto() & _
                           ". Corrija y vuelva a pulsar ACTUALIZAR Y REVISAR hasta que no quede nada; luego GENERAR."
    r = 4
    LeerEquipos
    LeerAuto

    ' DATOS
    Set ws = Hoja("DATOS")
    If Limpio(ws.Range(D_RESID).Value) = "" Then Obs wr, r, "DATOS", ws.Range(D_RESID).Row, "RESIDENTE DE OBRA", "Falta el nombre (va en los informes)."
    If Limpio(ws.Range(D_SUPER).Value) = "" Then Obs wr, r, "DATOS", ws.Range(D_SUPER).Row, "SUPERVISOR DE OBRA", "Falta el nombre (va en los informes)."
    If Maestro() = "" And auModo <> "NO" Then Obs wr, r, "DATOS", ws.Range(D_MAESTRO).Row, "MAESTRO DE OBRA", _
        "Falta el nombre: es quien pide los materiales en el kardex cuando DETALLE SALIDAS no dice otro."

    ' 4.10: saldos negativos
    Set ws = HojaPref("4.10")
    last = UltimaFila(ws, 6, 13, 8000)
    For i = 13 To last
        If Txt(ws.Cells(i, 6).Value) = "ENTRADA" Then
            If EsNum(ws.Cells(i, 40).Value) Then
                If Num(ws.Cells(i, 40).Value) < -0.00001 Then Obs wr, r, ws.Name, i, Txt(ws.Cells(i, 2).Value) & " " & _
                    Left(Txt(ws.Cells(i, 3).Value), 40), "Saldo negativo (" & Format(Num(ws.Cells(i, 40).Value), "0.00") & "): revise las salidas."
            End If
        End If
    Next i

    ' DETALLE SALIDAS
    Set ws = HojaAuto(DS_HOJA)
    If Not ws Is Nothing Then
        For i = DS0 To DS1
            If Fecha(ws.Cells(i, 2).Value) > 0 Then
                dato = Format(CDate(Fecha(ws.Cells(i, 2).Value)), "dd/mm") & "  ítem " & Txt(ws.Cells(i, 3).Value)
                If Not EsNum(ws.Cells(i, 6).Value) Then
                    Obs wr, r, ws.Name, i, dato, "Esta salida ya no está en el 4.10 (bórrela o corrija la fecha / el ítem)."
                Else
                    If Limpio(ws.Cells(i, 8).Value) = "" And auModo = "NO" Then Obs wr, r, ws.Name, i, dato, "Falta quién pide."
                    If Limpio(ws.Cells(i, 9).Value) = "" And auObs = "" Then Obs wr, r, ws.Name, i, dato, _
                        "Falta el destino / partida (va en OBSERVACIONES del kardex)."
                End If
            End If
        Next i
    End If

    ' MOV. EQUIPOS
    Set ws = HojaAuto(MV_HOJA)
    If Not ws Is Nothing Then
        LeerMovEq
        ReDim ult(1 To enEq + 1): ReDim ultF(1 To enEq + 1)
        For i = MQ0 To MQ1
            If Txt(ws.Cells(i, MQ_FEC).Value) <> "" Or Txt(ws.Cells(i, MQ_COD).Value) <> "" Then
                d = Fecha(ws.Cells(i, MQ_FEC).Value)
                c = Txt(ws.Cells(i, MQ_COD).Value)
                dato = c
                If d > 0 Then dato = Format(CDate(d), "dd/mm") & "  " & c
                k = Buscar(eCol, UCase(c))
                If d = 0 Then
                    Obs wr, r, ws.Name, i, dato, "Falta la fecha."
                ElseIf k = 0 Then
                    Obs wr, r, ws.Name, i, dato, "El código no está en EQUIPOS."
                ElseIf d < MesInicio() Or d > MesFin() Then
                    Obs wr, r, ws.Name, i, dato, "Fecha fuera del mes que se controla."
                End If
            End If
        Next i
        For x = 1 To mqN
            k = Buscar(eCol, UCase(mqCod(x)))
            If k > 0 Then
                dato = Format(CDate(mqFec(x)), "dd/mm") & "  " & mqCod(x)
                hh = 0
                If mqHorario(x) Then hh = Tr(mqMDE(x), mqMA(x)) + Tr(mqTDE(x), mqTA(x))
                hr = 0
                If Not IsEmpty(mqHI(x)) And Not IsEmpty(mqHF(x)) Then hr = mqHF(x) - mqHI(x)
                If Not mqHorario(x) Then Obs wr, r, MV_HOJA, mqFila(x), dato, "Falta el horario (mañana / tarde DE – A) para el parte diario."
                If Not EsUnidadDia(eUnd(k)) And InStr(1, Normal(eUnd(k)), "H.M", vbTextCompare) > 0 Then
                    If IsEmpty(mqHI(x)) Or IsEmpty(mqHF(x)) Then
                        Obs wr, r, MV_HOJA, mqFila(x), dato, "Falta el horómetro inicial o final."
                    ElseIf hr <= 0 Then
                        Obs wr, r, MV_HOJA, mqFila(x), dato, "El horómetro final no es mayor que el inicial."
                    End If
                    If Not IsEmpty(mqHI(x)) And Not IsEmpty(ult(k)) Then
                        If Abs(mqHI(x) - ult(k)) > 0.05 And mqFec(x) >= ultF(k) Then Obs wr, r, MV_HOJA, mqFila(x), dato, _
                            "El horómetro inicial (" & Format(mqHI(x), "0.00") & ") no sigue al final anterior (" & Format(ult(k), "0.00") & ")."
                    End If
                    If Not IsEmpty(mqHF(x)) Then
                        ult(k) = mqHF(x)
                        ultF(k) = mqFec(x)
                    End If
                End If
                If hh > 0 And hr > 0 Then
                    If Abs(hh - hr) > 0.5 Then Obs wr, r, MV_HOJA, mqFila(x), dato, "El horario (" & Format(hh, "0.00") & _
                        " h) y el horómetro (" & Format(hr, "0.00") & " h) no coinciden."
                End If
                If mqH(x) <= 0 Then Obs wr, r, MV_HOJA, mqFila(x), dato, "Sin horas (horario, horómetro u horas digitadas)."
                operD = ""
                DatosEquipo mqCod(x), operD, m1, m2, t1, t2, trabD
                If mqOper(x) = "" And eOper(k) = "" And operD = "" Then Obs wr, r, MV_HOJA, mqFila(x), dato, "Falta el operador."
                If mqTrab(x) = "" And trabD = "" Then Obs wr, r, MV_HOJA, mqFila(x), dato, "Falta el trabajo realizado."
            End If
        Next x
    End If

    ' EQUIPOS: horas solo en el 4.15 y saldo de la O/S
    Set ws = Hoja("EQUIPOS")
    For k = 1 To enEq
        i = 0
        For x = EQ0 To EQMAX
            If UCase(Txt(ws.Cells(x, EC_COD).Value)) = UCase(eCod(k)) Then i = x: Exit For
        Next x
        If i > 0 Then
            If Not TieneMov(eCod(k)) And Num(ws.Cells(i, EC_DIAS).Value) > 0 Then Obs wr, r, "EQUIPOS", i, eCod(k), _
                "Tiene horas o combustible en el 4.15 pero ninguna fila en MOV. EQUIPOS (el parte saldrá sin horario real)."
            If EsNum(ws.Cells(i, EC_SALDO).Value) Then
                If Num(ws.Cells(i, EC_SALDO).Value) < 0 Then Obs wr, r, "EQUIPOS", i, eCod(k), "Se pasó de la cantidad contratada en la O/S."
            End If
        End If
    Next k
    RevisarDatos = r - 4
    If r = 4 Then wr.Range("C5").Value = "Todo completo: puede GENERAR los formatos."
End Function

Public Sub ActualizarYRevisar()
    Dim s As String, n As Long
    s = ActualizarCuadros(True)
    n = RevisarDatos()
    If n = 0 Then
        Aviso s & vbCrLf & vbCrLf & "REVISIÓN: todo completo. Ya puede GENERAR los formatos (paso 3) y luego el PDF (paso 4).", _
              "ACTUALIZAR Y REVISAR"
    Else
        Aviso s & vbCrLf & vbCrLf & "REVISIÓN: " & n & " observación(es). Están en la hoja REVISION (con enlace a cada fila)." & _
              vbCrLf & "Corríjalas y vuelva a pulsar ACTUALIZAR Y REVISAR.", "ACTUALIZAR Y REVISAR", 48
        If Not gSilencio Then
            HojaAuto(RV_HOJA).Activate
            HojaAuto(RV_HOJA).Range("B5").Select
        End If
    End If
End Sub

Public Sub IrMovEquipos()
    Dim ws As Worksheet
    Set ws = HojaAuto(MV_HOJA)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("B7").Select
End Sub
