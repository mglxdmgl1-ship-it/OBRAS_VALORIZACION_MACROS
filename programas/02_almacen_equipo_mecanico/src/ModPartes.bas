Attribute VB_Name = "ModPartes"
Option Explicit

'==========================================================================================
'  EQUIPO MECÁNICO: lectura de EQUIPOS y PARTES y armado de los partes diarios (hojas + PDF)
'==========================================================================================

Public Const PD_ALTO As Long = 50             ' filas de un parte diario (1..50, formato PARTE DIARIO DE MAQUINARIA)
Public Const PD_L0 As Long = 15               ' primera fila de trabajos
Public Const PD_N As Long = 8                 ' trabajos por hoja (filas 15..22)
Public Const PD_COLS As Long = 16             ' columnas A..P

' ---- equipos leídos
Public enEq As Long
Public eCod() As String, eDes() As String, eMarca() As String, eModelo() As String, eSerie() As String, eAnio() As String
Public eProv() As String, eOS() As String, eUnd() As String, eCond() As String, eOper() As String, eComb() As String
Public eCol As Collection
' ---- partes válidos (pOrden: por equipo y fecha)
Public pV As Variant
Public pnPar As Long, pnMal As Long
Public pR() As Long, pEq() As Long, pFec() As Double, pOrden() As Long


Public Sub LeerEquipos()
    Dim ws As Worksheet, v As Variant, i As Long, n As Long, last As Long, cod As String
    Set eCol = New Collection
    enEq = 0
    Set ws = Hoja("EQUIPOS")
    last = UltimaFila(ws, EC_COD, EQ0, EQMAX)
    n = last - EQ0 + 1
    If n < 1 Then n = 1
    ReDim eCod(1 To n)
    ReDim eDes(1 To n)
    ReDim eMarca(1 To n)
    ReDim eModelo(1 To n)
    ReDim eSerie(1 To n)
    ReDim eAnio(1 To n)
    ReDim eProv(1 To n)
    ReDim eOS(1 To n)
    ReDim eUnd(1 To n)
    ReDim eCond(1 To n)
    ReDim eOper(1 To n)
    ReDim eComb(1 To n)
    If last < EQ0 Then Exit Sub
    v = ws.Range(ws.Cells(EQ0, 1), ws.Cells(last, EC_ULT)).Value
    For i = 1 To UBound(v, 1)
        cod = Txt(v(i, EC_COD))
        If cod <> "" Then
            If Buscar(eCol, UCase(cod)) = 0 Then
                enEq = enEq + 1
                eCod(enEq) = cod
                eDes(enEq) = Limpio(v(i, EC_DES))
                eMarca(enEq) = Limpio(v(i, EC_MARCA))
                eModelo(enEq) = Limpio(v(i, EC_MODELO))
                eSerie(enEq) = Limpio(v(i, EC_SERIE))
                eAnio(enEq) = Limpio(v(i, EC_ANIO))
                eProv(enEq) = Limpio(v(i, EC_PROV))
                eOS(enEq) = Limpio(v(i, EC_OS))
                eUnd(enEq) = Txt(v(i, EC_UND))
                eCond(enEq) = Limpio(v(i, EC_COND))
                eOper(enEq) = Limpio(v(i, EC_OPER))
                eComb(enEq) = Txt(v(i, EC_COMB))
                eCol.Add enEq, UCase(cod)
            End If
        End If
    Next i
End Sub

Public Sub LeerPartesHoja()
    ' (formato anterior) partes de la hoja PARTES
    Dim ws As Worksheet, last As Long, n As Long, i As Long, k As Long, t As Double
    Dim llave() As Double
    LeerEquipos
    pnPar = 0
    pnMal = 0
    Set ws = Hoja("PARTES")
    last = UltimaFilaVarias(ws, PC_FEC, PC_EQ, PT0, PTMAX)
    n = last - PT0 + 1
    If n < 1 Then n = 1
    ReDim pR(1 To n)
    ReDim pEq(1 To n)
    ReDim pFec(1 To n)
    ReDim pOrden(1 To n)
    ReDim llave(1 To n)
    If last < PT0 Then
        pV = Empty
        Exit Sub
    End If
    pV = ws.Range(ws.Cells(PT0, 1), ws.Cells(last, PC_ULT)).Value
    For i = 1 To UBound(pV, 1)
        If Txt(pV(i, PC_FEC)) <> "" Or Txt(pV(i, PC_EQ)) <> "" Then
            t = Fecha(pV(i, PC_FEC))
            k = Buscar(eCol, UCase(Txt(pV(i, PC_EQ))))
            If t > 0 And k > 0 Then
                pnPar = pnPar + 1
                pR(pnPar) = i
                pEq(pnPar) = k
                pFec(pnPar) = t
                llave(pnPar) = (CDbl(k) * 100000# + t) * 10000# + i
                pOrden(pnPar) = pnPar
            Else
                pnMal = pnMal + 1
            End If
        End If
    Next i
    If pnPar > 1 Then Ordenar llave, pOrden, pnPar
End Sub

Public Sub LeerPartes()
    ' partes del mes desde el cuadro 4.15 (combustible y horas de cada día)
    LeerPartes415
End Sub

Public Function HorasFila(ByVal i As Long) As Double
    ' horas a valorizar de una fila de PARTES: las reales o, si no hay, las del A.C.U.
    If EsNum(pV(i, PC_HREAL)) Then
        HorasFila = Num(pV(i, PC_HREAL))
    ElseIf EsNum(pV(i, PC_MET)) And EsNum(pV(i, PC_REND)) Then
        HorasFila = Redondear(Num(pV(i, PC_MET)) * Num(pV(i, PC_REND)), 2)
    Else
        HorasFila = 0
    End If
End Function

Public Function Hora(ByVal v As Variant) As String
    Hora = ""
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        Hora = Trim(CStr(v))
    ElseIf EsNum(v) Then
        Hora = Format(CDate(Num(v) - Int(Num(v))), "hh:mm")
    End If
End Function

Public Function Tramo(ByVal de As Variant, ByVal a As Variant) As Double
    ' horas entre dos horas (si ambas son números)
    Tramo = 0
    If EsNum(de) And EsNum(a) Then
        If Num(a) > Num(de) Then Tramo = (Num(a) - Num(de)) * 24
    End If
End Function


Public Sub GenerarPartes()
    ' arma los partes en la hoja PARTE_DIARIO (sin PDF: se revisan y se exportan con PDF PARTES DIARIOS)
    Dim n As Long, s As String
    s = ActualizarCuadros(True)
    n = ArmarPartes(False, False)
End Sub


Private Function DescCatalogo(ByVal cod As String) As String
    ' descripción de un código del CATÁLOGO ("" si no está)
    Dim ws As Worksheet, r As Variant
    DescCatalogo = ""
    If cod = "" Then Exit Function
    Set ws = Hoja("CATALOGO")
    On Error Resume Next
    r = Application.Match(cod, ws.Range(ws.Cells(CT0, CC_COD), ws.Cells(CTMAX, CC_COD)), 0)
    If Not IsError(r) Then DescCatalogo = UCase(Txt(ws.Cells(CT0 + CLng(r) - 1, CC_DES).Value))
    On Error GoTo 0
End Function

Private Sub Poner(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal txt As String)
    ' escribe un texto del parte; si no entra en el renglón se reduce la letra
    ws.Cells(r, c).Value = txt
    On Error Resume Next
    ws.Cells(r, c).MergeArea.ShrinkToFit = True
    ws.Cells(r, c).MergeArea.WrapText = False
End Sub

Private Function Partir(ByVal s As String, ByVal lim As Long, ByRef resto As String) As String
    ' corta s en un espacio antes de lim caracteres; devuelve la primera parte y deja el resto
    Dim k As Long
    If Len(s) <= lim Then
        Partir = s
        resto = ""
        Exit Function
    End If
    k = InStrRev(s, " ", lim)
    If k < lim \ 2 Then k = lim
    Partir = RTrim(Left(s, k))
    resto = LTrim(Mid(s, k + 1))
End Function


Public Function ArmarPartes(ByVal conPDF As Boolean, ByVal porEquipo As Boolean) As Long
    ' un parte diario por equipo y día, con el formato PARTE DIARIO DE MAQUINARIA del informe (hoja FMT_PARTE)
    Dim cm As Long, wsF As Worksheet, wsP As Worksheet, j As Long, g0 As Long, g1 As Long, k As Long, x As Long
    Dim r0 As Long, i As Long, c As Long, nG As Long, nLin As Long, nPag As Long, p As Long, li As Long, rr As Long
    Dim desde() As Long, hasta() As Long, carpeta As String, archivo As String, area As String, msg As String
    Dim totH As Double, totC As Double, hIni As Double, hFin As Double, f As Double, hHor As Double
    Dim nReal As Long, nParte As String, oper As String, mDe As String, mA As String, tDe As String, tA As String
    Dim obs As String, fuente As String, nPdf As Long, ok As Boolean, t As String, totPag As Long
    Dim proy As String, cui As String, f3 As String, f4 As String, info As String, combDes As String
    Dim mDeN As Double, mAN As Double, tDeN As Double, tAN As Double
    Dim lineas(1 To 3) As String, resto As String, nl As Long, trab As String, copObj As Boolean, sub2 As String

    ArmarPartes = 0
    Inicio cm, "Armando los partes diarios..."
    LeerPartes
    If pnPar = 0 Then
        Fin cm
        Aviso "No hay combustible ni horas en el cuadro 4.15 del mes (registre los equipos en EQUIPOS, pulse ACTUALIZAR " & _
              "CUADROS y anote cada día el combustible y las horas).", "PARTES DIARIOS", 48
        Exit Function
    End If
    proy = Limpio(Hoja("DATOS").Range(D_PROY).Value)
    cui = Txt(Hoja("DATOS").Range(D_CUI).Value)
    f3 = Txt(Hoja("DATOS").Range(D_F3).Value)
    f4 = Txt(Hoja("DATOS").Range(D_F4).Value)
    Set wsF = Hoja("FMT_PARTE")
    Set wsP = Hoja("PARTE_DIARIO")
    copObj = True
    On Error Resume Next
    copObj = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = True
    wsP.ResetAllPageBreaks
    BorrarFormas wsP
    BorrarFilas wsP
    On Error GoTo 0
    wsP.Cells.Clear
    For c = 1 To PD_COLS
        wsP.Columns(c).ColumnWidth = wsF.Columns(c).ColumnWidth
    Next c
    CopiarPagina wsF, wsP
    ReDim desde(1 To enEq)
    ReDim hasta(1 To enEq)
    r0 = 1
    j = 1
    Do While j <= pnPar
        g0 = j
        k = pEq(pOrden(j))
        f = pFec(pOrden(j))
        Do While j <= pnPar
            If pEq(pOrden(j)) <> k Or pFec(pOrden(j)) <> f Then Exit Do
            j = j + 1
        Loop
        g1 = j - 1
        nG = nG + 1
        If nG Mod 10 = 0 Then Estado "Armando partes diarios: " & nG & " ..."
        nLin = g1 - g0 + 1
        nPag = (nLin + PD_N - 1) \ PD_N
        nParte = ""
        oper = ""
        mDe = ""
        mA = ""
        tDe = ""
        tA = ""
        hIni = 0
        hFin = 0
        hHor = 0
        mDeN = -1: mAN = 0: tDeN = -1: tAN = 0
        totH = 0
        totC = 0
        nReal = 0
        obs = ""
        For x = g0 To g1
            i = pR(pOrden(x))
            If nParte = "" Then nParte = Txt(pV(i, PC_NPARTE))
            If oper = "" Then oper = Limpio(pV(i, PC_OPER))
            ' horario del parte: la hora más temprana y la más tarde de las líneas del día
            If EsNum(pV(i, PC_MDE)) Then
                If mDeN < 0 Or Num(pV(i, PC_MDE)) < mDeN Then mDeN = Num(pV(i, PC_MDE))
            ElseIf mDe = "" Then
                mDe = Hora(pV(i, PC_MDE))
            End If
            If EsNum(pV(i, PC_MA)) Then
                If Num(pV(i, PC_MA)) > mAN Then mAN = Num(pV(i, PC_MA))
            ElseIf mA = "" Then
                mA = Hora(pV(i, PC_MA))
            End If
            If EsNum(pV(i, PC_TDE)) Then
                If tDeN < 0 Or Num(pV(i, PC_TDE)) < tDeN Then tDeN = Num(pV(i, PC_TDE))
            ElseIf tDe = "" Then
                tDe = Hora(pV(i, PC_TDE))
            End If
            If EsNum(pV(i, PC_TA)) Then
                If Num(pV(i, PC_TA)) > tAN Then tAN = Num(pV(i, PC_TA))
            ElseIf tA = "" Then
                tA = Hora(pV(i, PC_TA))
            End If
            hHor = hHor + Tramo(pV(i, PC_MDE), pV(i, PC_MA)) + Tramo(pV(i, PC_TDE), pV(i, PC_TA))
            If EsNum(pV(i, PC_HINI)) Then
                If hIni = 0 Or Num(pV(i, PC_HINI)) < hIni Then hIni = Num(pV(i, PC_HINI))
            End If
            If EsNum(pV(i, PC_HFIN)) Then
                If Num(pV(i, PC_HFIN)) > hFin Then hFin = Num(pV(i, PC_HFIN))
            End If
            totH = totH + HorasFila(i)
            totC = totC + Num(pV(i, PC_COMB))
            If EsNum(pV(i, PC_HREAL)) Then nReal = nReal + 1
            t = Limpio(pV(i, PC_OBS))
            If t <> "" Then
                If InStr(1, obs, t, vbTextCompare) = 0 Then
                    If obs = "" Then
                        obs = t
                    Else
                        obs = obs & "; " & t
                    End If
                End If
            End If
        Next x
        If mDeN >= 0 Then mDe = Hora(mDeN)
        If mAN > 0 Then mA = Hora(mAN)
        If tDeN >= 0 Then tDe = Hora(tDeN)
        If tAN > 0 Then tA = Hora(tAN)
        If totH > 0 And Abs(hHor - totH) <= 0.05 Then hHor = totH   ' horario redondeado al minuto
        If oper = "" Then oper = eOper(k)
        If nReal = nLin Then
            fuente = "PARTE DIARIO (HORAS REALES)"
        ElseIf nReal = 0 Then
            fuente = "A.C.U. x METRADO (ESTIMADO)"
        Else
            fuente = "MIXTA: REAL Y A.C.U."
        End If
        If nReal < nLin Then
            t = "Horas"
            If totC > 0 Then t = t & " y combustible"
            t = t & " estimados según A.C.U. x metrado: reemplazar con la lectura del horómetro"
            If totC > 0 Then t = t & " y el vale de combustible"
            t = t & "."
            If obs = "" Then
                obs = t
            ElseIf Right(obs, 1) = "." Then
                obs = obs & " " & t
            Else
                obs = obs & ". " & t
            End If
        End If
        If nParte <> "" Then obs = "PARTE N° " & nParte & ". " & obs
        info = "CÓDIGO: " & eCod(k) & " (" & eUnd(k) & ")"
        If eOS(k) <> "" Then
            If UCase(Left(eOS(k), 3)) = "O/S" Then
                info = info & "   ·   " & eOS(k)
            Else
                info = info & "   ·   O/S: " & eOS(k)
            End If
        End If
        If eProv(k) <> "" Then info = info & "   ·   PROVEEDOR: " & eProv(k)
        If eCond(k) <> "" Then info = info & "   ·   " & eCond(k)
        combDes = ""
        If totC > 0 Then combDes = DescCatalogo(eComb(k))

        If desde(k) = 0 Then desde(k) = r0
        For p = 1 To nPag
            For i = 1 To PD_ALTO
                wsP.Rows(r0 + i - 1).RowHeight = wsF.Rows(i).RowHeight
            Next i
            wsF.Range(wsF.Cells(1, 1), wsF.Cells(PD_ALTO, PD_COLS)).Copy Destination:=wsP.Cells(r0, 1)
            ' encabezado
            wsP.Cells(r0 + 2, 5).Value = "OBRA: “" & proy & "”- CUI N° " & cui
            wsP.Cells(r0 + 4, 13).Value = Format(CDate(f), "dd")
            wsP.Cells(r0 + 4, 14).Value = Format(CDate(f), "mm")
            wsP.Cells(r0 + 4, 15).Value = Format(CDate(f), "yyyy")
            If oper <> "" Then Poner wsP, r0 + 5, 3, "NOMBRE DEL OPERADOR: " & oper
            If eDes(k) <> "" Then Poner wsP, r0 + 6, 3, "EQUIPO: " & eDes(k)
            If eAnio(k) <> "" Then Poner wsP, r0 + 6, 11, "CAPACIDAD: " & eAnio(k)
            t = eMarca(k)
            If eModelo(k) <> "" Then t = Trim(t & "   MODELO: " & eModelo(k))
            If t <> "" Then Poner wsP, r0 + 7, 3, "MARCA: " & t
            If eSerie(k) <> "" Then Poner wsP, r0 + 7, 11, "PLACA: " & eSerie(k)
            ' horario y horas de trabajo
            If mDe <> "" Then wsP.Cells(r0 + 10, 10).Value = "De: " & mDe
            If mA <> "" Then wsP.Cells(r0 + 11, 10).Value = "A: " & mA
            If tDe <> "" Then wsP.Cells(r0 + 10, 12).Value = "De: " & tDe
            If tA <> "" Then wsP.Cells(r0 + 11, 12).Value = "A: " & tA
            If hHor > 0 Then wsP.Cells(r0 + 9, 14).Value = "HORAS DE TRABAJO:" & vbLf & Format(Redondear(hHor, 2), "0.00")
            ' trabajos realizados
            For li = 1 To PD_N
                x = g0 + (p - 1) * PD_N + li - 1
                If x > g1 Then Exit For
                i = pR(pOrden(x))
                rr = r0 + PD_L0 + li - 2
                trab = Limpio(pV(i, PC_TRAB))
                If Txt(pV(i, PC_PART)) <> "" Then trab = Txt(pV(i, PC_PART)) & " – " & trab
                If EsNum(pV(i, PC_MET)) Then
                    t = Format(Num(pV(i, PC_MET)), "#,##0.00") & " " & Txt(pV(i, PC_UND))
                    If EsNum(pV(i, PC_REND)) Then t = t & " x " & Format(Num(pV(i, PC_REND)), "0.0000") & " h/" & Txt(pV(i, PC_UND))
                    trab = trab & "  (" & Trim(t) & ")"
                End If
                wsP.Cells(rr, 7).Value = trab
                If UBound(pV, 2) >= PC_ULT + 2 Then          ' horas DE / A (autollenado)
                    If Not IsEmpty(pV(i, PC_ULT + 1)) Then
                        wsP.Cells(rr, 3).NumberFormat = "hh:mm"
                        wsP.Cells(rr, 5).NumberFormat = "hh:mm"
                        If EsNum(pV(i, PC_ULT + 1)) Then wsP.Cells(rr, 3).Value = Num(pV(i, PC_ULT + 1))
                        If EsNum(pV(i, PC_ULT + 2)) Then wsP.Cells(rr, 5).Value = Num(pV(i, PC_ULT + 2))
                    End If
                End If
                AjustarTexto wsP.Cells(rr, 7), 140, 140, 10
                wsP.Cells(rr, 14).Value = HorasFila(i)
            Next li
            ' horómetro
            If hIni > 0 Then wsP.Cells(r0 + 27, 12).Value = hIni
            If hFin > 0 Then wsP.Cells(r0 + 27, 14).Value = hFin
            If p = nPag Then
                wsP.Cells(r0 + 23, 14).Value = Redondear(totH, 2)
                wsP.Cells(r0 + 26, 3).Value = "Horas trabajadas: (H.T.)   " & Format(Redondear(totH, 2), "0.00") & " h"
                ' observaciones (3 renglones)
                lineas(1) = Partir("OBSERVACIONES: " & obs, 88, resto)
                lineas(2) = ""
                lineas(3) = ""
                If resto <> "" Then
                    lineas(2) = Partir(resto, 98, resto)
                    lineas(3) = info
                Else
                    lineas(2) = info
                End If
                If resto <> "" Then lineas(3) = resto & "   ·   " & info
                For nl = 1 To 3
                    If lineas(nl) <> "" Then Poner wsP, r0 + 29 + nl, 3, lineas(nl)
                Next nl
                ' combustible
                If totC > 0 Then
                    If InStr(1, combDes, "GASOL", vbTextCompare) > 0 Then
                        wsP.Cells(r0 + 36, 3).Value = "GASOLINA: " & Format(Redondear(totC, 2), "0.00") & " GAL"
                    Else
                        wsP.Cells(r0 + 35, 3).Value = "PETROLEO: " & Format(Redondear(totC, 2), "0.00") & " GAL"
                    End If
                End If
            Else
                Poner wsP, r0 + 30, 3, "OBSERVACIONES: continúa en la hoja siguiente."
                Poner wsP, r0 + 31, 3, info
            End If
            wsP.Cells(r0 + 48, 9).Value = f3
            wsP.Cells(r0 + 48, 13).Value = f4
            If r0 > 1 Then
                On Error Resume Next
                wsP.HPageBreaks.Add Before:=wsP.Cells(r0, 1)
                On Error GoTo 0
            End If
            r0 = r0 + PD_ALTO
            totPag = totPag + 1
        Next p
        hasta(k) = r0 - 1
    Loop
    On Error Resume Next
    Application.CutCopyMode = False
    Application.CopyObjectsWithCells = copObj
    On Error GoTo 0

    For k = 1 To enEq
        If desde(k) > 0 Then MarcarBloque wsP, desde(k), eCod(k) & " " & eDes(k), hasta(k)
    Next k
    OcultarMarcas wsP
    area = "$B$1:$P$" & (r0 - 1)
    On Error Resume Next
    wsP.PageSetup.PrintArea = area
    On Error GoTo 0
    Fin cm
    PaginaFormato wsP, "PARTES"                   ' cada parte entra completo en la hoja, con márgenes
    ArmarPartes = nG

    msg = "Se armaron " & nG & " partes diarios (" & totPag & " hojas) del mes de " & MesTexto() & "."
    If pnMal > 0 Then msg = msg & vbCrLf & "Atención: " & pnMal & " equipo(s) del 4.15 ya no están en EQUIPOS (pulse ACTUALIZAR CUADROS)."
    If nG > 0 And Not gAcum Then msg = msg & vbCrLf & vbCrLf & "Revise (y corrija si hace falta) la hoja PARTE_DIARIO; luego use el botón " & _
                       "PDF PARTES DIARIOS (A5, o A4 con dos por cara)."
    If Not gSilencio Then
        wsP.Activate
        wsP.Range("A1").Select
    End If
    Aviso msg, "PARTES DIARIOS"
    If conPDF And nG > 0 Then msg = ExportarFormato("PARTES")
End Function
