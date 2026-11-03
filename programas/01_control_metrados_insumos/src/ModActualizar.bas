Attribute VB_Name = "ModActualizar"
Option Explicit

'==========================================================================================
'  CONTROL DE METRADOS DIARIOS - ACTUALIZAR TODO
'  Arma las planillas diarias con sus títulos por nivel, calcula los insumos según el A.C.U.
'  frente a los ingresos, las horas de maquinaria y el combustible, las HH de mano de obra y
'  las alertas (sin O/S ni ingreso, rendimiento, metrado del expediente).
'==========================================================================================

' ---- RESUMEN
Private nr As Long
Private rTipo() As String, rItem() As String, rKey() As String, rDesc() As String, rUnd() As String
Private rMet() As Double, rPU() As Double, rAnt() As Double, rMaxD() As Double
Private rLvl() As Long, rLast() As Long
Private rVal() As Double
Private dF(1 To 31) As Double
Private fIni As Double, fFin As Double, fMes As Double, fFinMes As Double, umbral As Double
Private sFirma1 As String, sFirma2 As String, sFirma3 As String
' ---- INSUMOS
Private nI As Long
Private iNom() As String, iUnd() As String, iCtrl() As String, iPre() As Double
Private iAny() As Double, iConf() As Double, iVig() As Double, iQty() As Double, iMaxFin() As Double
Private iOpen() As Boolean, iDated() As Boolean, iBlank() As Boolean, iFac() As Boolean
Private iCons() As Double, iAnt() As Double
' ---- ACU
Private nA As Long
Private aR() As Long, aK() As Long, aCat() As Long, aSig() As Long
Private aMO() As Boolean, aPct() As Boolean
Private aNom() As String
Private aApo() As Double, aPre() As Double, aSh() As Double, aL() As Double, aM() As Double, aN() As Double
Private aAnt() As Double, aMes() As Double
Private pIni() As Long
Private pAny() As Double, pConf() As Double
' ---- categorías de mano de obra
Private nc As Long
Private cNom() As String
Private colC As Collection
' ---- resultados
Private miss() As Double, pcf() As Double
Private hh() As Double
Private colR As Collection, colI As Collection


'========================================================================== BOTONES
Public Sub ActualizarTodo()
    Dim msg As String
    msg = Ejecutar()
    If Len(msg) > 0 Then MsgBox msg, IIf(Left(msg, 8) = "No se pu", vbExclamation, vbInformation), "Actualizar todo"
End Sub

Public Sub ActualizarSilencioso()
    Dim msg As String
    gSilencio = True
    msg = Ejecutar()
    gSilencio = False
    ThisWorkbook.Worksheets("DATOS").Range(D_MSGACT).Value = msg
End Sub

Public Sub PlanillaDelDia()
    Dim d As Long, c As Long, s As String, msg As String, ws As Worksheet
    d = 0
    On Error Resume Next
    If ActiveSheet.Name = "RESUMEN" Then
        c = ActiveCell.Column
        If c >= CD1 And c < CD1 + ND Then d = c - CD1 + 1
    ElseIf Len(ActiveSheet.Name) = 2 And IsNumeric(ActiveSheet.Name) Then
        d = CLng(ActiveSheet.Name)
    End If
    On Error GoTo 0
    If d = 0 Then
        s = InputBox("Día del mes (1 a 31) de la planilla que desea ver:" & vbCrLf & vbCrLf & _
                     "Sugerencia: en el RESUMEN seleccione una celda del día y presione este botón.", "Ver planilla del día")
        If Len(Trim(s)) = 0 Then Exit Sub
        If Not IsNumeric(s) Then Exit Sub
        d = CLng(s)
        If d < 1 Or d > ND Then Exit Sub
    End If
    msg = Ejecutar()
    If Left(msg, 8) = "No se pu" Then
        MsgBox msg, vbExclamation, "Ver planilla del día"
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(Format(d, "00"))
    If ws.Visible <> XL_VISIBLE Then
        MsgBox "El día " & d & " no tiene metrados registrados en el RESUMEN.", vbInformation, "Ver planilla del día"
    Else
        ws.Activate
        ws.Range("A1").Select
    End If
End Sub


'========================================================================== PROCESO
Public Function Ejecutar() As String
    Dim t0 As Double, calcMode As Long, huella As Variant, nPl As Long, nRojo As Long, nSin As Long
    t0 = Timer
    Inicio calcMode, "Actualizando metrados, insumos y planillas..."
    On Error GoTo Fallo
    huella = ThisWorkbook.Worksheets("DATOS").Range(D_HUELLAHOY).Value
    LeerDatos
    LeerResumen
    LeerInsumos
    LeerIngresos
    LeerAcu
    Calcular
    EscribirResumen
    EscribirAux
    EscribirInsumos nSin
    EscribirAcu
    EscribirMaquinaria
    EscribirManoObra
    GenerarPlanillas nPl, nRojo
    With ThisWorkbook.Worksheets("DATOS")
        .Range(D_ULTACT).NumberFormat = "@"
        .Range(D_ULTACT).Value = Format(Now, "dd/mm/yyyy hh:mm")
        .Range(D_HUELLA).Value = huella
    End With
    Ejecutar = "Actualizado en " & Format(Timer - t0, "0.0") & " s." & vbCrLf & vbCrLf & _
               nPl & " planilla(s) diaria(s) con metrados." & vbCrLf & _
               nRojo & " partida(s)-día en ROJO (no valorizar): vea la columna ALERTAS del RESUMEN." & vbCrLf & _
               nSin & " insumo(s) con consumo y sin ingreso ni O/S (hoja INSUMOS)."
Salir:
    Fin calcMode
    Exit Function
Fallo:
    Ejecutar = "No se pudo completar la actualización." & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    Resume Salir
End Function


'========================================================================== LECTURA
Private Sub LeerDatos()
    Dim ws As Worksheet, d As Long, x As Double
    Set ws = ThisWorkbook.Worksheets("DATOS")
    fIni = Num(ws.Range(D_INICIO).Value)
    fFin = fIni + Num(ws.Range(D_PLAZO).Value) - 1
    fMes = Num(ws.Range(D_MES).Value)
    umbral = Num(ws.Range(D_UMBRAL).Value)
    sFirma1 = Txt(ws.Range(D_FIRMA1).Value)
    sFirma2 = Txt(ws.Range(D_FIRMA2).Value)
    sFirma3 = Txt(ws.Range(D_FIRMA3).Value)
    If fMes <= 0 Then Err.Raise 1001, , "Falta el MES QUE SE CONTROLA en la hoja DATOS."
    If fIni <= 0 Then Err.Raise 1001, , "Falta la FECHA DE INICIO DE OBRA en la hoja DATOS."
    fMes = CDbl(DateSerial(Year(fMes), Month(fMes), 1))
    fFinMes = CDbl(DateSerial(Year(fMes), Month(fMes) + 1, 1)) - 1
    For d = 1 To ND
        x = CDbl(DateSerial(Year(fMes), Month(fMes), d))
        If x <= fFinMes Then
            dF(d) = x
        Else
            dF(d) = 0
        End If
    Next d
End Sub

Private Sub LeerResumen()
    Dim ws As Worksheet, v As Variant, i As Long, j As Long, d As Long, cU As Double, k As String
    Set ws = ThisWorkbook.Worksheets("RESUMEN")
    v = ws.Range(ws.Cells(RR0, 2), ws.Cells(RRMAX, 2)).Value
    nr = 0
    For i = 1 To UBound(v, 1)
        If Txt(v(i, 1)) = "" Then Exit For
        nr = i
    Next i
    If nr = 0 Then Err.Raise 1002, , "El RESUMEN no tiene partidas: cargue el PRESUPUESTO y el A.C.U. y presione PREPARAR OBRA."
    v = ws.Range(ws.Cells(RR0, 1), ws.Cells(RR0 + nr - 1, 47)).Value
    ReDim rTipo(1 To nr): ReDim rItem(1 To nr): ReDim rKey(1 To nr): ReDim rDesc(1 To nr): ReDim rUnd(1 To nr)
    ReDim rMet(1 To nr): ReDim rPU(1 To nr): ReDim rAnt(1 To nr): ReDim rMaxD(1 To nr)
    ReDim rLvl(1 To nr): ReDim rLast(1 To nr)
    ReDim rVal(1 To nr, 1 To ND)
    Set colR = New Collection
    For i = 1 To nr
        rTipo(i) = UCase(Txt(v(i, 1)))
        rItem(i) = Txt(v(i, 2))
        rKey(i) = Clave(rItem(i))
        rDesc(i) = Txt(v(i, 3))
        rUnd(i) = Txt(v(i, 4))
        rMet(i) = Num(v(i, 5))
        rPU(i) = Num(v(i, 6))
        rAnt(i) = Num(v(i, 8))
        cU = Num(v(i, 47))
        If cU <= 0 Then cU = 1
        rMaxD(i) = Num(v(i, 46)) * cU
        rLvl(i) = Nivel(rItem(i))
        For d = 1 To ND
            rVal(i, d) = Num(v(i, CD1 + d - 1))
            If rVal(i, d) < 0 Then rVal(i, d) = 0
        Next d
        If rKey(i) <> "" Then
            If Buscar(colR, "R" & rKey(i)) = 0 Then colR.Add i, "R" & rKey(i)
        End If
    Next i
    For i = 1 To nr
        rLast(i) = i
        If rTipo(i) = "T" Then
            k = rKey(i) & "."
            j = i + 1
            Do While j <= nr
                If Left(rKey(j), Len(k)) <> k Then Exit Do
                rLast(i) = j
                j = j + 1
            Loop
        End If
    Next i
End Sub

Private Sub LeerInsumos()
    Dim ws As Worksheet, v As Variant, i As Long
    Set ws = ThisWorkbook.Worksheets("INSUMOS")
    v = ws.Range(ws.Cells(IN0, 2), ws.Cells(INMAX, 2)).Value
    nI = 0
    For i = 1 To UBound(v, 1)
        If Txt(v(i, 1)) = "" Then Exit For
        nI = i
    Next i
    If nI = 0 Then
        ReDim iNom(0 To 0)
        Set colI = New Collection
        Exit Sub
    End If
    v = ws.Range(ws.Cells(IN0, 1), ws.Cells(IN0 + nI - 1, 6)).Value
    ReDim iNom(1 To nI): ReDim iUnd(1 To nI): ReDim iCtrl(1 To nI): ReDim iPre(1 To nI)
    ReDim iAny(1 To nI): ReDim iConf(1 To nI): ReDim iVig(1 To nI): ReDim iQty(1 To nI): ReDim iMaxFin(1 To nI)
    ReDim iOpen(1 To nI): ReDim iDated(1 To nI): ReDim iBlank(1 To nI): ReDim iFac(1 To nI)
    ReDim iCons(1 To nI, 1 To ND): ReDim iAnt(1 To nI)
    Set colI = New Collection
    For i = 1 To nI
        iNom(i) = Txt(v(i, 2))
        iUnd(i) = Txt(v(i, 3))
        iPre(i) = Num(v(i, 5))
        iCtrl(i) = UCase(Txt(v(i, 6)))
        iAny(i) = SENT
        iConf(i) = SENT
        iVig(i) = SENT
        If Buscar(colI, "I" & UCase(iNom(i))) = 0 Then colI.Add i, "I" & UCase(iNom(i))
    Next i
End Sub

Private Sub LeerIngresos()
    Dim ws As Worksheet, v As Variant, g As Long, k As Long, p As Long
    Dim fe As Double, fi As Double, fa As Double, est As String, unDoc As String, ins As String, par As String
    Set ws = ThisWorkbook.Worksheets("INGRESOS")
    v = ws.Range(ws.Cells(IG0, 1), ws.Cells(IG1, 15)).Value
    ReDim pAny(1 To nr): ReDim pConf(1 To nr)
    For p = 1 To nr
        pAny(p) = SENT
        pConf(p) = SENT
    Next p
    For g = 1 To UBound(v, 1)
        fe = Num(v(g, 2))
        unDoc = UCase(Txt(v(g, 7)))
        ins = Txt(v(g, 9))
        par = Txt(v(g, 13))
        fi = Num(v(g, 14))
        est = UCase(Txt(v(g, 15)))
        k = 0
        p = 0
        If ins <> "" And nI > 0 Then k = Buscar(colI, "I" & UCase(ins))
        If par <> "" Then p = Buscar(colR, "R" & Clave(par))
        If k > 0 And fe > 0 Then
            fa = 1
            If Txt(v(g, 11)) <> "" Then fa = Num(v(g, 11))
            If fe < iAny(k) Then iAny(k) = fe
            If est = "CONFIRMADO" And fe < iConf(k) Then iConf(k) = fe
            iDated(k) = True
            If fi > 0 Then
                If fi > iMaxFin(k) Then iMaxFin(k) = fi
            Else
                iOpen(k) = True
            End If
            If Txt(v(g, 8)) = "" Then
                iBlank(k) = True
            ElseIf fe <= fFinMes Then
                iQty(k) = iQty(k) + Num(v(g, 8)) * fa
            End If
            If unDoc <> "" And Txt(v(g, 11)) = "" And unDoc <> UCase(iUnd(k)) Then iFac(k) = True
        End If
        If p > 0 And fe > 0 Then
            If fe < pAny(p) Then pAny(p) = fe
            If est = "CONFIRMADO" And fe < pConf(p) Then pConf(p) = fe
        End If
    Next g
    For k = 1 To nI
        If iDated(k) And (Not iOpen(k)) And iMaxFin(k) > 0 Then
            iVig(k) = iMaxFin(k)
        Else
            iVig(k) = SENT
        End If
    Next k
End Sub

Private Sub LeerAcu()
    Dim ws As Worksheet, v As Variant, j As Long, r As Long, tP As String, nm As String
    Dim lastJ() As Long
    Set ws = ThisWorkbook.Worksheets("ACU")
    v = ws.Range(ws.Cells(AC0, 1), ws.Cells(ACMAX, 1)).Value
    nA = 0
    For j = 1 To UBound(v, 1)
        If Txt(v(j, 1)) = "" Then Exit For
        nA = j
    Next j
    ReDim pIni(1 To nr): ReDim lastJ(1 To nr)
    nc = 0
    ReDim cNom(1 To 1)
    Set colC = New Collection
    If nA = 0 Then Exit Sub
    v = ws.Range(ws.Cells(AC0, 1), ws.Cells(AC0 + nA - 1, 6)).Value
    ReDim aR(1 To nA): ReDim aK(1 To nA): ReDim aCat(1 To nA): ReDim aSig(1 To nA)
    ReDim aMO(1 To nA): ReDim aPct(1 To nA): ReDim aNom(1 To nA): ReDim aApo(1 To nA): ReDim aPre(1 To nA)
    For j = 1 To nA
        aR(j) = Buscar(colR, "R" & Clave(Txt(v(j, 1))))
        aNom(j) = Txt(v(j, 2))
        tP = UCase(Txt(v(j, 4)))
        aMO(j) = (tP = "MO")
        aPct(j) = (tP = "%MO")
        aApo(j) = Num(v(j, 5))
        aPre(j) = Num(v(j, 6))
        If aMO(j) Then
            nm = UCase(aNom(j))
            aCat(j) = Buscar(colC, "C" & nm)
            If aCat(j) = 0 Then
                nc = nc + 1
                ReDim Preserve cNom(1 To nc)
                cNom(nc) = nm
                colC.Add nc, "C" & nm
                aCat(j) = nc
            End If
        ElseIf tP = "MAT" Or tP = "EQ" Or tP = "SC" Then
            If nI > 0 Then aK(j) = Buscar(colI, "I" & UCase(aNom(j)))
            r = aR(j)
            If r > 0 Then
                If pIni(r) = 0 Then
                    pIni(r) = j
                Else
                    aSig(lastJ(r)) = j
                End If
                lastJ(r) = j
            End If
        End If
    Next j
End Sub


'========================================================================== CÁLCULO
Private Sub Calcular()
    Dim j As Long, r As Long, k As Long, d As Long, t As Double, s As Double, q As Double
    ReDim miss(1 To nr, 1 To ND): ReDim pcf(1 To nr, 1 To ND)
    If nc > 0 Then
        ReDim hh(1 To nc, 1 To ND)
    Else
        ReDim hh(1 To 1, 1 To ND)
    End If
    If nA > 0 Then
        ReDim aSh(1 To nA): ReDim aL(1 To nA): ReDim aM(1 To nA): ReDim aN(1 To nA)
        ReDim aAnt(1 To nA): ReDim aMes(1 To nA)
    End If
    For j = 1 To nA
        r = aR(j)
        k = aK(j)
        aSh(j) = 0
        aL(j) = SENT
        aM(j) = SENT
        aN(j) = SENT
        If (Not aMO(j)) And r > 0 And k > 0 Then
            If iCtrl(k) <> "NO" And rPU(r) > 0 Then aSh(j) = Redondear(aApo(j) * aPre(j) / rPU(r), 4)
            aL(j) = Minimo(iAny(k), pAny(r))
            aM(j) = Minimo(iConf(k), pConf(r))
            If pAny(r) < SENT Then
                aN(j) = SENT
            Else
                aN(j) = iVig(k)
            End If
        End If
    Next j
    For r = 1 To nr
        For d = 1 To ND
            miss(r, d) = -1
            pcf(r, d) = -1
            If rTipo(r) = "P" And rVal(r, d) > 0 Then
                t = dF(d)
                s = 0
                q = 0
                j = pIni(r)
                Do While j > 0
                    If t = 0 Then
                        s = s + aSh(j)
                    ElseIf aL(j) > t Or aN(j) < t Then
                        s = s + aSh(j)
                    ElseIf aM(j) > t Then
                        q = q + aSh(j)
                    End If
                    j = aSig(j)
                Loop
                miss(r, d) = s
                pcf(r, d) = q
            End If
        Next d
    Next r
    If nI > 0 Then
        ReDim iCons(1 To nI, 1 To ND): ReDim iAnt(1 To nI)
    End If
    For j = 1 To nA
        r = aR(j)
        If r > 0 And (Not aPct(j)) Then
            aAnt(j) = aApo(j) * rAnt(r)
            For d = 1 To ND
                If rVal(r, d) > 0 Then
                    q = aApo(j) * rVal(r, d)
                    aMes(j) = aMes(j) + q
                    If aMO(j) Then
                        If aCat(j) > 0 Then hh(aCat(j), d) = hh(aCat(j), d) + q
                    ElseIf aK(j) > 0 Then
                        iCons(aK(j), d) = iCons(aK(j), d) + q
                    End If
                End If
            Next d
            If (Not aMO(j)) And aK(j) > 0 Then iAnt(aK(j)) = iAnt(aK(j)) + aAnt(j)
        End If
    Next j
End Sub


'========================================================================== ESCRITURA
Private Sub EscribirResumen()
    Dim ws As Worksheet, o() As Variant, r As Long, d As Long
    Dim f1 As Double, f2 As Double, mx As Double, mv As Double, sm As Double, ac As Double
    Dim fuera As Boolean, partes As String, lst As String
    Set ws = ThisWorkbook.Worksheets("RESUMEN")
    ReDim o(1 To nr, 1 To 4)
    For r = 1 To nr
        If rTipo(r) = "P" Then
            f1 = 0
            f2 = 0
            mx = -1
            mv = 0
            sm = 0
            fuera = False
            For d = 1 To ND
                If rVal(r, d) > 0 Then
                    sm = sm + rVal(r, d)
                    If rVal(r, d) > mv Then mv = rVal(r, d)
                    If miss(r, d) > mx Then mx = miss(r, d)
                    If dF(d) = 0 Then
                        fuera = True
                    Else
                        If dF(d) < fIni Or dF(d) > fFin Then fuera = True
                        If f1 = 0 Or dF(d) < f1 Then f1 = dF(d)
                        If dF(d) > f2 Then f2 = dF(d)
                    End If
                End If
            Next d
            If sm > 0 Then
                o(r, 1) = mx
                If f1 > 0 Then
                    o(r, 3) = f1
                    o(r, 4) = f2
                End If
                partes = ""
                If fuera Then partes = Unir(partes, "FUERA DEL PLAZO")
                lst = ListaMes(r, f1, f2, 1, 150)
                If lst <> "" Then partes = Unir(partes, "SIN O/S NI INGRESO: " & lst)
                lst = ListaMes(r, f1, f2, 2, 120)
                If lst <> "" Then partes = Unir(partes, "POR CONFIRMAR: " & lst)
                If rMaxD(r) > 0 And mv > rMaxD(r) * 1.0001 Then
                    partes = Unir(partes, "SUPERA RENDIMIENTO (MÁX. " & Format(rMaxD(r), "0.00") & " " & rUnd(r) & "/DÍA)")
                End If
                ac = rAnt(r) + sm
                If ac > rMet(r) * 1.0001 + 0.001 Then partes = Unir(partes, "SUPERA EL METRADO DEL EXP. TÉC.")
                If partes <> "" Then o(r, 2) = partes
            End If
        End If
    Next r
    ws.Range(ws.Cells(RR0, 49), ws.Cells(RRMAX, 52)).ClearContents
    ws.Range(ws.Cells(RR0, 49), ws.Cells(RR0 + nr - 1, 52)).Value = o
End Sub

Private Sub EscribirAux()
    Dim ws As Worksheet, o() As Variant, r As Long, d As Long
    Set ws = ThisWorkbook.Worksheets("AUX")
    ReDim o(1 To nr, 1 To 94)
    For r = 1 To nr
        o(r, 1) = rItem(r)
        If rTipo(r) = "P" Then
            For d = 1 To ND
                If rVal(r, d) > 0 Then
                    o(r, 1 + d) = rVal(r, d)
                    o(r, 32 + d) = miss(r, d)
                    o(r, 63 + d) = pcf(r, d)
                End If
            Next d
        End If
    Next r
    ws.Range(ws.Cells(RR0, 1), ws.Cells(RRMAX, 94)).ClearContents
    ws.Range(ws.Cells(RR0, 1), ws.Cells(RRMAX, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(RR0, 1), ws.Cells(RR0 + nr - 1, 94)).Value = o
End Sub

Private Sub EscribirInsumos(ByRef nSin As Long)
    Dim ws As Worksheet, o() As Variant, w() As Variant, k As Long, d As Long
    Dim mes As Double, acu As Double, sal As Double, est As String
    nSin = 0
    Set ws = ThisWorkbook.Worksheets("INSUMOS")
    ws.Range(ws.Cells(IN0, 7), ws.Cells(INMAX, 47)).ClearContents
    ws.Range(ws.Cells(IN0, 49), ws.Cells(INMAX, 49)).ClearContents
    If nI = 0 Then Exit Sub
    ReDim o(1 To nI, 1 To 41)
    ReDim w(1 To nI, 1 To 1)
    For k = 1 To nI
        o(k, 1) = iAny(k)
        o(k, 2) = iConf(k)
        o(k, 3) = iVig(k)
        o(k, 4) = iQty(k)
        o(k, 5) = iAnt(k)
        mes = 0
        For d = 1 To ND
            If iCons(k, d) <> 0 Then o(k, 5 + d) = iCons(k, d)
            mes = mes + iCons(k, d)
        Next d
        acu = iAnt(k) + mes
        o(k, 37) = mes
        o(k, 38) = acu
        o(k, 40) = Redondear(mes * iPre(k), 2)
        If acu <> 0 Then
            sal = iQty(k) - acu
            o(k, 39) = sal
            If iCtrl(k) = "NO" Then
                est = "NO SE CONTROLA"
            ElseIf iAny(k) >= SENT Then
                est = "SIN INGRESO NI O/S"
                nSin = nSin + 1
            ElseIf iBlank(k) Then
                est = "REGISTRAR CANTIDAD"
            ElseIf iFac(k) Then
                est = "INDICAR FACTOR EN INGRESOS"
            ElseIf sal < -0.005 Then
                est = "INGRESO INSUFICIENTE"
            ElseIf iConf(k) >= SENT Then
                est = "POR CONFIRMAR"
            Else
                est = "CONFORME"
            End If
            o(k, 41) = est
            w(k, 1) = "SI"
        End If
    Next k
    ws.Range(ws.Cells(IN0, 7), ws.Cells(IN0 + nI - 1, 47)).Value = o
    ws.Range(ws.Cells(IN0, 49), ws.Cells(IN0 + nI - 1, 49)).Value = w
End Sub

Private Sub EscribirAcu()
    Dim ws As Worksheet, o() As Variant, j As Long
    Set ws = ThisWorkbook.Worksheets("ACU")
    ws.Range(ws.Cells(AC0, 7), ws.Cells(ACMAX, 16)).ClearContents
    If nA = 0 Then Exit Sub
    ReDim o(1 To nA, 1 To 10)
    For j = 1 To nA
        If aPct(j) Then
            o(j, 1) = Redondear(aApo(j) * aPre(j) / 100, 6)
        Else
            o(j, 1) = Redondear(aApo(j) * aPre(j), 6)
        End If
        If aMO(j) Then
            o(j, 3) = "MANO DE OBRA"
        ElseIf aK(j) = 0 Or aR(j) = 0 Then
            o(j, 3) = "NO SE CONTROLA"
        Else
            o(j, 2) = aSh(j)
            If aSh(j) = 0 Then
                o(j, 3) = "NO SE CONTROLA"
            ElseIf aSh(j) >= umbral Then
                o(j, 3) = "PRINCIPAL"
            Else
                o(j, 3) = "MENOR"
            End If
            o(j, 4) = aL(j)
            o(j, 5) = aM(j)
            o(j, 6) = aN(j)
        End If
        If Not aPct(j) Then
            o(j, 7) = aAnt(j)
            o(j, 8) = aMes(j)
            o(j, 9) = aAnt(j) + aMes(j)
        End If
        If aR(j) > 0 Then
            o(j, 10) = rDesc(aR(j))
        Else
            o(j, 10) = "(la partida no figura en el RESUMEN)"
        End If
    Next j
    ws.Range(ws.Cells(AC0, 7), ws.Cells(AC0 + nA - 1, 16)).Value = o
End Sub

Private Sub EscribirMaquinaria()
    Dim ws As Worksheet, v As Variant, o() As Variant, r As Long, k As Long, d As Long
    Set ws = ThisWorkbook.Worksheets("MAQ_COMB")
    v = ws.Range(ws.Cells(1, 1), ws.Cells(MQMAX, 4)).Value
    For r = MQ0 To MQMAX
        If Txt(v(r, 4)) = LB_ACU Then
            k = 0
            If nI > 0 Then k = Buscar(colI, "I" & UCase(Txt(v(r, 2))))
            ws.Range(ws.Cells(r, 5), ws.Cells(r, 38)).ClearContents
            If k > 0 Then
                ReDim o(1 To 1, 1 To 34)
                o(1, 1) = iQty(k)
                o(1, 2) = iAny(k)
                o(1, 3) = iVig(k)
                For d = 1 To ND
                    If iCons(k, d) <> 0 Then o(1, 3 + d) = iCons(k, d)
                Next d
                ws.Range(ws.Cells(r, 5), ws.Cells(r, 38)).Value = o
            End If
        End If
    Next r
End Sub

Private Sub EscribirManoObra()
    Dim ws As Worksheet, v As Variant, o() As Variant, r As Long, c As Long, d As Long
    Set ws = ThisWorkbook.Worksheets("MANO_OBRA")
    v = ws.Range(ws.Cells(1, 1), ws.Cells(MOMAX, 4)).Value
    For r = MO0 To MOMAX
        If Txt(v(r, 4)) = LB_HHACU Then
            c = Buscar(colC, "C" & UCase(Txt(v(r, 2))))
            ws.Range(ws.Cells(r, 8), ws.Cells(r, 38)).ClearContents
            If c > 0 Then
                ReDim o(1 To 1, 1 To ND)
                For d = 1 To ND
                    If hh(c, d) <> 0 Then o(1, d) = hh(c, d)
                Next d
                ws.Range(ws.Cells(r, 8), ws.Cells(r, 38)).Value = o
            End If
        End If
    Next r
End Sub


'========================================================================== PLANILLAS
Private Sub GenerarPlanillas(ByRef nPl As Long, ByRef nRojo As Long)
    Dim d As Long
    nPl = 0
    nRojo = 0
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo 0
    For d = 1 To ND
        Planilla d, nPl, nRojo
    Next d
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

Private Sub Planilla(ByVal d As Long, ByRef nPl As Long, ByRef nRojo As Long)
    Dim ws As Worksheet, cnt() As Long, filas() As Long, stt() As Long, o() As Variant
    Dim r As Long, n As Long, k As Long, dd As Long, sig As Long, fila As Long, ultima As Long
    Dim t As Double, tot As Double, ac As Double, nP As Long, nNo As Long, nOk As Long
    Dim fuera As Boolean, est As String
    Set ws = ThisWorkbook.Worksheets(Format(d, "00"))
    t = dF(d)
    ReDim cnt(0 To nr)
    For r = 1 To nr
        cnt(r) = cnt(r - 1)
        If rTipo(r) = "P" And rVal(r, d) > 0 Then cnt(r) = cnt(r) + 1
    Next r
    LimpiarPlanilla ws
    If t > 0 Then
        ws.Range("A2").Value = "PLANILLA DE METRADOS DIARIOS N° " & Format(d, "00") & "  –  " & Format(CDate(t), "dd/mm/yyyy")
        ws.Range("C5").Value = t
        ws.Range("D5").Value = NombreDia(t)
    Else
        ws.Range("A2").Value = "PLANILLA DE METRADOS DIARIOS N° " & Format(d, "00")
        ws.Range("C5").Value = "(no corresponde al mes)"
        ws.Range("D5").Value = Empty
    End If
    fuera = (t = 0)
    If Not fuera Then fuera = (t < fIni Or t > fFin)
    If fuera Then
        ws.Range("I5").Value = "FUERA DEL PLAZO"
    Else
        ws.Range("I5").Value = t - fIni + 1
    End If
    If cnt(nr) = 0 Then
        ws.Range("K5").Value = "SIN METRADOS REGISTRADOS"
        ws.Range("D6").Value = 0
        ws.Range("I6").Value = 0
        On Error Resume Next
        ws.PageSetup.PrintArea = "$A$1:$K$" & (PL0 + 5)
        If ws.Visible = XL_VISIBLE And ActiveSheet.Name <> ws.Name Then ws.Visible = XL_HIDDEN
        On Error GoTo 0
        Exit Sub
    End If
    ReDim filas(1 To nr)
    n = 0
    For r = 1 To nr
        If rTipo(r) = "P" Then
            If rVal(r, d) > 0 Then
                n = n + 1
                filas(n) = r
            End If
        ElseIf cnt(rLast(r)) - cnt(r) > 0 Then
            n = n + 1
            filas(n) = r
        End If
    Next r
    If n > PLMAX - PL0 - 8 Then Err.Raise 1005, , "La planilla del día " & d & " tiene demasiadas filas (" & n & ")."
    ReDim o(1 To n, 1 To 11)
    ReDim stt(1 To n)
    tot = 0
    nP = 0
    nNo = 0
    nOk = 0
    For k = 1 To n
        r = filas(k)
        o(k, 1) = k
        o(k, 2) = rItem(r)
        o(k, 3) = String(2 * (rLvl(r) - 1), " ") & LTrim(rDesc(r))
        If rTipo(r) = "P" Then
            ac = rAnt(r)
            For dd = 1 To d
                ac = ac + rVal(r, dd)
            Next dd
            o(k, 4) = rUnd(r)
            o(k, 5) = rVal(r, d)
            o(k, 6) = ac
            o(k, 7) = rMet(r)
            If rMet(r) > 0 Then o(k, 8) = ac / rMet(r)
            o(k, 9) = rPU(r)
            o(k, 10) = Redondear(rVal(r, d) * rPU(r), 2)
            tot = tot + Redondear(rVal(r, d) * rPU(r), 2)
            nP = nP + 1
            If fuera Then
                stt(k) = 1
                est = "FUERA DEL PLAZO: NO VALORIZAR"
            ElseIf miss(r, d) > 0 And miss(r, d) >= umbral Then
                stt(k) = 1
                est = "NO VALORIZAR – falta: " & ListaDia(r, t, 1, 230)
            ElseIf ac > rMet(r) * 1.0001 + 0.001 Then
                stt(k) = 2
                est = "SUPERA EL METRADO DEL EXP. TÉC."
            ElseIf rMaxD(r) > 0 And rVal(r, d) > rMaxD(r) * 1.0001 Then
                stt(k) = 3
                est = "SUPERA RENDIMIENTO (MÁX. " & Format(rMaxD(r), "#,##0.00") & " " & rUnd(r) & "/día)"
            ElseIf miss(r, d) > 0 Then
                stt(k) = 4
                est = "REVISAR – falta: " & ListaDia(r, t, 1, 230)
            ElseIf pcf(r, d) > 0 Then
                stt(k) = 4
                est = "POR CONFIRMAR: " & ListaDia(r, t, 2, 230)
            Else
                stt(k) = 5
                est = "CONFORME"
            End If
            o(k, 11) = est
            If stt(k) = 1 Then nNo = nNo + 1
            If stt(k) = 5 Then nOk = nOk + 1
        End If
    Next k
    ws.Range(ws.Cells(PL0, 1), ws.Cells(PL0 + n - 1, 11)).Value = o
    ' firmas (solo el cargo; se firma sobre la línea)
    sig = PL0 + n + 3
    ultima = sig + 1
    If sFirma1 <> "" Then
        ws.Cells(sig, 3).Value = "______________________________"
        ws.Cells(sig + 1, 3).Value = sFirma1
    End If
    If sFirma2 <> "" Then
        ws.Cells(sig, 11).Value = "______________________________"
        ws.Cells(sig + 1, 11).Value = sFirma2
    End If
    If sFirma3 <> "" Then
        ws.Cells(sig, 7).Value = "______________________________"
        ws.Cells(sig + 1, 7).Value = sFirma3
    End If
    If nNo > 0 Then
        ws.Range("K5").Value = nNo & " PARTIDA(S) PARA NO VALORIZAR"
    ElseIf nOk = nP Then
        ws.Range("K5").Value = "CONFORME"
    Else
        ws.Range("K5").Value = "REVISAR ALERTAS"
    End If
    ws.Range("D6").Value = tot
    ws.Range("I6").Value = nP
    nPl = nPl + 1
    nRojo = nRojo + nNo
    ' formato de las filas (si alguna propiedad no existe en otra versión, se omite)
    On Error Resume Next
    For k = 1 To n
        r = filas(k)
        fila = PL0 + k - 1
        With ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 11))
            .Borders.LineStyle = XL_CONT
            .Borders.Weight = XL_THIN
            .Borders.Color = RGB(158, 158, 158)
            If rTipo(r) = "T" Then
                .Font.Bold = True
                .Font.Color = ColorNivel(rLvl(r))
            End If
        End With
        If rTipo(r) = "P" Then
            ws.Cells(fila, 5).Font.Bold = True
            Pintar ws.Cells(fila, 5), stt(k)
            Pintar ws.Cells(fila, 11), stt(k)
            If Len(CStr(o(k, 11))) > 70 Then
                ws.Cells(fila, 11).WrapText = True
                ws.Rows(fila).RowHeight = 11.5 * (Int((Len(CStr(o(k, 11))) - 1) / 72) + 1)
            End If
        End If
    Next k
    With ws.Range(ws.Cells(sig + 1, 3), ws.Cells(sig + 1, 11))
        .Font.Bold = True
    End With
    ws.Cells(sig, 7).HorizontalAlignment = XL_CENTER
    ws.Cells(sig + 1, 7).HorizontalAlignment = XL_CENTER
    ws.PageSetup.PrintArea = "$A$1:$K$" & ultima
    If ws.Visible <> XL_VISIBLE Then ws.Visible = XL_VISIBLE
    On Error GoTo 0
End Sub

Private Sub LimpiarPlanilla(ByVal ws As Worksheet)
    Dim rg As Range
    Set rg = ws.Range(ws.Cells(PL0, 1), ws.Cells(PLMAX, 11))
    rg.ClearContents
    ws.Range(ws.Cells(PL0, 2), ws.Cells(PLMAX, 2)).NumberFormat = "@"
    On Error Resume Next
    rg.Interior.ColorIndex = XL_NONE
    rg.Font.Bold = False
    rg.Font.Color = 0
    rg.Borders.LineStyle = XL_NONE
    rg.WrapText = False
    rg.RowHeight = 12.75
    rg.HorizontalAlignment = 1
    ws.Range(ws.Cells(PL0, 1), ws.Cells(PLMAX, 1)).HorizontalAlignment = XL_CENTER
    ws.Range(ws.Cells(PL0, 4), ws.Cells(PLMAX, 4)).HorizontalAlignment = XL_CENTER
    On Error GoTo 0
End Sub

Private Sub Pintar(ByVal c As Range, ByVal st As Long)
    Select Case st
        Case 1
            c.Interior.Color = RGB(255, 199, 206)
            c.Font.Color = RGB(156, 0, 6)
            c.Font.Bold = True
        Case 2
            c.Interior.Color = RGB(228, 223, 236)
            c.Font.Color = RGB(91, 44, 111)
            c.Font.Bold = True
        Case 3
            c.Interior.Color = RGB(248, 203, 173)
            c.Font.Color = RGB(131, 60, 11)
            c.Font.Bold = True
        Case 4
            c.Interior.Color = RGB(255, 235, 156)
            c.Font.Color = RGB(156, 87, 0)
        Case 5
            c.Interior.Color = RGB(198, 239, 206)
            c.Font.Color = RGB(0, 97, 0)
    End Select
End Sub


'========================================================================== LISTAS DE INSUMOS
Private Function ListaMes(ByVal r As Long, ByVal t1 As Double, ByVal t2 As Double, ByVal tipo As Long, ByVal lim As Long) As String
    ' tipo 1: insumos sin O/S ni ingreso en algún día con metrado; tipo 2: con ingreso por confirmar al primer día
    Dim j As Long, s As String, ok As Boolean
    j = pIni(r)
    Do While j > 0
        If aSh(j) > 0 Then
            ok = False
            If tipo = 1 Then
                If t1 = 0 Then
                    ok = True
                ElseIf aL(j) > t1 Or aN(j) < t2 Then
                    ok = True
                End If
            ElseIf t1 > 0 Then
                If aL(j) <= t1 And aM(j) > t1 Then ok = True
            End If
            If ok Then s = Agregar(s, aNom(j))
        End If
        j = aSig(j)
    Loop
    ListaMes = Cortar(s, lim)
End Function

Private Function ListaDia(ByVal r As Long, ByVal t As Double, ByVal tipo As Long, ByVal lim As Long) As String
    ' tipo 1: sin O/S ni ingreso a la fecha; tipo 2: ingreso por confirmar a la fecha
    Dim j As Long, s As String, ok As Boolean
    j = pIni(r)
    Do While j > 0
        If aSh(j) > 0 Then
            ok = False
            If tipo = 1 Then
                If t = 0 Then
                    ok = True
                ElseIf aL(j) > t Or aN(j) < t Then
                    ok = True
                End If
            ElseIf t > 0 Then
                If aL(j) <= t And aM(j) > t And aN(j) >= t Then ok = True
            End If
            If ok Then s = Agregar(s, aNom(j))
        End If
        j = aSig(j)
    Loop
    ListaDia = Cortar(s, lim)
End Function
