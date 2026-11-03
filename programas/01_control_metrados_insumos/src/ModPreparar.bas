Attribute VB_Name = "ModPreparar"
Option Explicit

'==========================================================================================
'  CONTROL DE METRADOS DIARIOS - PREPARAR OBRA
'  Arma el RESUMEN, INSUMOS, MAQ_COMB y MANO_OBRA a partir de las hojas PRESUPUESTO y ACU.
'  Conserva lo ya digitado (metrados de los días, acumulado anterior, cuadrillas, CONTROLAR,
'  observaciones, partes diarios, vales y tareo) buscándolo por ítem, insumo o categoría.
'==========================================================================================

' ---- presupuesto
Private pN As Long
Private pCod() As String, pKey() As String, pDesc() As String, pUnd() As String, pTipo() As String
Private pMet() As Double, pPU() As Double, pPar() As Double, pRend() As Double
Private pLvl() As Long, pLast() As Long, pFila() As Long
Private pTieneAcu() As Boolean, pSumAcu() As Double
Private colP As Collection
Private nSinItem As Long, sDup As String, nDup As Long
' ---- insumos y categorías del A.C.U.
Private qN As Long
Private qNom() As String, qUnd() As String, qTipo() As String, qPre() As Double, qOrd() As String
Private colQ As Collection
Private cN As Long
Private cNomP() As String
Private nAcuFilas As Long, nAcuSinPres As Long, sAcuSinPres As String, nTipoAuto As Long
' ---- datos anteriores (para conservarlos)
Private nOld As Long
Private oH() As Variant, oD() As Variant, oAU() As Variant, oUsada() As Boolean, oItem() As String, oTiene() As Boolean
Private colO As Collection
Private colOI As Collection
Private oCtrl() As Variant, oObs() As Variant
Private colOM As Collection
Private oMq() As Variant
Private colOT As Collection
Private oTa() As Variant
' ---- colores
Private cGris As Long, cCalc As Long, cAma As Long, cSubh As Long, cAzul As Long, cBorde As Long


'========================================================================== BOTONES
Public Sub PrepararObra()
    Dim msg As String
    msg = Preparar()
    If Len(msg) > 0 Then MsgBox msg, IIf(Left(msg, 8) = "No se pu", vbExclamation, vbInformation), "Preparar obra"
End Sub

Public Sub PrepararSilencioso()
    Dim msg As String
    gSilencio = True
    msg = Preparar()
    gSilencio = False
    ThisWorkbook.Worksheets("DATOS").Range(D_MSGPREP).Value = msg
End Sub


'========================================================================== PROCESO
Private Function Preparar() As String
    Dim calcMode As Long, t0 As Double, msg As String, aviso As String, perdidos As String, nPerd As Long
    Dim msgAct As String, i As Long, nT As Long, nP As Long, cD As Double, nSinAcu As Long, nPU As Long
    Dim nMat As Long, nEq As Long, nSc As Long, nComb As Long, nCons As Long
    t0 = Timer
    cGris = RGB(242, 242, 242)
    cCalc = RGB(244, 247, 251)
    cAma = RGB(255, 242, 204)
    cSubh = RGB(217, 225, 242)
    cAzul = RGB(31, 56, 100)
    cBorde = RGB(158, 158, 158)
    Inicio calcMode, "Preparando la obra..."
    On Error GoTo Fallo
    LeerPresupuesto
    If pN = 0 Then Err.Raise 2001, , "La hoja PRESUPUESTO está vacía. Use IMPORTAR PRESUPUESTO o pegue el presupuesto (columnas A a F)."
    If nDup > 0 Then Err.Raise 2002, , "El PRESUPUESTO tiene ítems repetidos (" & nDup & "): " & Cortar(sDup, 300) & _
                                      ". Corrija los ítems y vuelva a presionar PREPARAR OBRA."
    If pN > RRMAX - RR0 + 1 Then Err.Raise 2003, , "El presupuesto tiene " & pN & " filas; el programa admite hasta " & (RRMAX - RR0 + 1) & "."
    LeerAcuPrep
    If qN > INMAX - IN0 + 1 Then Err.Raise 2004, , "El A.C.U. tiene " & qN & " insumos; el programa admite hasta " & (INMAX - IN0 + 1) & "."
    GuardarAnteriores nPerd, perdidos
    If nPerd > 0 And Not Silencio() Then
        Application.ScreenUpdating = True
        If MsgBox(nPerd & " partida(s) con metrados digitados ya no figuran en el PRESUPUESTO y sus metrados se perderán:" & _
                  vbCrLf & vbCrLf & Cortar(perdidos, 400) & vbCrLf & vbCrLf & "¿Desea continuar?", vbYesNo + vbExclamation, _
                  "Preparar obra") <> vbYes Then
            Preparar = ""
            GoTo Salir
        End If
        Application.ScreenUpdating = False
    End If
    RevisarPresupuesto nSinAcu, nPU
    ArmarResumen
    ArmarInsumos nMat, nEq, nSc
    ArmarMaquinaria nComb
    ArmarManoObra
    With ThisWorkbook.Worksheets("DATOS")
        .Range(D_ULTPREP).NumberFormat = "@"
        .Range(D_ULTPREP).Value = Format(Now, "dd/mm/yyyy hh:mm")
        .Range(D_HUELLAP).Value = .Range(D_HUELLAPHOY).Value
    End With
    nT = 0
    nP = 0
    cD = 0
    For i = 1 To pN
        If pTipo(i) = "P" Then
            nP = nP + 1
            cD = cD + pPar(i)
        Else
            nT = nT + 1
        End If
    Next i
    nCons = 0
    For i = 1 To nOld
        If oUsada(i) And oTiene(i) Then nCons = nCons + 1
    Next i
    msg = "Obra preparada en " & Format(Timer - t0, "0.0") & " s." & vbCrLf & vbCrLf & _
          "• Presupuesto: " & nT & " títulos y " & nP & " partidas (C.D. S/ " & Format(cD, "#,##0.00") & ")." & vbCrLf & _
          "• A.C.U.: " & nAcuFilas & " filas; " & qN & " insumos (" & nMat & " materiales, " & nEq & " equipos, " & nSc & _
          " servicios) y " & cN & " categoría(s) de mano de obra." & vbCrLf & _
          "• Maquinaria y combustible: " & nEq & " equipo(s) y " & nComb & " combustible(s)."
    aviso = ""
    If nSinAcu > 0 Then aviso = aviso & vbCrLf & "• " & nSinAcu & " partida(s) sin A.C.U. (no controlan insumos)."
    If nPU > 0 Then aviso = aviso & vbCrLf & "• " & nPU & " partida(s) cuyo A.C.U. no suma el P.U. del presupuesto."
    If nAcuSinPres > 0 Then aviso = aviso & vbCrLf & "• " & nAcuSinPres & " partida(s) del A.C.U. no están en el presupuesto: " & _
                                    Cortar(sAcuSinPres, 150)
    If nSinItem > 0 Then aviso = aviso & vbCrLf & "• " & nSinItem & " fila(s) del PRESUPUESTO sin ÍTEM (no se usan)."
    If nTipoAuto > 0 Then aviso = aviso & vbCrLf & "• " & nTipoAuto & " fila(s) del A.C.U. sin TIPO: se asignó por la unidad."
    If aviso <> "" Then msg = msg & vbCrLf & vbCrLf & "Revise (columna REVISIÓN de la hoja PRESUPUESTO):" & aviso
    If nCons > 0 Then msg = msg & vbCrLf & vbCrLf & "Se conservaron los metrados digitados de " & nCons & " partida(s)."
    msgAct = Ejecutar()
    Preparar = msg & vbCrLf & vbCrLf & "ACTUALIZAR TODO: " & msgAct
Salir:
    Fin calcMode
    Exit Function
Fallo:
    Preparar = "No se pudo preparar la obra." & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    Resume Salir
End Function


'========================================================================== LECTURA
Private Function CodTexto(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Or IsError(v) Then
        CodTexto = ""
        Exit Function
    End If
    If VarType(v) = vbString Then
        s = Limpio(v)
    Else
        s = Replace(Trim(CStr(v)), ",", ".")
    End If
    If Right(s, 1) = "." Then s = Left(s, Len(s) - 1)
    CodTexto = s
End Function

Private Sub LeerPresupuesto()
    Dim ws As Worksheet, v As Variant, i As Long, j As Long, n As Long, ult As Long, cod As String, k As String, ds As String
    Set ws = ThisWorkbook.Worksheets("PRESUPUESTO")
    ult = PP0 - 1
    v = ws.Range(ws.Cells(PP0, 1), ws.Cells(PPMAX, 2)).Value
    For i = UBound(v, 1) To 1 Step -1
        If Txt(v(i, 1)) <> "" Or Txt(v(i, 2)) <> "" Then
            ult = PP0 + i - 1
            Exit For
        End If
    Next i
    pN = 0
    nSinItem = 0
    nDup = 0
    sDup = ""
    Set colP = New Collection
    If ult < PP0 Then Exit Sub
    v = ws.Range(ws.Cells(PP0, 1), ws.Cells(ult, 7)).Value
    n = UBound(v, 1)
    ReDim pCod(1 To n): ReDim pKey(1 To n): ReDim pDesc(1 To n): ReDim pUnd(1 To n): ReDim pTipo(1 To n)
    ReDim pMet(1 To n): ReDim pPU(1 To n): ReDim pPar(1 To n): ReDim pRend(1 To n)
    ReDim pLvl(1 To n): ReDim pLast(1 To n): ReDim pFila(1 To n)
    ReDim pTieneAcu(1 To n): ReDim pSumAcu(1 To n)
    For i = 1 To n
        cod = CodTexto(v(i, 1))
        ds = Limpio(v(i, 2))
        If cod = "" Then
            If ds <> "" Then nSinItem = nSinItem + 1
        Else
            k = Clave(cod)
            If Buscar(colP, "P" & k) > 0 Then
                nDup = nDup + 1
                sDup = Agregar(sDup, cod)
            Else
                pN = pN + 1
                pCod(pN) = cod
                pKey(pN) = k
                pDesc(pN) = ds
                pUnd(pN) = Limpio(v(i, 3))
                pFila(pN) = PP0 + i - 1
                If pUnd(pN) <> "" Then
                    pTipo(pN) = "P"
                    pMet(pN) = Num(v(i, 4))
                    pPU(pN) = Num(v(i, 5))
                    pPar(pN) = Num(v(i, 6))
                    If pPar(pN) = 0 Then pPar(pN) = Redondear(pMet(pN) * pPU(pN), 2)
                    pRend(pN) = Num(v(i, 7))
                Else
                    pTipo(pN) = "T"
                End If
                pLvl(pN) = Nivel(cod)
                colP.Add pN, "P" & k
            End If
        End If
    Next i
    For i = 1 To pN
        pLast(i) = i
        If pTipo(i) = "T" Then
            k = pKey(i) & "."
            j = i + 1
            Do While j <= pN
                If Left(pKey(j), Len(k)) <> k Then Exit Do
                pLast(i) = j
                j = j + 1
            Loop
        End If
    Next i
End Sub

Public Function TipoPorUnidad(ByVal und As String) As String
    Dim u As String
    u = Replace(Replace(Replace(Normal(und), " ", ""), ".", ""), "-", "")
    Select Case u
        Case "HH", "HHOMBRE", "HORAHOMBRE"
            TipoPorUnidad = "MO"
        Case "HM", "HE", "HMAQ", "HORAMAQUINA", "HEQ"
            TipoPorUnidad = "EQ"
        Case "%MO", "%"
            TipoPorUnidad = "%MO"
        Case Else
            TipoPorUnidad = "MAT"
    End Select
End Function

Private Sub LeerAcuPrep()
    Dim ws As Worksheet, v As Variant, j As Long, n As Long, r As Long, k As Long, cod As String, ins As String, und As String
    Dim tP As String, cant As Double, pre As Double, nm As String, w() As Variant, cambio As Boolean
    Dim prio As Variant, c As Long, cc As Long, otros() As String, nOtros As Long, ordC() As String
    Set ws = ThisWorkbook.Worksheets("ACU")
    qN = 0
    cN = 0
    nAcuFilas = 0
    nAcuSinPres = 0
    sAcuSinPres = ""
    nTipoAuto = 0
    Set colQ = New Collection
    ReDim qNom(1 To 1): ReDim qUnd(1 To 1): ReDim qTipo(1 To 1): ReDim qPre(1 To 1): ReDim qOrd(1 To 1)
    ReDim cNomP(1 To 1)
    v = ws.Range(ws.Cells(AC0, 1), ws.Cells(ACMAX, 1)).Value
    n = 0
    For j = 1 To UBound(v, 1)
        If Txt(v(j, 1)) = "" Then Exit For
        n = j
    Next j
    nAcuFilas = n
    If n = 0 Then Exit Sub
    v = ws.Range(ws.Cells(AC0, 1), ws.Cells(AC0 + n - 1, 6)).Value
    ReDim w(1 To n, 1 To 1)
    cambio = False
    ReDim otros(1 To n)
    nOtros = 0
    For j = 1 To n
        cod = CodTexto(v(j, 1))
        ins = Txt(v(j, 2))
        und = Txt(v(j, 3))
        tP = UCase(Txt(v(j, 4)))
        If tP = "" And ins <> "" Then
            tP = TipoPorUnidad(und)
            nTipoAuto = nTipoAuto + 1
            cambio = True
        End If
        w(j, 1) = tP
        cant = Num(v(j, 5))
        pre = Num(v(j, 6))
        r = Buscar(colP, "P" & Clave(cod))
        If r = 0 Then
            If InStr(1, SEP & sAcuSinPres & SEP, SEP & cod & SEP, vbBinaryCompare) = 0 Then
                nAcuSinPres = nAcuSinPres + 1
                sAcuSinPres = Agregar(sAcuSinPres, cod)
            End If
        Else
            pTieneAcu(r) = True
            If tP = "%MO" Then
                pSumAcu(r) = pSumAcu(r) + cant * pre / 100
            Else
                pSumAcu(r) = pSumAcu(r) + cant * pre
            End If
        End If
        If ins <> "" Then
            If tP = "MO" Then
                nm = UCase(ins)
                If Buscar(colQ, "C" & nm) = 0 Then
                    colQ.Add 1, "C" & nm
                    nOtros = nOtros + 1
                    otros(nOtros) = nm
                End If
            ElseIf tP = "MAT" Or tP = "EQ" Or tP = "SC" Then
                nm = UCase(ins)
                k = Buscar(colQ, "I" & nm)
                If k = 0 Then
                    qN = qN + 1
                    ReDim Preserve qNom(1 To qN): ReDim Preserve qUnd(1 To qN): ReDim Preserve qTipo(1 To qN)
                    ReDim Preserve qPre(1 To qN): ReDim Preserve qOrd(1 To qN)
                    qNom(qN) = ins
                    qUnd(qN) = und
                    qTipo(qN) = tP
                    qPre(qN) = pre
                    Select Case tP
                        Case "MAT"
                            qOrd(qN) = "1|" & nm
                        Case "SC"
                            qOrd(qN) = "2|" & nm
                        Case Else
                            qOrd(qN) = "3|" & nm
                    End Select
                    colQ.Add qN, "I" & nm
                End If
            End If
        End If
    Next j
    If cambio Then ws.Range(ws.Cells(AC0, 4), ws.Cells(AC0 + n - 1, 4)).Value = w
    ' categorías de mano de obra: primero las habituales, luego en el orden en que aparecen
    prio = Array("CAPATAZ", "OPERARIO", "OFICIAL", "PEON")
    cN = 0
    ReDim ordC(1 To nOtros + 1)
    For c = 0 To 3
        For cc = 1 To nOtros
            If otros(cc) <> "" Then
                If Normal(otros(cc)) = prio(c) Then
                    cN = cN + 1
                    ordC(cN) = otros(cc)
                    otros(cc) = ""
                End If
            End If
        Next cc
    Next c
    For cc = 1 To nOtros
        If otros(cc) <> "" Then
            cN = cN + 1
            ordC(cN) = otros(cc)
        End If
    Next cc
    If cN > 0 Then
        ReDim cNomP(1 To cN)
        For c = 1 To cN
            cNomP(c) = ordC(c)
        Next c
    End If
    OrdenarInsumos
End Sub

Private Sub OrdenarInsumos()
    ' ordena por tipo (material, servicio, equipo) y nombre (inserción binaria estable)
    Dim i As Long, j As Long, tN As String, tu As String, tT As String, tOrd As String, tP As Double
    For i = 2 To qN
        tN = qNom(i): tu = qUnd(i): tT = qTipo(i): tOrd = qOrd(i): tP = qPre(i)
        j = i - 1
        Do While j >= 1
            If StrComp(qOrd(j), tOrd, vbBinaryCompare) <= 0 Then Exit Do
            qNom(j + 1) = qNom(j): qUnd(j + 1) = qUnd(j): qTipo(j + 1) = qTipo(j): qOrd(j + 1) = qOrd(j): qPre(j + 1) = qPre(j)
            j = j - 1
        Loop
        qNom(j + 1) = tN: qUnd(j + 1) = tu: qTipo(j + 1) = tT: qOrd(j + 1) = tOrd: qPre(j + 1) = tP
    Next i
End Sub


'========================================================================== DATOS ANTERIORES
Private Sub GuardarAnteriores(ByRef nPerd As Long, ByRef perdidos As String)
    Dim ws As Worksheet, v As Variant, i As Long, d As Long, n As Long, k As String, tiene As Boolean, nm As String
    ' ---- RESUMEN
    Set ws = ThisWorkbook.Worksheets("RESUMEN")
    Set colO = New Collection
    nPerd = 0
    perdidos = ""
    v = ws.Range(ws.Cells(RR0, 2), ws.Cells(RRMAX, 2)).Value
    n = 0
    For i = 1 To UBound(v, 1)
        If Txt(v(i, 1)) = "" Then Exit For
        n = i
    Next i
    nOld = n
    ReDim oH(1 To n + 1): ReDim oD(1 To n + 1, 1 To ND): ReDim oAU(1 To n + 1): ReDim oUsada(1 To n + 1)
    ReDim oItem(1 To n + 1): ReDim oTiene(1 To n + 1)
    If n > 0 Then
        v = ws.Range(ws.Cells(RR0, 1), ws.Cells(RR0 + n - 1, 47)).Value
        For i = 1 To n
            oItem(i) = Txt(v(i, 2))
            k = Clave(oItem(i))
            If UCase(Txt(v(i, 1))) = "P" And Buscar(colO, "O" & k) = 0 Then
                oH(i) = v(i, 8)
                tiene = (Num(v(i, 8)) <> 0)
                For d = 1 To ND
                    oD(i, d) = v(i, CD1 + d - 1)
                    If Num(oD(i, d)) <> 0 Then tiene = True
                Next d
                oAU(i) = v(i, 47)
                oTiene(i) = tiene
                colO.Add i, "O" & k
                If tiene And Buscar(colP, "P" & k) = 0 Then
                    nPerd = nPerd + 1
                    perdidos = Agregar(perdidos, oItem(i))
                ElseIf tiene Then
                    If pTipo(Buscar(colP, "P" & k)) <> "P" Then
                        nPerd = nPerd + 1
                        perdidos = Agregar(perdidos, oItem(i))
                    End If
                End If
            End If
        Next i
    End If
    ' ---- INSUMOS (CONTROLAR y OBSERVACIÓN)
    Set ws = ThisWorkbook.Worksheets("INSUMOS")
    Set colOI = New Collection
    v = ws.Range(ws.Cells(IN0, 1), ws.Cells(INMAX, 48)).Value
    ReDim oCtrl(1 To UBound(v, 1)): ReDim oObs(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        nm = UCase(Txt(v(i, 2)))
        If nm = "" Then Exit For
        If Buscar(colOI, "I" & nm) = 0 Then
            colOI.Add i, "I" & nm
            oCtrl(i) = v(i, 6)
            oObs(i) = v(i, 48)
        End If
    Next i
    ' ---- MAQ_COMB (partes diarios y vales)
    Set ws = ThisWorkbook.Worksheets("MAQ_COMB")
    Set colOM = New Collection
    v = ws.Range(ws.Cells(1, 1), ws.Cells(MQMAX, 38)).Value
    ReDim oMq(1 To MQMAX, 1 To ND)
    For i = MQ0 + 1 To MQMAX
        If Txt(v(i, 4)) = LB_PARTE Or Txt(v(i, 4)) = LB_VALES Then
            nm = UCase(Txt(v(i - 1, 2)))
            If nm <> "" And Buscar(colOM, "M" & nm) = 0 Then
                colOM.Add i, "M" & nm
                For d = 1 To ND
                    oMq(i, d) = v(i, 7 + d)
                Next d
            End If
        End If
    Next i
    ' ---- MANO_OBRA (tareo)
    Set ws = ThisWorkbook.Worksheets("MANO_OBRA")
    Set colOT = New Collection
    v = ws.Range(ws.Cells(1, 1), ws.Cells(MOMAX, 38)).Value
    ReDim oTa(1 To MOMAX, 1 To ND)
    For i = MO0 + 1 To MOMAX
        If Txt(v(i, 4)) = LB_TRAB Then
            nm = UCase(Txt(v(i - 1, 2)))
            If nm <> "" And Buscar(colOT, "T" & nm) = 0 Then
                colOT.Add i, "T" & nm
                For d = 1 To ND
                    oTa(i, d) = v(i, 7 + d)
                Next d
            End If
        End If
    Next i
End Sub


'========================================================================== PRESUPUESTO: REVISIÓN
Private Sub RevisarPresupuesto(ByRef nSinAcu As Long, ByRef nPU As Long)
    Dim ws As Worksheet, i As Long, ult As Long, o() As Variant, f As Long, tol As Double, v As Variant
    Set ws = ThisWorkbook.Worksheets("PRESUPUESTO")
    nSinAcu = 0
    nPU = 0
    ult = UltimaFila(ws, 2, PP0, PPMAX)
    If UltimaFila(ws, 1, PP0, PPMAX) > ult Then ult = UltimaFila(ws, 1, PP0, PPMAX)
    ws.Range(ws.Cells(PP0, 8), ws.Cells(PPMAX, 9)).ClearContents
    If ult < PP0 Then Exit Sub
    ReDim o(1 To ult - PP0 + 1, 1 To 2)
    v = ws.Range(ws.Cells(PP0, 1), ws.Cells(ult, 2)).Value
    For f = 1 To ult - PP0 + 1
        If Txt(v(f, 1)) = "" And Txt(v(f, 2)) <> "" Then o(f, 2) = "SIN ÍTEM: NO SE USA"
    Next f
    For i = 1 To pN
        f = pFila(i) - PP0 + 1
        o(f, 1) = pTipo(i)
        If pTipo(i) = "P" Then
            If Not pTieneAcu(i) Then
                o(f, 2) = "SIN A.C.U."
                nSinAcu = nSinAcu + 1
            Else
                tol = pPU(i) * 0.01
                If tol < 0.1 Then tol = 0.1
                If Abs(pSumAcu(i) - pPU(i)) > tol Then
                    o(f, 2) = "EL A.C.U. SUMA S/ " & Format(pSumAcu(i), "#,##0.00") & " (P.U. S/ " & Format(pPU(i), "#,##0.00") & ")"
                    nPU = nPU + 1
                End If
            End If
            If pMet(i) <= 0 Then o(f, 2) = Unir(Txt(o(f, 2)), "METRADO CERO")
        End If
    Next i
    ws.Range(ws.Cells(PP0, 8), ws.Cells(ult, 9)).Value = o
End Sub


'========================================================================== RESUMEN
Private Sub ArmarResumen()
    Dim ws As Worksheet, n As Long, hasta As Long, i As Long, d As Long, r As Long, k As Long
    Dim o() As Variant, dv() As Variant, tu() As Variant, f0 As Long, f1 As Long
    Set ws = ThisWorkbook.Worksheets("RESUMEN")
    n = pN
    hasta = RR0 + n - 1
    If nOld > n Then hasta = RR0 + nOld - 1
    hasta = hasta + 1
    With ws.Range(ws.Cells(RR0, 1), ws.Cells(hasta, 52))
        .ClearContents
        .Interior.ColorIndex = XL_NONE
        .Font.Bold = False
        .Font.Color = 0
    End With
    ws.Range(ws.Cells(RR0, 1), ws.Cells(hasta, 1)).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(RR0, 49), ws.Cells(hasta, 49)).Font.Color = 0
    ws.Range(ws.Cells(RR0, 50), ws.Cells(hasta, 50)).Font.Color = RGB(156, 0, 6)
    ws.Range(ws.Cells(RR0, 2), ws.Cells(hasta, 4)).NumberFormat = "@"
    ReDim o(1 To n, 1 To 8)
    ReDim dv(1 To n, 1 To ND)
    ReDim tu(1 To n, 1 To 2)
    For i = 1 To n
        o(i, 1) = pTipo(i)
        o(i, 2) = pCod(i)
        o(i, 3) = String(2 * (pLvl(i) - 1), " ") & pDesc(i)
        If pTipo(i) = "P" Then
            o(i, 4) = pUnd(i)
            o(i, 5) = pMet(i)
            o(i, 6) = pPU(i)
            o(i, 7) = pPar(i)
            o(i, 8) = 0
            If pRend(i) > 0 Then tu(i, 1) = pRend(i)
            tu(i, 2) = 1
            k = Buscar(colO, "O" & pKey(i))
            If k > 0 Then
                oUsada(k) = True
                If Not IsEmpty(oH(k)) Then o(i, 8) = oH(k)
                For d = 1 To ND
                    dv(i, d) = oD(k, d)
                Next d
                If Num(oAU(k)) > 0 Then tu(i, 2) = oAU(k)
            End If
        End If
    Next i
    ws.Range(ws.Cells(RR0, 1), ws.Cells(RR0 + n - 1, 8)).Value = o
    ws.Range(ws.Cells(RR0, CD1), ws.Cells(RR0 + n - 1, CD1 + ND - 1)).Value = dv
    ws.Range(ws.Cells(RR0, 46), ws.Cells(RR0 + n - 1, 47)).Value = tu
    ' fórmulas de las partidas (columnas AN a AS y AV)
    ws.Range(ws.Cells(RR0, 40), ws.Cells(RR0 + n - 1, 40)).FormulaR1C1 = "=SUM(RC9:RC39)"
    ws.Range(ws.Cells(RR0, 41), ws.Cells(RR0 + n - 1, 41)).FormulaR1C1 = "=RC8+RC40"
    ws.Range(ws.Cells(RR0, 42), ws.Cells(RR0 + n - 1, 42)).FormulaR1C1 = "=IF(RC5>0,RC41/RC5,"""")"
    ws.Range(ws.Cells(RR0, 43), ws.Cells(RR0 + n - 1, 43)).FormulaR1C1 = "=RC5-RC41"
    ws.Range(ws.Cells(RR0, 44), ws.Cells(RR0 + n - 1, 44)).FormulaR1C1 = "=ROUND(RC40*RC6,2)"
    ws.Range(ws.Cells(RR0, 45), ws.Cells(RR0 + n - 1, 45)).FormulaR1C1 = "=ROUND(RC41*RC6,2)"
    ws.Range(ws.Cells(RR0, 48), ws.Cells(RR0 + n - 1, 48)).FormulaR1C1 = "=IF(RC46="""","""",RC46*RC47)"
    ' fondo de las columnas calculadas
    ws.Range(ws.Cells(RR0, 5), ws.Cells(RR0 + n - 1, 7)).Interior.Color = cCalc
    ws.Range(ws.Cells(RR0, 40), ws.Cells(RR0 + n - 1, 52)).Interior.Color = cCalc
    ws.Range(ws.Cells(RR0, 46), ws.Cells(RR0 + n - 1, 46)).Interior.Color = cCalc
    ws.Range(ws.Cells(RR0, 47), ws.Cells(RR0 + n - 1, 47)).Interior.ColorIndex = XL_NONE
    ws.Range(ws.Cells(RR0, 40), ws.Cells(RR0 + n - 1, 40)).Font.Bold = True
    ws.Range(ws.Cells(RR0, 44), ws.Cells(RR0 + n - 1, 44)).Font.Bold = True
    ' títulos: sumas de sus partidas, negrita y color por nivel
    For i = 1 To n
        If pTipo(i) = "T" Then
            r = RR0 + i - 1
            ws.Range(ws.Cells(r, 40), ws.Cells(r, 48)).ClearContents
            f0 = r + 1
            f1 = RR0 + pLast(i) - 1
            If f1 >= f0 Then
                ws.Cells(r, 7).Formula = "=SUMIFS(G" & f0 & ":G" & f1 & ",$A" & f0 & ":$A" & f1 & ",""P"")"
                ws.Cells(r, 44).Formula = "=SUMIFS(AR" & f0 & ":AR" & f1 & ",$A" & f0 & ":$A" & f1 & ",""P"")"
                ws.Cells(r, 45).Formula = "=SUMIFS(AS" & f0 & ":AS" & f1 & ",$A" & f0 & ":$A" & f1 & ",""P"")"
                ws.Cells(r, 42).Formula = "=IF(G" & r & ">0,AS" & r & "/G" & r & ","""")"
            End If
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 52))
                .Interior.Color = cGris
                .Font.Bold = True
                .Font.Color = ColorNivel(pLvl(i))
            End With
            ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)
        End If
    Next i
    ' hoja auxiliar
    ThisWorkbook.Worksheets("AUX").Range("A" & RR0 & ":CP" & RRMAX).ClearContents
End Sub


'========================================================================== INSUMOS
Private Function NoControlar(ByVal nom As String) As Boolean
    Dim p() As String, i As Long, t As String, s As String
    NoControlar = False
    s = Normal(ThisWorkbook.Worksheets("DATOS").Range(D_NOCTRL).Value)
    If s = "" Then Exit Function
    p = Split(s, ";")
    For i = LBound(p) To UBound(p)
        t = Trim(p(i))
        If t <> "" Then
            If InStr(1, Normal(nom), t, vbBinaryCompare) > 0 Then
                NoControlar = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function EsCombustible(ByVal nom As String) As Boolean
    Dim p() As String, i As Long, t As String, s As String
    EsCombustible = False
    s = Normal(ThisWorkbook.Worksheets("DATOS").Range(D_COMB).Value)
    If s = "" Then Exit Function
    p = Split(s, ";")
    For i = LBound(p) To UBound(p)
        t = Trim(p(i))
        If t <> "" Then
            If InStr(1, Normal(nom), t, vbBinaryCompare) > 0 Then
                EsCombustible = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub ArmarInsumos(ByRef nMat As Long, ByRef nEq As Long, ByRef nSc As Long)
    Dim ws As Worksheet, o() As Variant, ob() As Variant, i As Long, k As Long, ctrl As String
    Set ws = ThisWorkbook.Worksheets("INSUMOS")
    nMat = 0
    nEq = 0
    nSc = 0
    With ws.Range(ws.Cells(IN0, 1), ws.Cells(INMAX, 49))
        .ClearContents
    End With
    ws.Range(ws.Cells(IN0, 6), ws.Cells(INMAX, 6)).Interior.ColorIndex = XL_NONE
    ws.Range(ws.Cells(IN0, 48), ws.Cells(INMAX, 48)).Interior.ColorIndex = XL_NONE
    ws.Range(ws.Cells(IN0, 2), ws.Cells(INMAX, 4)).NumberFormat = "@"
    If qN = 0 Then Exit Sub
    ReDim o(1 To qN, 1 To 6)
    ReDim ob(1 To qN, 1 To 1)
    For i = 1 To qN
        o(i, 1) = i
        o(i, 2) = qNom(i)
        o(i, 3) = qUnd(i)
        Select Case qTipo(i)
            Case "MAT"
                o(i, 4) = "MATERIAL"
                nMat = nMat + 1
            Case "SC"
                o(i, 4) = "SERVICIO"
                nSc = nSc + 1
            Case Else
                o(i, 4) = "EQUIPO"
                nEq = nEq + 1
        End Select
        o(i, 5) = qPre(i)
        k = Buscar(colOI, "I" & UCase(qNom(i)))
        ctrl = ""
        If k > 0 Then
            ctrl = UCase(Txt(oCtrl(k)))
            ob(i, 1) = oObs(k)
        End If
        If ctrl <> "" Then
            o(i, 6) = ctrl
        ElseIf NoControlar(qNom(i)) Then
            o(i, 6) = "NO"
        Else
            o(i, 6) = "SI"
        End If
    Next i
    ws.Range(ws.Cells(IN0, 1), ws.Cells(IN0 + qN - 1, 6)).Value = o
    ws.Range(ws.Cells(IN0, 48), ws.Cells(IN0 + qN - 1, 48)).Value = ob
    ws.Range(ws.Cells(IN0, 6), ws.Cells(IN0 + qN - 1, 6)).Interior.Color = cAma
    ws.Range(ws.Cells(IN0, 48), ws.Cells(IN0 + qN - 1, 48)).Interior.Color = cAma
End Sub


'========================================================================== MAQUINARIA Y COMBUSTIBLE
Private Sub ArmarMaquinaria(ByRef nComb As Long)
    Dim ws As Worksheet, r As Long, i As Long, idx As Long, ini As Long, nEq As Long
    Set ws = ThisWorkbook.Worksheets("MAQ_COMB")
    nComb = 0
    LimpiarZona ws, MQ0, MQMAX, 42
    r = MQ0
    Seccion ws, r, 42, "A. MAQUINARIA Y EQUIPOS (horas según A.C.U. frente al parte diario)"
    r = r + 1
    ini = r
    idx = 0
    For i = 1 To qN
        If qTipo(i) = "EQ" Then
            idx = idx + 1
            BloqueMaq ws, r, idx, i, LB_PARTE, LB_DIFP
            r = r + 3
        End If
    Next i
    nEq = idx
    If nEq = 0 Then
        ws.Cells(r, 2).Value = "(el A.C.U. no tiene equipos con horas)"
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    End If
    ws.Cells(r, 4).Value = LB_TOTEQ
    If nEq > 0 Then ws.Cells(r, 42).Formula = "=SUMIFS(AP" & ini & ":AP" & (r - 1) & ",$D" & ini & ":$D" & (r - 1) & ",""" & LB_ACU & """)"
    With ws.Range(ws.Cells(r, 4), ws.Cells(r, 42))
        .Interior.Color = cSubh
        .Font.Bold = True
    End With
    ws.Cells(r, 4).HorizontalAlignment = XL_RIGHT
    r = r + 2
    Seccion ws, r, 42, "B. COMBUSTIBLE (galones según A.C.U. frente a los vales de despacho)"
    r = r + 1
    idx = 0
    For i = 1 To qN
        If qTipo(i) = "MAT" Then
            If EsCombustible(qNom(i)) Then
                idx = idx + 1
                BloqueMaq ws, r, idx, i, LB_VALES, LB_DIFV
                r = r + 3
            End If
        End If
    Next i
    nComb = idx
    If nComb = 0 Then
        ws.Cells(r, 2).Value = "(no hay combustibles en el A.C.U.; revise las palabras clave en la hoja DATOS)"
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    End If
    r = r + 1
    ws.Cells(r, 2).Value = "Notas: 1) «SEGÚN A.C.U.» lo calcula ACTUALIZAR TODO con los metrados del RESUMEN. 2) Digite en las filas " & _
                           "amarillas las horas del parte diario y los galones de los vales. 3) Si la diferencia es negativa, el A.C.U. pide " & _
                           "más horas o galones que los reales: no valorizar por encima de lo sustentado."
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Size = 7
End Sub

Private Sub LimpiarZona(ByVal ws As Worksheet, ByVal f0 As Long, ByVal f1 As Long, ByVal nc As Long)
    With ws.Range(ws.Cells(f0, 1), ws.Cells(f1, nc))
        .ClearContents
        .Interior.ColorIndex = XL_NONE
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = 0
        .Font.Size = 7
        .Borders.LineStyle = XL_NONE
        .WrapText = False
        .HorizontalAlignment = 1
    End With
    ws.Range(ws.Cells(f0, 2), ws.Cells(f1, 2)).NumberFormat = "@"
End Sub

Private Sub Seccion(ByVal ws As Worksheet, ByVal r As Long, ByVal nc As Long, ByVal texto As String)
    ws.Cells(r, 1).Value = texto
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nc))
        .Interior.Color = cSubh
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = cAzul
    End With
End Sub

Private Sub BloqueMaq(ByVal ws As Worksheet, ByVal r As Long, ByVal idx As Long, ByVal i As Long, ByVal lblReal As String, ByVal lblDif As String)
    Dim ra As Long, rp As Long, rf As Long, kr As Long, k As Long, d As Long, o() As Variant
    ra = r
    rp = r + 1
    rf = r + 2
    kr = IN0 + i - 1
    ws.Cells(ra, 1).Value = idx
    ws.Cells(ra, 2).Value = qNom(i)
    ws.Cells(ra, 3).Value = qUnd(i)
    ws.Cells(ra, 4).Value = LB_ACU
    ws.Cells(rp, 4).Value = lblReal
    ws.Cells(rf, 4).Value = lblDif
    ws.Range(ws.Cells(rf, 8), ws.Cells(rf, 38)).FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C-R[-2]C)"
    ws.Cells(ra, 39).FormulaR1C1 = "=SUM(RC8:RC38)"
    ws.Cells(rp, 39).FormulaR1C1 = "=SUM(RC8:RC38)"
    ws.Cells(rf, 39).FormulaR1C1 = "=IF(COUNT(R[-1]C8:R[-1]C38)=0,"""",R[-1]C-R[-2]C)"
    ws.Cells(ra, 40).Formula = "=INSUMOS!$AR$" & kr
    ws.Cells(ra, 41).FormulaR1C1 = "=IF(RC5=0,"""",RC5-RC40)"
    ws.Cells(ra, 42).Formula = "=ROUND(AM" & ra & "*INSUMOS!$E$" & kr & ",2)"
    k = Buscar(colOM, "M" & UCase(qNom(i)))
    If k > 0 Then
        ReDim o(1 To 1, 1 To ND)
        For d = 1 To ND
            o(1, d) = oMq(k, d)
        Next d
        ws.Range(ws.Cells(rp, 8), ws.Cells(rp, 38)).Value = o
    End If
    ' formato del bloque
    With ws.Range(ws.Cells(ra, 1), ws.Cells(rf, 42))
        .Borders.LineStyle = XL_CONT
        .Borders.Weight = XL_THIN
        .Borders.Color = cBorde
        .NumberFormat = "#,##0.00;-#,##0.00;"""""
    End With
    On Error Resume Next
    ws.Range(ws.Cells(ra, 1), ws.Cells(rf, 3)).Borders(12).LineStyle = XL_NONE
    On Error GoTo 0
    ws.Cells(ra, 1).NumberFormat = "0"
    ws.Cells(ra, 1).HorizontalAlignment = XL_CENTER
    ws.Cells(ra, 1).Font.Bold = True
    ws.Cells(ra, 2).Font.Bold = True
    ws.Cells(ra, 2).NumberFormat = "@"
    ws.Cells(ra, 2).WrapText = True
    ws.Cells(ra, 3).HorizontalAlignment = XL_CENTER
    ws.Range(ws.Cells(ra, 4), ws.Cells(ra, 42)).Interior.Color = cCalc
    ws.Cells(ra, 4).Font.Bold = True
    ws.Cells(ra, 4).Font.Color = cAzul
    ws.Range(ws.Cells(ra, 6), ws.Cells(ra, 6)).NumberFormat = "[=73051]""SIN INGRESO"";dd/mm/yyyy"
    ws.Range(ws.Cells(ra, 7), ws.Cells(ra, 7)).NumberFormat = "[=73051]""—"";dd/mm/yyyy"
    ws.Range(ws.Cells(ra, 6), ws.Cells(ra, 7)).HorizontalAlignment = XL_CENTER
    ws.Cells(rp, 4).Interior.Color = cAma
    ws.Cells(rp, 4).Font.Bold = True
    ws.Cells(rp, 4).Font.Color = RGB(127, 96, 0)
    ws.Range(ws.Cells(rp, 8), ws.Cells(rp, 39)).Interior.Color = cAma
    ws.Cells(rf, 4).Font.Italic = True
    ws.Range(ws.Cells(rf, 8), ws.Cells(rf, 39)).Font.Italic = True
    ws.Range(ws.Cells(rf, 8), ws.Cells(rf, 39)).NumberFormat = "#,##0.00;[Red]-#,##0.00;""0.00"""
    ws.Cells(ra, 41).NumberFormat = "#,##0.00;[Red]-#,##0.00;""0.00"""
    ws.Cells(ra, 39).Font.Bold = True
    ws.Cells(rp, 39).Font.Bold = True
    ws.Cells(ra, 42).Font.Bold = True
End Sub


'========================================================================== MANO DE OBRA
Private Sub ArmarManoObra()
    Dim ws As Worksheet, r As Long, c As Long, k As Long, d As Long, o() As Variant, t1 As Long, filFin As Long
    Set ws = ThisWorkbook.Worksheets("MANO_OBRA")
    LimpiarZona ws, MO0, MOMAX, 39
    r = MO0
    For c = 1 To cN
        ws.Cells(r, 1).Value = c
        ws.Cells(r, 2).Value = cNomP(c)
        ws.Cells(r, 3).Value = "hh"
        ws.Cells(r + 1, 3).Value = "N°"
        ws.Cells(r + 2, 3).Value = "hh"
        ws.Cells(r, 4).Value = LB_HHACU
        ws.Cells(r + 1, 4).Value = LB_TRAB
        ws.Cells(r + 2, 4).Value = LB_HHTAR
        ws.Cells(r + 3, 4).Value = LB_USO
        ws.Range(ws.Cells(r + 2, 8), ws.Cells(r + 2, 38)).FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C*DATOS!R25C3)"
        ws.Range(ws.Cells(r + 3, 8), ws.Cells(r + 3, 38)).FormulaR1C1 = "=IF(N(R[-1]C)=0,IF(N(R[-3]C)>0,""S/P"",""""),N(R[-3]C)/R[-1]C)"
        ws.Cells(r, 39).FormulaR1C1 = "=SUM(RC8:RC38)"
        ws.Cells(r + 2, 39).FormulaR1C1 = "=SUM(RC8:RC38)"
        k = Buscar(colOT, "T" & UCase(cNomP(c)))
        If k > 0 Then
            ReDim o(1 To 1, 1 To ND)
            For d = 1 To ND
                o(1, d) = oTa(k, d)
            Next d
            ws.Range(ws.Cells(r + 1, 8), ws.Cells(r + 1, 38)).Value = o
        End If
        With ws.Range(ws.Cells(r, 1), ws.Cells(r + 3, 39))
            .Borders.LineStyle = XL_CONT
            .Borders.Weight = XL_THIN
            .Borders.Color = cBorde
            .NumberFormat = "#,##0.00;-#,##0.00;"""""
        End With
        On Error Resume Next
        ws.Range(ws.Cells(r, 1), ws.Cells(r + 3, 2)).Borders(12).LineStyle = XL_NONE
        On Error GoTo 0
        ws.Cells(r, 1).NumberFormat = "0"
        ws.Cells(r, 1).HorizontalAlignment = XL_CENTER
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Font.Bold = True
        ws.Cells(r, 2).Font.Size = 8
        ws.Range(ws.Cells(r, 3), ws.Cells(r + 3, 3)).HorizontalAlignment = XL_CENTER
        ws.Range(ws.Cells(r, 4), ws.Cells(r + 3, 4)).Font.Bold = True
        ws.Range(ws.Cells(r, 4), ws.Cells(r, 39)).Interior.Color = cCalc
        ws.Range(ws.Cells(r + 1, 4), ws.Cells(r + 1, 38)).Interior.Color = cAma
        ws.Range(ws.Cells(r + 1, 8), ws.Cells(r + 1, 38)).NumberFormat = "0;-0;"""""
        ws.Range(ws.Cells(r + 1, 8), ws.Cells(r + 1, 38)).HorizontalAlignment = XL_CENTER
        ws.Range(ws.Cells(r + 1, 8), ws.Cells(r + 1, 38)).Font.Bold = True
        ws.Range(ws.Cells(r + 2, 4), ws.Cells(r + 3, 39)).Interior.Color = cCalc
        ws.Range(ws.Cells(r + 3, 8), ws.Cells(r + 3, 38)).NumberFormat = "0%;-0%;"""""
        ws.Range(ws.Cells(r + 3, 8), ws.Cells(r + 3, 38)).HorizontalAlignment = XL_CENTER
        ws.Cells(r, 39).Font.Bold = True
        ws.Cells(r + 2, 39).Font.Bold = True
        r = r + 4
    Next c
    If cN = 0 Then
        ws.Cells(r, 2).Value = "(el A.C.U. no tiene mano de obra)"
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    End If
    filFin = r - 1
    r = r + 1
    t1 = r
    ws.Cells(t1, 2).Value = "TOTAL DEL DÍA"
    ws.Cells(t1, 4).Value = LB_THHACU
    ws.Cells(t1 + 1, 4).Value = LB_THHTAR
    ws.Cells(t1 + 2, 4).Value = LB_TUSO
    ws.Cells(t1 + 3, 4).Formula = "=""HH DISPONIBLES (USO MÁX. ""&TEXT(DATOS!$C$26,""0%"")&"")"""
    If filFin < MO0 Then filFin = MO0
    ws.Range(ws.Cells(t1, 8), ws.Cells(t1, 39)).FormulaR1C1 = "=SUMIFS(R" & MO0 & "C:R" & filFin & "C,R" & MO0 & "C4:R" & filFin & "C4,""" & LB_HHACU & """)"
    ws.Range(ws.Cells(t1 + 1, 8), ws.Cells(t1 + 1, 39)).FormulaR1C1 = "=SUMIFS(R" & MO0 & "C:R" & filFin & "C,R" & MO0 & "C4:R" & filFin & "C4,""" & LB_HHTAR & """)"
    ws.Range(ws.Cells(t1 + 2, 8), ws.Cells(t1 + 2, 39)).FormulaR1C1 = "=IF(R[-1]C=0,IF(R[-2]C>0,""SIN TAREO"",""""),R[-2]C/R[-1]C)"
    ws.Range(ws.Cells(t1 + 3, 8), ws.Cells(t1 + 3, 38)).FormulaR1C1 = "=R[-2]C*DATOS!R26C3"
    With ws.Range(ws.Cells(t1, 1), ws.Cells(t1 + 3, 39))
        .Interior.Color = cSubh
        .Font.Bold = True
        .Font.Size = 8
        .Borders.LineStyle = XL_CONT
        .Borders.Weight = XL_THIN
        .Borders.Color = cBorde
        .NumberFormat = "#,##0.00;-#,##0.00;"""""
    End With
    ws.Range(ws.Cells(t1 + 2, 8), ws.Cells(t1 + 2, 39)).NumberFormat = "0%;-0%;"""""
    ws.Range(ws.Cells(t1 + 2, 8), ws.Cells(t1 + 2, 39)).HorizontalAlignment = XL_CENTER
    ws.Range(ws.Cells(t1 + 3, 8), ws.Cells(t1 + 3, 39)).Font.Size = 7
    ws.Range(ws.Cells(t1 + 3, 39), ws.Cells(t1 + 3, 39)).ClearContents
    ws.Cells(t1, 2).HorizontalAlignment = XL_CENTER
    ws.Cells(t1 + 5, 4).Value = "S/P = el A.C.U. pide horas de esta categoría y el tareo no registra trabajadores de ella. " & _
                                "El control principal es el % DE USO TOTAL."
    ws.Cells(t1 + 5, 4).Font.Italic = True
End Sub
