Attribute VB_Name = "ModMenu"
Option Explicit

'==========================================================================================
'  CONTROL DE METRADOS DIARIOS - MENÚ, FIN DE MES Y NUEVA OBRA
'==========================================================================================


'========================================================================== NAVEGACIÓN
Public Sub IrInicio()
    IrA "INICIO"
End Sub

Public Sub IrResumen()
    IrA "RESUMEN"
End Sub

Private Sub IrA(ByVal nombre As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(nombre).Activate
    ThisWorkbook.Worksheets(nombre).Range("A1").Select
End Sub


'========================================================================== PDF DE LAS PLANILLAS
Public Sub ExportarPlanillasPDF()
    Dim msg As String, d As Long, n As Long, nom() As Variant, ruta As String, f As Variant, mes As Double, ws As Worksheet
    msg = Ejecutar()
    If Left(msg, 8) = "No se pu" Then
        MsgBox msg, vbExclamation, "Exportar planillas a PDF"
        Exit Sub
    End If
    ReDim nom(1 To ND)
    n = 0
    For d = 1 To ND
        Set ws = ThisWorkbook.Worksheets(Format(d, "00"))
        If ws.Visible = XL_VISIBLE Then
            n = n + 1
            nom(n) = ws.Name
        End If
    Next d
    If n = 0 Then
        MsgBox "No hay planillas con metrados en el mes.", vbInformation, "Exportar planillas a PDF"
        Exit Sub
    End If
    ReDim Preserve nom(1 To n)
    mes = Num(ThisWorkbook.Worksheets("DATOS").Range(D_MES).Value)
    ruta = "PLANILLAS_METRADOS_DIARIOS_" & Format(mes, "yyyy_mm") & ".pdf"
    If ThisWorkbook.Path <> "" Then ruta = ThisWorkbook.Path & Application.PathSeparator & ruta
    f = Application.GetSaveAsFilename(ruta, "Archivo PDF (*.pdf),*.pdf", , "Guardar las planillas del mes en PDF")
    If VarType(f) = vbBoolean Then Exit Sub
    ruta = CStr(f)
    On Error GoTo Fallo
    ThisWorkbook.Worksheets(nom).Select
    ActiveSheet.ExportAsFixedFormat Type:=0, Filename:=ruta, Quality:=0, IncludeDocProperties:=True, _
                                    IgnorePrintAreas:=False, OpenAfterPublish:=True
    ThisWorkbook.Worksheets("INICIO").Select
    MsgBox n & " planilla(s) guardada(s) en:" & vbCrLf & ruta, vbInformation, "Exportar planillas a PDF"
    Exit Sub
Fallo:
    On Error Resume Next
    ThisWorkbook.Worksheets("INICIO").Select
    MsgBox "No se pudo crear el PDF." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, _
           "Exportar planillas a PDF"
End Sub


'========================================================================== CERRAR EL MES
Public Sub CerrarMes()
    CerrarMesProceso False
End Sub

Public Sub CerrarMesSilencioso()
    gSilencio = True
    CerrarMesProceso True
    gSilencio = False
End Sub

Private Function SiguienteNumero(ByVal v As Variant) As Variant
    Dim s As String, n As Long
    s = Txt(v)
    If s = "" Then
        SiguienteNumero = ""
    ElseIf EsDigitos(s) Then
        n = CLng(s) + 1
        SiguienteNumero = Format(n, String(Len(s), "0"))
    Else
        SiguienteNumero = s
    End If
End Function

Private Sub CerrarMesProceso(ByVal sinPreguntar As Boolean)
    Dim ws As Worksheet, mes As Double, sig As Double, ruta As String, msg As String, n As Long, i As Long
    Dim v As Variant, h() As Variant, r As Long, calcMode As Long, nomBase As String, d As Long
    Set ws = ThisWorkbook.Worksheets("DATOS")
    mes = Num(ws.Range(D_MES).Value)
    If mes <= 0 Then
        If Not sinPreguntar Then MsgBox "Falta el MES QUE SE CONTROLA en la hoja DATOS.", vbExclamation, "Cerrar mes"
        Exit Sub
    End If
    sig = CDbl(DateSerial(Year(mes), Month(mes) + 1, 1))
    If Not sinPreguntar Then
        If MsgBox("Se cerrará el mes de " & NombreMes(mes) & " " & Year(mes) & ":" & vbCrLf & vbCrLf & _
                  "1) Se guarda una copia del libro tal como está (respaldo del mes)." & vbCrLf & _
                  "2) El METRADO ACUMULADO de cada partida pasa a METRADO ACUM. ANTERIOR." & vbCrLf & _
                  "3) Se borran los metrados de los días, los partes diarios, los vales y el tareo." & vbCrLf & _
                  "4) El mes pasa a " & NombreMes(sig) & " " & Year(sig) & " y la valorización al número siguiente." & vbCrLf & _
                  vbCrLf & "Los INGRESOS se conservan. ¿Continuar?", vbYesNo + vbQuestion, "Cerrar mes") <> vbYes Then Exit Sub
    End If
    msg = Ejecutar()
    If Left(msg, 8) = "No se pu" Then
        If Not sinPreguntar Then MsgBox msg, vbExclamation, "Cerrar mes"
        Exit Sub
    End If
    ' copia de respaldo del mes
    ruta = ""
    If ThisWorkbook.Path <> "" And Not sinPreguntar Then
        nomBase = ThisWorkbook.Name
        If InStrRev(nomBase, ".") > 0 Then nomBase = Left(nomBase, InStrRev(nomBase, ".") - 1)
        ruta = ThisWorkbook.Path & Application.PathSeparator & nomBase & "_CIERRE_" & Format(mes, "yyyy_mm") & ".xlsm"
        On Error Resume Next
        ThisWorkbook.SaveCopyAs ruta
        If Err.Number <> 0 Then ruta = ""
        On Error GoTo 0
    End If
    Inicio calcMode, "Cerrando el mes..."
    ' el acumulado pasa a acumulado anterior y se borran los días
    Set ws = ThisWorkbook.Worksheets("RESUMEN")
    n = UltimaFila(ws, 2, RR0, RRMAX) - RR0 + 1
    If n > 0 Then
        Application.Calculate
        v = ws.Range(ws.Cells(RR0, 1), ws.Cells(RR0 + n - 1, 41)).Value
        ReDim h(1 To n, 1 To 1)
        For i = 1 To n
            If UCase(Txt(v(i, 1))) = "P" Then
                h(i, 1) = Num(v(i, 8))
                For d = 1 To ND
                    h(i, 1) = h(i, 1) + Num(v(i, CD1 + d - 1))
                Next d
            End If
        Next i
        ws.Range(ws.Cells(RR0, 8), ws.Cells(RR0 + n - 1, 8)).Value = h
        ws.Range(ws.Cells(RR0, CD1), ws.Cells(RR0 + n - 1, CD1 + ND - 1)).ClearContents
    End If
    ' partes diarios, vales y tareo
    Set ws = ThisWorkbook.Worksheets("MAQ_COMB")
    v = ws.Range(ws.Cells(1, 4), ws.Cells(MQMAX, 4)).Value
    For r = MQ0 To MQMAX
        If Txt(v(r, 1)) = LB_PARTE Or Txt(v(r, 1)) = LB_VALES Then ws.Range(ws.Cells(r, 8), ws.Cells(r, 38)).ClearContents
    Next r
    Set ws = ThisWorkbook.Worksheets("MANO_OBRA")
    v = ws.Range(ws.Cells(1, 4), ws.Cells(MOMAX, 4)).Value
    For r = MO0 To MOMAX
        If Txt(v(r, 1)) = LB_TRAB Then ws.Range(ws.Cells(r, 8), ws.Cells(r, 38)).ClearContents
    Next r
    ' mes siguiente
    With ThisWorkbook.Worksheets("DATOS")
        .Range(D_MES).Value = DateSerial(Year(sig), Month(sig), 1)
        .Range(D_VALN).Value = SiguienteNumero(.Range(D_VALN).Value)
    End With
    Fin calcMode
    msg = Ejecutar()
    If Not sinPreguntar Then
        ThisWorkbook.Worksheets("INICIO").Activate
        MsgBox "Mes cerrado. Ahora se controla " & NombreMes(sig) & " " & Year(sig) & "." & vbCrLf & vbCrLf & _
               IIf(ruta <> "", "Copia de respaldo del mes anterior:" & vbCrLf & ruta, "No se guardó copia de respaldo " & _
               "(guarde el libro primero).") & vbCrLf & vbCrLf & msg, vbInformation, "Cerrar mes"
    Else
        ThisWorkbook.Worksheets("DATOS").Range(D_MSGACT).Value = msg
    End If
End Sub


'========================================================================== NUEVA OBRA
Public Sub NuevaObra()
    Dim ruta As String, nomBase As String, calcMode As Long, d As Long, ws As Worksheet
    If MsgBox("Se VACIARÁ el programa para cargar OTRA OBRA: datos de la obra, presupuesto, A.C.U., resumen, " & _
              "ingresos, insumos, maquinaria, mano de obra y planillas." & vbCrLf & vbCrLf & _
              "Antes se guardará una copia de respaldo del libro tal como está." & vbCrLf & vbCrLf & "¿Continuar?", _
              vbYesNo + vbExclamation, "Nueva obra") <> vbYes Then Exit Sub
    If MsgBox("¿Está seguro? Esta acción borra todos los datos de la obra actual de este libro.", _
              vbYesNo + vbCritical + vbDefaultButton2, "Nueva obra") <> vbYes Then Exit Sub
    If ThisWorkbook.Path <> "" Then
        nomBase = ThisWorkbook.Name
        If InStrRev(nomBase, ".") > 0 Then nomBase = Left(nomBase, InStrRev(nomBase, ".") - 1)
        ruta = ThisWorkbook.Path & Application.PathSeparator & nomBase & "_RESPALDO_" & Format(Now, "yyyymmdd_hhmm") & ".xlsm"
        On Error Resume Next
        ThisWorkbook.SaveCopyAs ruta
        If Err.Number <> 0 Then
            On Error GoTo 0
            If MsgBox("No se pudo guardar la copia de respaldo. ¿Vaciar de todas formas?", vbYesNo + vbCritical, _
                      "Nueva obra") <> vbYes Then Exit Sub
            ruta = ""
        End If
        On Error GoTo 0
    End If
    VaciarTodo
    ThisWorkbook.Worksheets("DATOS").Activate
    ThisWorkbook.Worksheets("DATOS").Range(D_ENTIDAD).Select
    MsgBox "Programa vacío y listo para otra obra." & vbCrLf & vbCrLf & _
           IIf(ruta <> "", "Respaldo de la obra anterior:" & vbCrLf & ruta & vbCrLf & vbCrLf, "") & _
           "Siga los pasos del INICIO: 1) datos de la obra, 2) presupuesto, 3) A.C.U., 4) PREPARAR OBRA.", vbInformation, "Nueva obra"
End Sub

Public Sub VaciarTodo()
    Dim calcMode As Long, d As Long, ws As Worksheet, celdas As Variant, i As Long
    Inicio calcMode, "Vaciando el programa..."
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DATOS")
    celdas = Array(D_ENTIDAD, D_DEPEND, D_PROY, D_CUI, D_META, D_UBIC, D_MODAL, D_INICIO, D_PLAZO, D_MES, D_VALN, _
                   D_ULTACT, D_HUELLA, D_MSGACT, D_ULTPREP, D_HUELLAP, D_MSGPREP)
    For i = LBound(celdas) To UBound(celdas)
        ws.Range(celdas(i)).ClearContents
    Next i
    With ThisWorkbook.Worksheets("PRESUPUESTO")
        .Range(.Cells(PP0, 1), .Cells(PPMAX, 9)).ClearContents
    End With
    With ThisWorkbook.Worksheets("ACU")
        .Range(.Cells(AC0, 1), .Cells(ACMAX, 16)).ClearContents
    End With
    With ThisWorkbook.Worksheets("RESUMEN")
        With .Range(.Cells(RR0, 1), .Cells(RRMAX, 52))
            .ClearContents
            .Interior.ColorIndex = XL_NONE
            .Font.Bold = False
            .Font.Color = 0
        End With
    End With
    With ThisWorkbook.Worksheets("INGRESOS")
        .Range(.Cells(IG0, 2), .Cells(IG1, 9)).ClearContents
        .Range(.Cells(IG0, 11), .Cells(IG1, 11)).ClearContents
        .Range(.Cells(IG0, 13), .Cells(IG1, 16)).ClearContents
    End With
    With ThisWorkbook.Worksheets("INSUMOS")
        .Range(.Cells(IN0, 1), .Cells(INMAX, 49)).ClearContents
        .Range(.Cells(IN0, 6), .Cells(INMAX, 6)).Interior.ColorIndex = XL_NONE
        .Range(.Cells(IN0, 48), .Cells(INMAX, 48)).Interior.ColorIndex = XL_NONE
    End With
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "MAQ_COMB" Or ws.Name = "MANO_OBRA" Then
            With ws.Range(ws.Cells(MQ0, 1), ws.Cells(MQMAX, 42))
                .ClearContents
                .Interior.ColorIndex = XL_NONE
                .Borders.LineStyle = XL_NONE
                .Font.Bold = False
                .Font.Italic = False
                .Font.Color = 0
            End With
        End If
    Next ws
    ThisWorkbook.Worksheets("AUX").Range("A" & RR0 & ":CP" & RRMAX).ClearContents
    For d = 1 To ND
        Set ws = ThisWorkbook.Worksheets(Format(d, "00"))
        With ws.Range(ws.Cells(PL0, 1), ws.Cells(PLMAX, 11))
            .ClearContents
            .Interior.ColorIndex = XL_NONE
            .Borders.LineStyle = XL_NONE
            .Font.Bold = False
            .Font.Color = 0
        End With
        ws.Range("A2").Value = "PLANILLA DE METRADOS DIARIOS N° " & Format(d, "00")
        ws.Range("C5").ClearContents
        ws.Range("D5").ClearContents
        ws.Range("I5").ClearContents
        ws.Range("K5").Value = "SIN METRADOS REGISTRADOS"
        ws.Range("D6").Value = 0
        ws.Range("I6").Value = 0
        ws.Visible = XL_HIDDEN
    Next d
    On Error GoTo 0
    Fin calcMode
End Sub
