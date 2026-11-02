Attribute VB_Name = "ModInforme"
Option Explicit

'==========================================================================================
'  ENVIAR LOS DATOS AL EXCEL DEL INFORME (hojas BD_ del informe 4.10 y 4.15, con fórmulas)
'==========================================================================================

Public Sub EnviarInforme()
    Dim msg As String, s As String, a1 As Boolean, a2 As Boolean, conPdf As Boolean
    msg = ""
    a1 = Pregunta("¿Enviar el CATÁLOGO y los MOVIMIENTOS al Excel del informe 4.10 (almacén de obra)?", _
                  "ENVIAR AL INFORME", True)
    a2 = Pregunta("¿Enviar los EQUIPOS y los PARTES al Excel del informe 4.15 (control de equipo mecánico)?", _
                  "ENVIAR AL INFORME", True)
    If Not (a1 Or a2) Then Exit Sub
    conPdf = Pregunta("¿Exportar también a PDF las hojas del informe?" & vbCrLf & vbCrLf & _
                      "Se guardan en la carpeta de salidas del mes (subcarpeta INFORME) para revisarlas.", _
                      "ENVIAR AL INFORME", True)
    If a1 Then
        msg = EnviarA(True, "", conPdf)
    End If
    If a2 Then
        s = EnviarA(False, "", conPdf)
        If msg = "" Then
            msg = s
        Else
            msg = msg & vbCrLf & vbCrLf & s
        End If
    End If
    If msg <> "" Then Aviso msg, "ENVIAR AL INFORME"
End Sub


Private Sub CopiarDatos(ByVal wb As Workbook)
    Dim ws As Worksheet, src As Worksheet, c As Variant
    If Not TieneHoja(wb, "DATOS") Then Exit Sub
    Set ws = wb.Worksheets("DATOS")
    Set src = Hoja("DATOS")
    For Each c In Array(D_ENTIDAD, D_GEREN, D_SUBG, D_PROY, D_CUI, D_META, D_UBIC, D_MODAL, D_NINF)
        If Txt(src.Range(c).Value) <> "" Then
            ws.Range(c).NumberFormat = "@"
            ws.Range(c).Value = Txt(src.Range(c).Value)
        End If
    Next c
    ws.Range(D_MES).Value = CDate(MesInicio())
    ws.Range(D_CORTE).Value = CDate(FechaCorte())
End Sub

Private Function CopiarBloque(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal f0 As Long, ByVal fmax As Long, _
                              ByVal colKey As Long, ByVal ucol As Long) As Long
    Dim last As Long, v As Variant, i As Long, n As Long
    CopiarBloque = 0
    last = UltimaFila(src, colKey, f0, fmax)
    dst.Range(dst.Cells(f0, 1), dst.Cells(fmax, ucol)).ClearContents
    If last < f0 Then Exit Function
    v = src.Range(src.Cells(f0, 1), src.Cells(last, ucol)).Value
    For i = 1 To UBound(v, 1)
        If IsError(v(i, 1)) Then v(i, 1) = Empty
        If Txt(v(i, colKey)) <> "" Then n = n + 1
    Next i
    dst.Range(dst.Cells(f0, 1), dst.Cells(last, ucol)).Value = v
    CopiarBloque = n
End Function

Private Function CopiarMovimientos(ByVal dst As Worksheet) As Long
    ' movimientos válidos en orden cronológico
    Dim src As Worksheet, last As Long, v As Variant, o() As Variant, j As Long, i As Long, c As Long
    CopiarMovimientos = 0
    Set src = Hoja("MOVIMIENTOS")
    dst.Range(dst.Cells(MV0, 1), dst.Cells(MVMAX, MC_ULT)).ClearContents
    LeerAlmacen
    If anMov = 0 Then Exit Function
    last = UltimaFilaVarias(src, MC_FEC, MC_CANT, MV0, MVMAX)
    v = src.Range(src.Cells(MV0, 1), src.Cells(last, MC_ULT)).Value
    ReDim o(1 To anMov, 1 To MC_ULT)
    For j = 1 To anMov
        i = vFila(vOrden(j)) - MV0 + 1
        For c = 1 To MC_ULT
            If IsError(v(i, c)) Then
                o(j, c) = Empty
            Else
                o(j, c) = v(i, c)
            End If
        Next c
        o(j, 1) = j
    Next j
    dst.Range(dst.Cells(MV0, 1), dst.Cells(MV0 + anMov - 1, MC_ULT)).Value = o
    CopiarMovimientos = anMov
End Function

Private Function CopiarPartes(ByVal dst As Worksheet) As Long
    ' partes válidos ordenados por equipo y fecha
    Dim o() As Variant, j As Long, i As Long, c As Long
    CopiarPartes = 0
    dst.Range(dst.Cells(PT0, 1), dst.Cells(PTMAX, PC_ULT)).ClearContents
    LeerPartes
    If pnPar = 0 Then Exit Function
    ReDim o(1 To pnPar, 1 To PC_ULT)
    For j = 1 To pnPar
        i = pR(pOrden(j))
        For c = 1 To PC_ULT
            If IsError(pV(i, c)) Then
                o(j, c) = Empty
            Else
                o(j, c) = pV(i, c)
            End If
        Next c
        o(j, 1) = j
    Next j
    dst.Range(dst.Cells(PT0, 1), dst.Cells(PT0 + pnPar - 1, PC_ULT)).Value = o
    CopiarPartes = pnPar
End Function


Private Function HojaDelInforme(ByVal nom As String) As Boolean
    ' hojas que van al PDF del informe (no DATOS, BD_, AUX ni las consultas KARDEX / BINCARD)
    Dim u As String
    u = UCase(Trim(nom))
    HojaDelInforme = Not (u = "DATOS" Or Left(u, 3) = "BD_" Or u = "AUX" Or u = "KARDEX" Or u = "BINCARD")
End Function

Private Sub AplicarFiltros(ByVal wb As Workbook)
    ' recalcula y vuelve a aplicar el filtro MOSTRAR de cada hoja (oculta las filas sin datos)
    Dim ws As Worksheet
    On Error Resume Next
    Application.Calculate
    For Each ws In wb.Worksheets
        If ws.AutoFilterMode Then ws.AutoFilter.ApplyFilter
    Next ws
    On Error GoTo 0
End Sub

Private Function PdfInforme(ByVal wb As Workbook) As String
    ' PDF de las hojas del informe en SALIDAS_AAAA_MM\INFORME (las demás hojas se ocultan un momento)
    Dim ws As Worksheet, ocultas As Collection, archivo As String, carpeta As String, k As Long, nom As String
    PdfInforme = ""
    If gSinPDF Then Exit Function
    carpeta = CarpetaSalida("INFORME")
    If carpeta = "" Then Exit Function
    nom = wb.Name
    k = InStrRev(nom, ".")
    If k > 0 Then nom = Left(nom, k - 1)
    archivo = carpeta & Sep() & NombreSeguro(nom, 120) & ".pdf"
    Set ocultas = New Collection
    On Error GoTo Falla
    For Each ws In wb.Worksheets
        If ws.Visible = -1 Then
            If Not HojaDelInforme(ws.Name) Then
                ws.Visible = 0
                ocultas.Add ws.Name
            End If
        End If
    Next ws
    wb.ExportAsFixedFormat Type:=0, Filename:=archivo, Quality:=0, IncludeDocProperties:=False, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=False
    PdfInforme = archivo
Restaurar:
    On Error Resume Next
    For k = 1 To ocultas.Count
        wb.Worksheets(CStr(ocultas(k))).Visible = -1
    Next k
    Exit Function
Falla:
    Resume Restaurar
End Function

Public Function EnviarA(ByVal almacen As Boolean, ByVal ruta As String, Optional ByVal conPdf As Boolean = False) As String
    Dim celda As String, r As Variant, wb As Workbook, ya As Boolean, cm As Long, h1 As String, h2 As String
    Dim n1 As Long, n2 As Long, nombre As String, msg As String, soloLec As Boolean, pdf As String

    If almacen Then
        celda = D_INF410
        h1 = "BD_CATALOGO"
        h2 = "BD_MOVIMIENTOS"
        nombre = "4.10 (ALMACÉN DE OBRA)"
    Else
        celda = D_INF415
        h1 = "BD_EQUIPOS"
        h2 = "BD_PARTES"
        nombre = "4.15 (CONTROL DE EQUIPO MECÁNICO)"
    End If
    If ruta = "" Then ruta = Txt(Hoja("DATOS").Range(celda).Value)
    If ruta <> "" Then
        On Error Resume Next
        If Dir(ruta) = "" Then ruta = ""
        On Error GoTo 0
    End If
    If ruta = "" Then
        If gSilencio Then
            EnviarA = "Informe " & nombre & ": no se indicó el archivo."
            Exit Function
        End If
        r = Application.GetOpenFilename("Libros de Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "Elija el Excel del informe " & nombre)
        If VarType(r) = vbBoolean Then
            EnviarA = "Informe " & nombre & ": envío cancelado."
            Exit Function
        End If
        ruta = CStr(r)
    End If

    Inicio cm, "Enviando los datos al informe " & nombre & "..."
    On Error GoTo Falla
    Set wb = AbrirLibroX(ruta, ya, False)
    If wb Is Nothing Then Err.Raise 3201, , "No se pudo abrir el archivo."
    soloLec = False
    On Error Resume Next
    soloLec = wb.ReadOnly
    On Error GoTo Falla
    If soloLec Then Err.Raise 3202, , "El archivo está abierto solo para lectura (¿lo tiene abierto otra persona?)."
    If Not (TieneHoja(wb, h1) And TieneHoja(wb, h2)) Then
        Err.Raise 3203, , "El archivo no tiene las hojas " & h1 & " y " & h2 & ": use la plantilla del informe con fórmulas."
    End If
    CopiarDatos wb
    If almacen Then
        n1 = CopiarBloque(Hoja("CATALOGO"), wb.Worksheets(h1), CT0, CTMAX, CC_COD, CC_ULT)
        n2 = CopiarMovimientos(wb.Worksheets(h2))
    Else
        n1 = CopiarBloque(Hoja("EQUIPOS"), wb.Worksheets(h1), EQ0, EQMAX, EC_COD, EC_ULT)
        n2 = CopiarPartes(wb.Worksheets(h2))
    End If
    AplicarFiltros wb
    wb.Save
    pdf = ""
    If conPdf Then pdf = PdfInforme(wb)
    If Not ya Then wb.Close False
    Set wb = Nothing
    On Error GoTo 0
    Fin cm
    Hoja("DATOS").Range(celda).Value = ruta
    If almacen Then
        msg = "Informe " & nombre & ": se enviaron " & n1 & " insumos y " & n2 & " movimientos."
        If anMal > 0 Then msg = msg & vbCrLf & "   (" & anMal & " filas de MOVIMIENTOS incompletas no se enviaron)"
    Else
        msg = "Informe " & nombre & ": se enviaron " & n1 & " equipos y " & n2 & " filas de partes."
        If pnMal > 0 Then msg = msg & vbCrLf & "   (" & pnMal & " filas de PARTES incompletas no se enviaron)"
    End If
    msg = msg & vbCrLf & "   " & ruta
    If conPdf And Not gSinPDF Then
        If pdf <> "" Then
            msg = msg & vbCrLf & "   PDF: " & pdf
        Else
            msg = msg & vbCrLf & "   (no se pudo exportar el PDF del informe: ¿está abierto en otro programa?)"
        End If
    End If
    Mensaje D_MSGINF, "Enviado el " & Format(Now, "dd/mm/yyyy hh:mm") & " al informe " & nombre
    EnviarA = msg
    Exit Function
Falla:
    msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then
        If Not ya Then wb.Close False
    End If
    Fin cm
    EnviarA = "Informe " & nombre & ": no se pudo enviar. " & msg
    Mensaje D_MSGINF, "No se pudo enviar al informe " & nombre & " (" & Format(Now, "dd/mm/yyyy hh:mm") & "): " & msg
End Function
