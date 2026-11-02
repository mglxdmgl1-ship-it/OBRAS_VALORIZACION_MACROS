Attribute VB_Name = "ModPDF"
Option Explicit

'==========================================================================================
'  PDF E IMPRESIÓN
'  Los formatos se arman primero en sus hojas (botones GENERAR), se revisan y recién se exportan
'  (botones PDF). Cada página se escala para que entre completa en el papel, con márgenes, sin cortes.
'  Bincard y partes diarios: en A5 (uno por hoja) o en A4 horizontal con dos por cara
'  (dos distintos o dos ejemplares del mismo: original y copia).
'==========================================================================================

Public Const MK_CLAVE As Long = 30            ' columna AD: código y descripción al inicio de cada tarjeta / equipo
Public Const MK_FIN As Long = 31              ' columna AE: última fila de esa tarjeta / equipo
Public Const MARGEN_PT As Double = 17         ' 0.6 cm de margen en cada lado (evita cortes de la impresora)
Public Const HOJA_A4 As String = "IMPRESION_A4"

Public Const IMP_NORMAL As Long = 1           ' en su papel: kardex A4 horizontal; bincard y parte A5 vertical
Public Const IMP_DOS As Long = 2              ' A4 horizontal: dos formatos distintos por cara
Public Const IMP_COPIA As Long = 3            ' A4 horizontal: dos ejemplares del mismo formato (original y copia)

Public gImpresion As Long                     ' forma de impresión ya elegida (PDF DE TODO y pruebas); 0 = preguntar
Public gPorUno As Integer                     ' PDF DE TODO: 1 = también un PDF por insumo/equipo, 2 = no; 0 = preguntar


'========================================================================== MARCAS DE CADA TARJETA
Public Sub MarcarBloque(ByVal ws As Worksheet, ByVal fila As Long, ByVal clave As String, ByVal fin As Long)
    ws.Cells(fila, MK_CLAVE).Value = clave
    ws.Cells(fila, MK_FIN).Value = fin
End Sub

Public Sub OcultarMarcas(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(ws.Cells(1, MK_CLAVE), ws.Cells(1, MK_FIN)).EntireColumn.Hidden = True
End Sub

Public Function FilasArmadas(ByVal ws As Worksheet) As Long
    ' última fila armada de la hoja de salida (0 si está vacía)
    Dim r As Long, i As Long, v As Variant, m As Double
    FilasArmadas = 0
    r = ws.Cells(ws.Rows.Count, MK_CLAVE).End(-4162).Row
    If Txt(ws.Cells(r, MK_CLAVE).Value) = "" Then Exit Function
    If r = 1 Then
        FilasArmadas = CLng(Num(ws.Cells(1, MK_FIN).Value))
        Exit Function
    End If
    v = ws.Range(ws.Cells(1, MK_FIN), ws.Cells(r, MK_FIN)).Value
    For i = 1 To r
        If Num(v(i, 1)) > m Then m = Num(v(i, 1))
    Next i
    FilasArmadas = CLng(m)
End Function


Public Sub BorrarFilas(ByVal ws As Worksheet)
    ' borra las filas usadas (alto de filas y contenido vuelven a lo normal)
    Dim u As Long
    On Error Resume Next
    u = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If u < 1 Then u = 1
    ws.Range(ws.Rows(1), ws.Rows(u)).Delete
End Sub


'========================================================================== PÁGINA SIN CORTES
Private Sub ComunicarImpresora(ByVal si As Boolean)
    ' Excel 2010+: con False los cambios de página se aplican juntos (más rápido)
    If gSilencio Then Exit Sub                    ' pruebas automáticas (LibreOffice no la tiene)
    On Error Resume Next
    Application.PrintCommunication = si
End Sub

Public Sub AjustarPagina(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, ByVal alto As Long, _
                         ByVal papel As Long, ByVal horizontal As Boolean)
    ' una página = filas 1..alto y columnas c1..c2; se calcula la escala para que entre completa en el papel
    Dim aP As Double, hP As Double, w As Double, h As Double, z As Double, t As Double, papReal As Long
    On Error Resume Next
    ComunicarImpresora True
    With ws.PageSetup
        .PaperSize = papel
        .Orientation = IIf(horizontal, 2, 1)
    End With
    papReal = ws.PageSetup.PaperSize                 ' si la impresora no tiene ese papel queda el que acepta
    If Err.Number <> 0 Then papReal = papel
    Err.Clear
    Select Case papReal
        Case 11  ' A5
            aP = 419.53
            hP = 595.28
        Case 1  ' carta
            aP = 612
            hP = 792
        Case 5  ' oficio
            aP = 612
            hP = 1008
        Case Else  ' A4
            aP = 595.28
            hP = 841.89
    End Select
    If horizontal Then
        t = aP
        aP = hP
        hP = t
    End If
    w = ws.Range(ws.Cells(1, c1), ws.Cells(1, c2)).Width
    h = ws.Range(ws.Cells(1, 1), ws.Cells(alto, 1)).Height
    z = 100
    If w > 0 Then z = (aP - 2 * MARGEN_PT) / w * 100
    If h > 0 Then
        t = (hP - 2 * MARGEN_PT) / h * 100
        If t < z Then z = t
    End If
    z = Int(z * 0.97)                             ' holgura para la diferencia entre pantalla e impresión
    If z < 10 Then z = 10
    If z > 400 Then z = 400
    ComunicarImpresora False
    With ws.PageSetup
        .Zoom = z
        .LeftMargin = MARGEN_PT
        .RightMargin = MARGEN_PT
        .TopMargin = MARGEN_PT
        .BottomMargin = MARGEN_PT
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .CenterFooter = ""
        .LeftFooter = ""
        .RightFooter = ""
    End With
    ComunicarImpresora True
    On Error GoTo 0
End Sub

Public Sub PaginaFormato(ByVal ws As Worksheet, ByVal cual As String)
    ' papel de cada formato impreso uno por hoja
    Select Case cual
        Case "KARDEX"
            AjustarPagina Hoja("KARDEX"), 2, 16, KX_ALTO, 9, True
        Case "BINCARD"
            AjustarPagina Hoja("BINCARD"), 2, 16, BC_ALTO, 11, False
        Case Else
            AjustarPagina Hoja("PARTE_DIARIO"), 2, 16, PD_ALTO, 11, False
    End Select
End Sub

Public Sub AjustarInforme(ByVal ws As Worksheet)
    ' cuadros del informe (INV., movimiento, ingresos – egresos, saldo valorizado, control de equipo): todo el ancho en una página, títulos repetidos, sin cortes
    Dim tit As String, nm As String, u As Long, horiz As Boolean
    nm = UCase(ws.Name)
    horiz = True
    If Left(nm, 4) = "INV." Then
        tit = "$1:$1"
        horiz = False
        u = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
        If u < 2 Then u = 2
        On Error Resume Next
        ws.PageSetup.PrintArea = "$A$1:$J$" & u
        On Error GoTo 0
    ElseIf Left(nm, 7) = "CONTROL" Then
        tit = "$1:$13"
    Else
        tit = "$1:$12"
    End If
    On Error Resume Next
    ComunicarImpresora False
    With ws.PageSetup
        .PaperSize = 9
        .Orientation = IIf(horiz, 2, 1)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = MARGEN_PT
        .RightMargin = MARGEN_PT
        .TopMargin = MARGEN_PT * 1.2
        .BottomMargin = MARGEN_PT * 1.6
        .HeaderMargin = 0
        .FooterMargin = MARGEN_PT * 0.5
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintTitleRows = tit
        .RightFooter = "Página &P de &N"
    End With
    ComunicarImpresora True
    On Error GoTo 0
End Sub


'========================================================================== A4 CON DOS POR CARA
Private Function HojaImpresion(ByVal despuesDe As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_A4)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=despuesDe)
        ws.Name = HOJA_A4
        On Error Resume Next
        ws.Tab.Color = RGB(192, 0, 0)
        ActiveWindow.DisplayGridlines = False
        On Error GoTo 0
    End If
    On Error Resume Next
    ws.ResetAllPageBreaks
    BorrarFormas ws
    BorrarFilas ws
    On Error GoTo 0
    ws.Cells.Clear
    Set HojaImpresion = ws
End Function

Public Function ArmarDoble(ByVal wsS As Worksheet, ByVal alto As Long, ByVal ult As Long, ByVal copia As Boolean) As Worksheet
    ' hoja IMPRESION_A4: A4 horizontal con dos páginas del formato (columnas B..P) una al lado de la otra
    Dim ws As Worksheet, nP As Long, nQ As Long, q As Long, a As Long, b As Long, r0 As Long, i As Long, c As Long
    Dim copObj As Boolean
    Set ws = HojaImpresion(wsS)
    nP = (ult + alto - 1) \ alto
    If copia Then nQ = nP Else nQ = (nP + 1) \ 2
    For c = 2 To 16
        ws.Columns(c - 1).ColumnWidth = wsS.Columns(c).ColumnWidth
        ws.Columns(c + 16).ColumnWidth = wsS.Columns(c).ColumnWidth
    Next c
    ws.Columns(16).ColumnWidth = 1.5
    ws.Columns(17).ColumnWidth = 1.5
    copObj = True
    On Error Resume Next
    copObj = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = True
    On Error GoTo 0
    For q = 1 To nQ
        If q Mod 10 = 0 Then Estado "Armando la hoja " & HOJA_A4 & ": " & q & " de " & nQ & " ..."
        If copia Then
            a = q
            b = q
        Else
            a = 2 * q - 1
            b = 2 * q
        End If
        r0 = (q - 1) * alto + 1
        For i = 1 To alto
            ws.Rows(r0 + i - 1).RowHeight = wsS.Rows((a - 1) * alto + i).RowHeight
        Next i
        wsS.Range(wsS.Cells((a - 1) * alto + 1, 2), wsS.Cells(a * alto, 16)).Copy Destination:=ws.Cells(r0, 1)
        If b <= nP Then
            wsS.Range(wsS.Cells((b - 1) * alto + 1, 2), wsS.Cells(b * alto, 16)).Copy Destination:=ws.Cells(r0, 18)
        End If
        ' línea de corte entre los dos ejemplares
        On Error Resume Next
        With ws.Range(ws.Cells(r0, 16), ws.Cells(r0 + alto - 1, 16)).Borders(10)
            .LineStyle = -4115
            .Weight = 2
            .Color = RGB(150, 150, 150)
        End With
        If r0 > 1 Then ws.HPageBreaks.Add Before:=ws.Cells(r0, 1)
        On Error GoTo 0
    Next q
    On Error Resume Next
    Application.CutCopyMode = False
    Application.CopyObjectsWithCells = copObj
    ws.PageSetup.PrintArea = "$A$1:$AF$" & (nQ * alto)
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets(HOJA_A4)
    AjustarPagina ws, 1, 32, alto, 9, True
    Set ArmarDoble = ws
End Function


'========================================================================== ELEGIR LA FORMA DE IMPRESIÓN
Public Function ElegirImpresion(ByVal nombre As String) As Long
    Dim v As Variant
    If gImpresion > 0 Then
        ElegirImpresion = gImpresion
        Exit Function
    End If
    If gSilencio Then
        ElegirImpresion = IMP_NORMAL
        Exit Function
    End If
    v = Application.InputBox("¿Cómo desea el PDF de " & nombre & "?" & vbCrLf & vbCrLf & _
                             "1 = En A5: un formato por hoja" & vbCrLf & _
                             "2 = En A4 horizontal: dos formatos distintos por cara (ahorra papel)" & vbCrLf & _
                             "3 = En A4 horizontal: dos ejemplares del mismo formato (original y copia)" & vbCrLf & vbCrLf & _
                             "Escriba 1, 2 o 3:", nombre & " - FORMA DE IMPRESIÓN", 1, Type:=1)
    If VarType(v) = vbBoolean Then
        ElegirImpresion = 0
        Exit Function
    End If
    If v < 1 Or v > 3 Then v = 1
    ElegirImpresion = CLng(v)
End Function


'========================================================================== EXPORTAR UN FORMATO
Private Sub ArmarAhora(ByVal cual As String)
    ' arma la hoja sin mostrar los avisos intermedios
    Dim acum As Boolean, txt0 As String, s As String, n As Long
    acum = gAcum
    txt0 = gTexto
    gAcum = True
    s = ActualizarCuadros(True)
    Select Case cual
        Case "KARDEX"
            n = ArmarTarjetas(True, False, False)
        Case "BINCARD"
            n = ArmarTarjetas(False, False, False)
        Case Else
            n = ArmarPartes(False, False)
    End Select
    If Not acum Then
        gAcum = False
        gTexto = txt0
    End If
End Sub

Public Function ExportarFormato(ByVal cual As String) As String
    ' cual = KARDEX, BINCARD o PARTES; exporta lo que está en la hoja (con las correcciones que se hayan hecho)
    Dim ws As Worksheet, wsX As Worksheet, alto As Long, ult As Long, modo As Long, porUno As Boolean
    Dim nombre As String, subc As String, que As String, suf As String, archivo As String, carpeta As String
    Dim msg As String, nPdf As Long, r As Long, area As String, ok As Boolean, sub2 As String, nPag As Long

    ExportarFormato = ""
    Select Case cual
        Case "KARDEX"
            Set ws = Hoja("KARDEX")
            alto = KX_ALTO
            nombre = "KARDEX"
            subc = "POR_INSUMO"
            que = "insumo"
        Case "BINCARD"
            Set ws = Hoja("BINCARD")
            alto = BC_ALTO
            nombre = "BINCARD"
            subc = "POR_INSUMO"
            que = "insumo"
        Case Else
            Set ws = Hoja("PARTE_DIARIO")
            alto = PD_ALTO
            nombre = "PARTES_DIARIOS"
            subc = "POR_EQUIPO"
            que = "equipo"
    End Select

    ult = FilasArmadas(ws)
    If ult = 0 Then
        If Not Pregunta("La hoja " & ws.Name & " está vacía (todavía no se generó)." & vbCrLf & vbCrLf & _
                        "¿Generarla ahora con los datos del mes y exportarla?", nombre, True) Then Exit Function
        ArmarAhora cual
        ult = FilasArmadas(ws)
        If ult = 0 Then
            Aviso "No hay datos del mes para " & nombre & ".", nombre, 48
            Exit Function
        End If
    End If
    nPag = (ult + alto - 1) \ alto

    modo = IMP_NORMAL
    If cual <> "KARDEX" Then
        modo = ElegirImpresion(Replace(nombre, "_", " "))
        If modo = 0 Then Exit Function
    End If
    porUno = False
    If modo = IMP_NORMAL Then
        If gPorUno = 1 Then
            porUno = True
        ElseIf gPorUno = 0 Then
            porUno = Pregunta("Se arma un PDF con todo lo que está en la hoja " & ws.Name & " (" & nPag & " hojas)." & _
                              vbCrLf & vbCrLf & "¿Desea además un PDF por cada " & que & " (carpeta " & subc & ")?", nombre, False)
        End If
    End If
    Select Case modo
        Case IMP_DOS
            suf = "_A4_DOS_POR_CARA"
        Case IMP_COPIA
            suf = "_A4_ORIGINAL_Y_COPIA"
        Case Else
            suf = "_TODOS"
    End Select
    archivo = GuardarComo(RutaSugerida(nombre, nombre & "_" & MesArchivo() & suf & ".pdf"), _
                          "Guardar el PDF de " & Replace(nombre, "_", " ") & " como", "PDF")
    If archivo = "" Then
        Aviso "No se generó el PDF (se canceló «Guardar como»). Lo armado sigue en la hoja " & ws.Name & ".", nombre, 48
        Exit Function
    End If
    carpeta = CarpetaDe(archivo)
    Estado "Exportando " & Replace(nombre, "_", " ") & " a PDF..."

    area = "$B$1:$P$" & ult
    PaginaFormato ws, cual
    On Error Resume Next
    ws.PageSetup.PrintArea = area
    On Error GoTo 0
    If modo = IMP_NORMAL Then
        Set wsX = ws
    Else
        Set wsX = ArmarDoble(ws, alto, ult, (modo = IMP_COPIA))
    End If
    ok = ExportarPDF(wsX, archivo)
    If ok Then
        nPdf = 1
        If porUno Then
            sub2 = SubcarpetaPor(carpeta, nombre, subc)
            If CrearCarpeta(sub2) Then
                For r = 1 To ult Step alto
                    If Txt(ws.Cells(r, MK_CLAVE).Value) <> "" Then
                        Estado "PDF por " & que & ": " & Txt(ws.Cells(r, MK_CLAVE).Value)
                        On Error Resume Next
                        ws.PageSetup.PrintArea = "$B$" & r & ":$P$" & CLng(Num(ws.Cells(r, MK_FIN).Value))
                        On Error GoTo 0
                        If ExportarPDF(ws, sub2 & Sep() & IIf(cual = "PARTES", "PARTES", nombre) & "_" & _
                                       NombreSeguro(Txt(ws.Cells(r, MK_CLAVE).Value), 70) & ".pdf") Then nPdf = nPdf + 1
                    End If
                Next r
                On Error Resume Next
                ws.PageSetup.PrintArea = area
                On Error GoTo 0
            End If
        End If
        msg = Replace(nombre, "_", " ") & " de " & MesTexto() & ": " & nPdf & " PDF" & vbCrLf & archivo
        Select Case modo
            Case IMP_DOS
                msg = msg & vbCrLf & vbCrLf & "A4 horizontal, dos por cara (" & nPag & " formatos en " & (nPag + 1) \ 2 & _
                      " hojas). La hoja " & HOJA_A4 & " queda lista para imprimir directo (Ctrl+P)."
            Case IMP_COPIA
                msg = msg & vbCrLf & vbCrLf & "A4 horizontal, original y copia en cada cara (" & nPag & " hojas). " & _
                      "La hoja " & HOJA_A4 & " queda lista para imprimir directo (Ctrl+P)."
            Case Else
                If cual <> "KARDEX" Then msg = msg & vbCrLf & vbCrLf & "Papel A5, uno por hoja."
        End Select
        Mensaje D_MSGPDF, Replace(nombre, "_", " ") & ": " & nPdf & " PDF (" & Format(Now, "dd/mm/yyyy hh:mm") & ") en " & carpeta
        ExportarFormato = archivo
    Else
        msg = "No se pudo crear el PDF en:" & vbCrLf & archivo & vbCrLf & "(¿el archivo está abierto en otro programa?)"
    End If
    Estado ""
    If Not gSilencio Then
        ThisWorkbook.Activate
        ws.Activate
        ws.Range("A1").Select
    End If
    Aviso msg, Replace(nombre, "_", " ")
End Function


'========================================================================== BOTONES
Public Sub PdfKardex()
    Dim s As String
    s = ExportarFormato("KARDEX")
End Sub

Public Sub PdfBincard()
    Dim s As String
    s = ExportarFormato("BINCARD")
End Sub

Public Sub PdfPartes()
    Dim s As String
    s = ExportarFormato("PARTES")
End Sub

Public Sub PdfTodo()
    ' los tres formatos (lo que está en sus hojas; si una está vacía se arma) y, si se pide, los informes
    Dim carpeta As String, inf As Boolean, s As String, modo As Long
    modo = ElegirImpresion("BINCARD y PARTES DIARIOS")
    If modo = 0 Then Exit Sub
    If modo = IMP_NORMAL Then
        If Pregunta("¿Desea además un PDF por cada insumo y por cada equipo?", "PDF DE TODO", False) Then
            gPorUno = 1
        Else
            gPorUno = 2
        End If
    Else
        gPorUno = 2
    End If
    inf = Pregunta("¿Exportar también el INFORME MENSUAL de almacén (inventario, movimiento, ingresos – egresos y saldo valorizado) y de equipo mecánico" & _
                   " en Excel y PDF?", "PDF DE TODO", True)
    carpeta = ElegirCarpetaDesde("Elija la carpeta donde se guardarán todos los archivos (Guardar en...)", CarpetaSalida(""))
    If carpeta = "" Then
        gPorUno = 0
        Aviso "Se canceló la elección de la carpeta: no se exportó nada.", "PDF DE TODO", 48
        Exit Sub
    End If
    gImpresion = modo
    gCarpetaFija = carpeta
    gAcum = True
    gTexto = ""
    On Error GoTo Sale
    s = ExportarFormato("KARDEX")
    s = ExportarFormato("BINCARD")
    s = ExportarFormato("PARTES")
    If inf Then
        s = ExportarInformeAlmacen(False)
        s = ExportarInformeEquipos(False)
        ThisWorkbook.Activate
    End If
Sale:
    gAcum = False
    gCarpetaFija = ""
    gImpresion = 0
    gPorUno = 0
    If Err.Number <> 0 Then gTexto = gTexto & "Error: " & Err.Description
    Aviso "Carpeta: " & carpeta & vbCrLf & vbCrLf & gTexto, "PDF DE TODO"
    gTexto = ""
End Sub

'========================================================================== PRUEBAS
Public Sub PruebaA4Dos()
    gImpresion = IMP_DOS
End Sub

Public Sub PruebaA4Copia()
    gImpresion = IMP_COPIA
End Sub
