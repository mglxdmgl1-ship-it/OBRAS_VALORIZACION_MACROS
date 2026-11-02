Attribute VB_Name = "ModReporte"
Option Explicit

'==========================================================================================
'  INFORME MENSUAL CON EL FORMATO DEL EJEMPLO (Excel + PDF, con «Guardar como»)
'   ALMACÉN : INV. (inventario de ingresos), 4.10 MOVIMIENTO DIARIO, 4.11 INGRESOS – EGRESOS,
'             4.12 SALDO VALORIZADO          (plantillas FMT_INV, FMT_410, FMT_411, FMT_412)
'   EQUIPOS : 4.15 CONTROL DE EQUIPO MECÁNICO Y COMBUSTIBLE y DETALLE por equipo
'                                            (plantillas FMT_415 y FMT_DET)
'  Las plantillas son copias de las hojas del informe de ejemplo; el programa copia sus filas
'  modelo (rubro, ítem, totales) y escribe los datos del mes.
'==========================================================================================

Private Const XL_VISIBLE As Long = -1
Private Const XL_HIDDEN As Long = 0
Private Const GRIS_DOMINGO As Long = 14277081      ' RGB(217,217,217)
Private Const OFS As Long = 2000                    ' las filas modelo se copian 2000 filas más abajo en la misma hoja

' ---- datos del almacén preparados para los cuadros
Private rN As Long                                  ' número de rubros
Private rNom() As String
Private iRub() As Long, iNum() As String, iTipo() As String, iIng() As Double, iSal() As Double
Private dIng() As Double, dSal() As Double          ' (insumo, día)
Private nFuera As Long
' ---- historial de ingresos (meses cerrados)
Private hN As Long
Private hFec() As Double, hCod() As String, hDoc() As String, hNDoc() As String, hOC() As String, hProc() As String
Private hCant() As Double


'========================================================================== AUXILIARES
Private Function MesNombre(ByVal t As Double) As String
    ' "Octubre"
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

Private Function LetraDia(ByVal t As Double) As String
    LetraDia = Choose(Weekday(CDate(t), vbMonday), "L", "M", "M", "J", "V", "S", "D")
End Function

Private Function MesCorto() As String
    ' "OCT26"
    MesCorto = Choose(Month(CDate(MesInicio())), "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", _
                      "NOV", "DIC") & Right(CStr(Year(CDate(MesInicio()))), 2)
End Function

Private Function Dato(ByVal celda As String) As String
    Dato = Limpio(Hoja("DATOS").Range(celda).Value)
End Function

Private Function ConDos(ByVal s As String) As String
    ' ": texto" como en los encabezados del ejemplo
    ConDos = ": " & s
End Function

Private Function Doc(ByVal s As String) As String
    ' "O/C N° 04766" -> "O/C 04766"
    s = Replace(s, "O/C N° ", "O/C ")
    s = Replace(s, "O/S N° ", "O/S ")
    s = Replace(s, "O/C N°", "O/C ")
    Doc = Trim(s)
End Function

Private Sub CopiarFilas(ByVal wsT As Worksheet, ByVal a As Long, ByVal b As Long, ByVal ws As Worksheet, ByVal r As Long)
    ' copia las filas modelo a..b (guardadas en la misma hoja, OFS filas más abajo) a la fila r
    Dim i As Long
    ws.Rows((a + OFS) & ":" & (b + OFS)).Copy Destination:=ws.Rows(r)
    For i = 0 To b - a
        ws.Rows(r + i).RowHeight = wsT.Rows(a + i).RowHeight
    Next i
End Sub

Private Function CopiarPlantilla(ByVal nombre As String, ByRef wb As Workbook, ByVal nuevoNombre As String, _
                                 ByVal m0 As Long, ByVal m1 As Long) As Worksheet
    ' copia una hoja FMT_ (oculta) al libro del informe; si wb es Nothing crea el libro
    Dim wsT As Worksheet, ws As Worksheet, vis As Long
    Set wsT = Hoja(nombre)
    vis = wsT.Visible
    wsT.Visible = XL_VISIBLE
    If wb Is Nothing Then
        wsT.Copy
        Set wb = ActiveWorkbook
        Set ws = wb.Worksheets(1)
    Else
        wsT.Copy After:=wb.Worksheets(wb.Worksheets.Count)
        Set ws = wb.Worksheets(wb.Worksheets.Count)
    End If
    wsT.Visible = vis
    ws.Visible = XL_VISIBLE
    ' filas modelo (rubro, ítem, totales) guardadas abajo para copiarlas dentro de la misma hoja
    ws.Rows(m0 & ":" & m1).Copy Destination:=ws.Rows(m0 + OFS)
    On Error Resume Next
    ws.Name = nuevoNombre
    On Error GoTo 0
    Set CopiarPlantilla = ws
End Function

Private Sub BorrarSobrantes(ByVal ws As Worksheet, ByVal desde As Long)
    ' quita las filas modelo que quedaron debajo de lo escrito
    ws.Rows((OFS + 1) & ":" & (OFS + 80)).Delete
    ws.Rows(desde & ":" & (desde + 60)).Delete
End Sub

Private Sub Domingos(ByVal ws As Worksheet, ByVal col1 As Long, ByVal f0 As Long, ByVal f1 As Long)
    ' columnas de los domingos en gris (como en el ejemplo); col1 = columna del día 1
    Dim d As Long
    If f1 < f0 Then Exit Sub
    For d = 1 To DiasMes()
        If Weekday(DateSerial(Year(CDate(MesInicio())), Month(CDate(MesInicio())), d), vbMonday) = 7 Then
            ws.Range(ws.Cells(f0, col1 + d - 1), ws.Cells(f1, col1 + d - 1)).Interior.Color = GRIS_DOMINGO
        End If
    Next d
End Sub

Private Sub CabeceraDias(ByVal ws As Worksheet, ByVal fLetra As Long, ByVal fNum As Long, ByVal col1 As Long)
    ' letras (L M M J V S D) y números de los días; oculta los días que no tiene el mes
    Dim d As Long, t As Double
    For d = 1 To 31
        If d <= DiasMes() Then
            t = CDbl(DateSerial(Year(CDate(MesInicio())), Month(CDate(MesInicio())), d))
            ws.Cells(fLetra, col1 + d - 1).Value = LetraDia(t)
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
    ' =SUM(Hr:ALr)
    SumaFila = "=SUM(" & ws.Cells(r, c1).Address(False, False) & ":" & ws.Cells(r, c2).Address(False, False) & ")"
End Function

Private Function NombreArchivoInforme(ByVal tema As String) As String
    Dim n As String
    n = Dato(D_NINF)
    If n <> "" Then n = n & ". "
    NombreArchivoInforme = n & tema & "_" & Format(CDate(FechaCorte()), "dd") & " DE " & NombreMes(MesInicio()) & " " & _
                           Year(CDate(MesInicio())) & ".xlsx"
End Function

Private Function GuardarLibro(ByVal wb As Workbook, ByVal tema As String, ByRef rutaPdf As String, _
                              ByVal preguntarPdf As Boolean) As String
    ' «Guardar como» del libro del informe (y PDF); devuelve la ruta del Excel o ""
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


'========================================================================== HISTORIAL DE INGRESOS
Public Function HojaHistorial() As Worksheet
    ' hoja oculta con los ingresos de los meses cerrados (para el inventario INV.)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("HIST_INGRESOS")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "HIST_INGRESOS"
        ws.Range("A1:J1").Value = Array("FECHA", "CÓDIGO", "DESCRIPCIÓN", "UND.", "DOCUMENTO", "N° DOCUMENTO", "O/C – O/S", _
                                        "PROCEDENCIA", "CANTIDAD", "MES")
        ws.Range("A1:J1").Font.Bold = True
        ws.Visible = XL_HIDDEN
    End If
    Set HojaHistorial = ws
End Function

Public Sub GuardarHistorial()
    ' al CERRAR MES: pasa los ingresos del mes al historial
    Dim ws As Worksheet, r As Long, m As Long, j As Long
    LeerAlmacen
    If anMov = 0 Then Exit Sub
    Set ws = HojaHistorial()
    r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    If r < 2 Then r = 2
    For j = 1 To anMov
        m = vOrden(j)
        If vTipo(m) = T_ING Then
            ws.Cells(r, 1).Value = CDate(vFec(m))
            ws.Cells(r, 2).Value = aCod(vItem(m))
            ws.Cells(r, 3).Value = aDes(vItem(m))
            ws.Cells(r, 4).Value = aUnd(vItem(m))
            ws.Cells(r, 5).Value = vDoc(m)
            ws.Cells(r, 6).NumberFormat = "@"
            ws.Cells(r, 6).Value = vNDoc(m)
            ws.Cells(r, 7).Value = vOC(m)
            ws.Cells(r, 8).Value = vProc(m)
            ws.Cells(r, 9).Value = vCant(m)
            ws.Cells(r, 10).Value = MesArchivo()
            r = r + 1
        End If
    Next j
End Sub

Private Sub LeerHistorial()
    Dim ws As Worksheet, v As Variant, last As Long, i As Long
    hN = 0
    ReDim hFec(1 To 1): ReDim hCod(1 To 1): ReDim hDoc(1 To 1): ReDim hNDoc(1 To 1)
    ReDim hOC(1 To 1): ReDim hProc(1 To 1): ReDim hCant(1 To 1)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("HIST_INGRESOS")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    If last < 2 Then Exit Sub
    v = ws.Range("A2:J" & last).Value
    ReDim hFec(1 To UBound(v, 1)): ReDim hCod(1 To UBound(v, 1)): ReDim hDoc(1 To UBound(v, 1))
    ReDim hNDoc(1 To UBound(v, 1)): ReDim hOC(1 To UBound(v, 1)): ReDim hProc(1 To UBound(v, 1))
    ReDim hCant(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If txt(v(i, 2)) <> "" Then
            ' no repetir el mes que se está informando (si se reabrió)
            If txt(v(i, 10)) <> MesArchivo() Then
                hN = hN + 1
                hFec(hN) = Fecha(v(i, 1))
                hCod(hN) = txt(v(i, 2))
                hDoc(hN) = txt(v(i, 5))
                hNDoc(hN) = txt(v(i, 6))
                hOC(hN) = txt(v(i, 7))
                hProc(hN) = txt(v(i, 8))
                hCant(hN) = Num(v(i, 9))
            End If
        End If
    Next i
End Sub


'========================================================================== PREPARAR LOS DATOS DEL ALMACÉN
Private Sub PrepararAlmacen()
    Dim k As Long, g As Long, m As Long, d As Long, cnt() As Long, f As Variant, ws As Worksheet, last As Long
    Dim t0 As Double, t1 As Double, i As Long, fue As String
    LeerAlmacen
    LeerHistorial
    ReDim rNom(1 To anCat + 1)
    ReDim iRub(1 To anCat + 1)
    ReDim iNum(1 To anCat + 1)
    ReDim iTipo(1 To anCat + 1)
    ReDim iIng(1 To anCat + 1)
    ReDim iSal(1 To anCat + 1)
    ReDim dIng(1 To anCat + 1, 1 To 31)
    ReDim dSal(1 To anCat + 1, 1 To 31)
    ReDim cnt(1 To anCat + 1)
    rN = 0
    nFuera = 0
    For k = 1 To anCat
        g = 0
        For i = 1 To rN
            If Normal(rNom(i)) = Normal(IIf(aGru(k) = "", "OTROS", aGru(k))) Then g = i: Exit For
        Next i
        If g = 0 Then
            rN = rN + 1
            rNom(rN) = UCase(IIf(aGru(k) = "", "OTROS", aGru(k)))
            g = rN
        End If
        iRub(k) = g
        cnt(g) = cnt(g) + 1
        iNum(k) = g & "." & cnt(g)
    Next k
    ' movimientos por día
    t0 = MesInicio()
    t1 = MesFin()
    For k = 1 To anCat
        m = aPrim(k)
        Do While m > 0
            If vFec(m) >= t0 And vFec(m) <= t1 Then
                d = Day(CDate(vFec(m)))
                If vTipo(m) = T_ING Then
                    dIng(k, d) = dIng(k, d) + vCant(m)
                    iIng(k) = iIng(k) + vCant(m)
                Else
                    dSal(k, d) = dSal(k, d) + vCant(m)
                    iSal(k) = iSal(k) + vCant(m)
                End If
            Else
                nFuera = nFuera + 1
            End If
            If iTipo(k) = "" And vTipo(m) = T_ING Then
                If vOC(m) <> "" Then
                    iTipo(k) = Doc(vOC(m))
                Else
                    iTipo(k) = vDoc(m)
                End If
            End If
            m = vSig(m)
        Loop
    Next k
    ' tipo de documento desde el historial o la fuente del P.U.
    Set ws = Hoja("CATALOGO")
    For k = 1 To anCat
        If iTipo(k) = "" Then
            For i = 1 To hN
                If UCase(hCod(i)) = UCase(aCod(k)) Then
                    If hOC(i) <> "" Then iTipo(k) = Doc(hOC(i)) Else iTipo(k) = hDoc(i)
                    Exit For
                End If
            Next i
        End If
        If iTipo(k) = "" Then
            f = Application.Match(aCod(k), ws.Range(ws.Cells(CT0, CC_COD), ws.Cells(CTMAX, CC_COD)), 0)
            If Not IsError(f) Then
                fue = txt(ws.Cells(CT0 + CLng(f) - 1, CC_FUE).Value)
                If UCase(Left(fue, 3)) = "O/C" Or UCase(Left(fue, 3)) = "O/S" Then iTipo(k) = Doc(fue)
            End If
        End If
    Next k
End Sub


'========================================================================== INFORME DE ALMACÉN
Public Sub GenerarInformeAlmacen()
    Dim r As String
    r = InformeAlmacen(True)
End Sub

Public Function InformeAlmacen(ByVal preguntarPdf As Boolean) As String
    Dim cm As Long, wb As Workbook, ws As Worksheet, ruta As String, rutaPdf As String, msg As String
    InformeAlmacen = ""
    Inicio cm, "Armando el informe de almacén..."
    PrepararAlmacen
    If anCat = 0 Then
        Fin cm
        Aviso "El CATÁLOGO está vacío: registre los insumos del almacén.", "INFORME DE ALMACÉN", 48
        Exit Function
    End If
    On Error GoTo Falla
    Set wb = Nothing
    Set ws = CopiarPlantilla("FMT_INV", wb, "INV. " & Year(CDate(MesInicio())), 2, 4)
    Estado "Informe de almacén: inventario..."
    HojaInventario ws
    Estado "Informe de almacén: 4.10..."
    Set ws = CopiarPlantilla("FMT_410", wb, "4.10 MOV_" & MesCorto(), 13, 16)
    HojaMovimiento ws, Hoja("FMT_410"), False
    Estado "Informe de almacén: 4.11..."
    Set ws = CopiarPlantilla("FMT_411", wb, "4.11 INGRESO - EGRESO", 13, 16)
    HojaMovimiento ws, Hoja("FMT_411"), True
    Estado "Informe de almacén: 4.12..."
    Set ws = CopiarPlantilla("FMT_412", wb, "4.12 SALD VAL_" & MesCorto(), 13, 20)
    HojaSaldo ws, Hoja("FMT_412")
    Application.CutCopyMode = False
    wb.Worksheets(2).Activate
    wb.Worksheets(2).Range("A1").Select
    Fin cm
    ruta = GuardarLibro(wb, "ALMACEN DE OBRA", rutaPdf, preguntarPdf)
    msg = "Informe de almacén de " & MesTexto() & ": INV., 4.10, 4.11 y 4.12 (" & anCat & " insumos en " & rN & " rubros)."
    If nFuera > 0 Then msg = msg & vbCrLf & "Atención: " & nFuera & " movimiento(s) con fecha fuera del mes no entran en los cuadros diarios."
    If ruta = "" Then
        msg = msg & vbCrLf & vbCrLf & "El libro quedó abierto sin guardar (se canceló «Guardar como»)."
    Else
        msg = msg & vbCrLf & vbCrLf & "Excel: " & ruta
        If rutaPdf <> "" Then msg = msg & vbCrLf & "PDF: " & rutaPdf
        Mensaje D_MSGINF, "Informe de almacén (" & Format(Now, "dd/mm/yyyy hh:mm") & "): " & ruta
    End If
    InformeAlmacen = ruta
    Aviso msg, "INFORME DE ALMACÉN"
    Exit Function
Falla:
    Fin cm
    Aviso "No se pudo armar el informe de almacén: " & Err.Description, "INFORME DE ALMACÉN", 16
End Function

Private Sub HojaInventario(ByVal ws As Worksheet)
    ' INV.: un renglón por ingreso (meses anteriores + mes actual), agrupado por rubro, como tabla
    Dim wsT As Worksheet, r As Long, g As Long, k As Long, m As Long, i As Long, lo As Object, n As Long
    Set wsT = Hoja("FMT_INV")
    ws.Cells(1, 9).Value = "N° DOCUMENTO"
    r = 2
    For g = 1 To rN
        CopiarFilas wsT, 2, 2, ws, r
        ws.Cells(r, 1).Value = g
        ws.Cells(r, 2).Value = rNom(g)
        r = r + 1
        For k = 1 To anCat
            If iRub(k) = g Then
                ' meses anteriores
                For i = 1 To hN
                    If UCase(hCod(i)) = UCase(aCod(k)) Then
                        CopiarFilas wsT, 3, 3, ws, r
                        FilaInv ws, r, k, hFec(i), IIf(hOC(i) <> "", Doc(hOC(i)), hDoc(i)), hCant(i), _
                                Juntar2(hDoc(i), hNDoc(i)), hProc(i)
                        r = r + 1
                    End If
                Next i
                ' mes actual
                m = aPrim(k)
                Do While m > 0
                    If vTipo(m) = T_ING Then
                        CopiarFilas wsT, 3, 3, ws, r
                        FilaInv ws, r, k, vFec(m), IIf(vOC(m) <> "", Doc(vOC(m)), vDoc(m)), vCant(m), _
                                Juntar2(vDoc(m), vNDoc(m)), vProc(m)
                        r = r + 1
                    End If
                    m = vSig(m)
                Loop
            End If
        Next k
    Next g
    BorrarSobrantes ws, r
    n = r - 1
    On Error Resume Next
    Set lo = ws.ListObjects.Add(1, ws.Range("A1:J" & n), , 1)
    If Not lo Is Nothing Then
        lo.Name = "BD"
        lo.TableStyle = "TableStyleMedium6"
        lo.ShowTableStyleRowStripes = False
        lo.ShowTotals = True
        lo.ListColumns(1).TotalsCalculation = 0
        lo.TotalsRowRange.Cells(1, 1).Value = "Total"
        lo.ListColumns(7).TotalsCalculation = 1
        ws.PageSetup.PrintArea = lo.Range.Address
    Else
        ws.PageSetup.PrintArea = "$A$1:$J$" & n
    End If
    ws.PageSetup.PrintTitleRows = "$1:$1"
    On Error GoTo 0
End Sub

Private Function Juntar2(ByVal a As String, ByVal b As String) As String
    If b = "" Then
        Juntar2 = a
    ElseIf a = "" Then
        Juntar2 = b
    Else
        Juntar2 = a & " " & b
    End If
End Function

Private Sub FilaInv(ByVal ws As Worksheet, ByVal r As Long, ByVal k As Long, ByVal f As Double, ByVal tipo As String, _
                    ByVal cant As Double, ByVal ndoc As String, ByVal prov As String)
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = iNum(k)
    ws.Cells(r, 2).Value = aDes(k)
    ws.Cells(r, 3).Value = aUnd(k)
    ws.Cells(r, 4).Value = tipo
    If aPU(k) <> 0 Then ws.Cells(r, 5).Value = aPU(k)
    ws.Cells(r, 6).Value = cant
    ws.Cells(r, 7).Formula = "=" & ws.Cells(r, 5).Address(False, False) & "*" & ws.Cells(r, 6).Address(False, False)
    If f > 0 Then ws.Cells(r, 8).Value = CDate(f)
    ws.Cells(r, 9).Value = ndoc
    ws.Cells(r, 10).Value = prov
End Sub

Private Sub HojaMovimiento(ByVal ws As Worksheet, ByVal wsT As Worksheet, ByVal es411 As Boolean)
    ' 4.10 (es411 = False) y 4.11 (es411 = True): dos filas por insumo (ENTRADA / SALIDA) con los días del mes
    Dim r As Long, g As Long, k As Long, d As Long, v() As Variant, cT As Long, cS As Long, ent As String
    Dim f0 As Long, txt As String, cAno As Long, cP As Long
    ent = Dato(D_ENTIDAD) & vbLf & Dato(D_GEREN) & vbLf
    ws.Cells(1, 2).Value = ent
    If es411 Then
        ws.Cells(3, 2).Value = "4.11 CUADRO RESUMEN DE ALMACÉN (INGRESOS – EGRESOS) - " & NombreMes(MesInicio()) & " " & _
                               Year(CDate(MesInicio()))
        cP = 3
        cAno = 42
    Else
        ws.Cells(3, 2).Value = "4.10 CUADRO DE MOVIMIENTO DIARIO DE ALMACEN - " & NombreMes(MesInicio()) & " " & _
                               Year(CDate(MesInicio()))
        cP = 4
        cAno = 40
    End If
    ws.Cells(5, cP).Value = ConDos("“" & Dato(D_PROY) & "”.")
    ws.Cells(6, cP).Value = ConDos(Dato(D_MODAL))
    ws.Cells(7, cP).Value = ConDos(Dato(D_RESID))
    ws.Cells(8, cP).Value = ConDos(Dato(D_SUPER))
    ws.Cells(8, 2).Value = Dato(D_F4)
    ws.Cells(5, cAno).Value = Year(CDate(MesInicio()))
    ws.Cells(7, cAno).Value = MesNombre(MesInicio())
    ws.Cells(10, 7).Value = "SALDO" & vbLf & NombreMes(MesAnterior()) & " " & Year(CDate(MesAnterior()))
    txt = "SALDO " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    If es411 Then
        ws.Cells(10, 39).Value = "TOTAL"
        ws.Cells(10, 40).Value = "INGRESOS"
        ws.Cells(10, 41).Value = "SALIDAS"
        ws.Cells(10, 42).Value = txt
    Else
        ws.Cells(10, 39).Value = "TOTAL"
        ws.Cells(10, 40).Value = txt
    End If
    CabeceraDias ws, 10, 11, 8
    r = 13
    f0 = r
    For g = 1 To rN
        CopiarFilas wsT, 13, 14, ws, r
        ws.Cells(r, 2).Value = g
        ws.Cells(r, 3).Value = rNom(g)
        r = r + 2
        For k = 1 To anCat
            If iRub(k) = g Then
                CopiarFilas wsT, 15, 16, ws, r
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 2).Value = iNum(k)
                ws.Cells(r, 3).Value = aDes(k)
                ws.Cells(r, 4).Value = aUnd(k)
                ws.Cells(r, 5).Value = iTipo(k)
                ws.Cells(r, 6).Value = "ENTRADA"
                ws.Cells(r + 1, 6).Value = "SALIDA"
                ws.Cells(r, 7).Value = Redondear(aAnt(k), 4)
                ReDim v(1 To 2, 1 To 31)
                For d = 1 To 31
                    If dIng(k, d) <> 0 Then v(1, d) = Redondear(dIng(k, d), 4)
                    If dSal(k, d) <> 0 Then v(2, d) = Redondear(dSal(k, d), 4)
                Next d
                ws.Range(ws.Cells(r, 8), ws.Cells(r + 1, 38)).Value = v
                ws.Cells(r, 39).Formula = SumaFila(ws, r, 8, 38)
                ws.Cells(r + 1, 39).Formula = SumaFila(ws, r + 1, 8, 38)
                If es411 Then
                    ws.Cells(r, 40).Formula = "=" & ws.Cells(r, 39).Address(False, False)
                    ws.Cells(r, 41).Formula = "=" & ws.Cells(r + 1, 39).Address(False, False)
                    ws.Cells(r, 42).Formula = "=" & ws.Cells(r, 7).Address(False, False) & "+" & _
                                              ws.Cells(r, 39).Address(False, False) & "-" & ws.Cells(r + 1, 39).Address(False, False)
                Else
                    ws.Cells(r, 40).Formula = "=" & ws.Cells(r, 7).Address(False, False) & "+" & _
                                              ws.Cells(r, 39).Address(False, False) & "-" & ws.Cells(r + 1, 39).Address(False, False)
                End If
                r = r + 2
            End If
        Next k
    Next g
    BorrarSobrantes ws, r
    Domingos ws, 8, f0, r - 1
    On Error Resume Next
    If es411 Then
        ws.PageSetup.PrintArea = "$B$1:$AP$" & (r - 1)
    Else
        ws.PageSetup.PrintArea = "$A$1:$AN$" & (r - 1)
    End If
    ws.PageSetup.PrintTitleRows = "$10:$12"
    On Error GoTo 0
End Sub

Private Sub HojaSaldo(ByVal ws As Worksheet, ByVal wsT As Worksheet)
    ' 4.12: saldo anterior, ingresos, salidas y saldo del mes, valorizados con el P.U. del catálogo
    Dim r As Long, g As Long, k As Long, f0 As Long, f1 As Long, c As Long, a As String, c2 As Variant
    ws.Cells(1, 2).Value = Dato(D_ENTIDAD) & vbLf & Dato(D_GEREN) & vbLf
    ws.Cells(3, 2).Value = "4.12 SALDO VALORIZADO DE MATERIALES - " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    ws.Cells(5, 4).Value = ConDos("“" & Dato(D_PROY) & "”.")
    ws.Cells(7, 4).Value = ConDos(Dato(D_MODAL))
    ws.Cells(8, 4).Value = ConDos(Dato(D_RESID))
    ws.Cells(9, 4).Value = ConDos(Dato(D_SUPER))
    ws.Cells(9, 2).Value = Dato(D_F4)
    ws.Cells(7, 14).Value = Year(CDate(MesInicio()))
    ws.Cells(9, 14).Value = MesNombre(MesInicio())
    ws.Cells(11, 7).Value = "SALDO " & NombreMes(MesAnterior()) & " " & Year(CDate(MesAnterior()))
    ws.Cells(11, 9).Value = "INGRESOS " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    ws.Cells(11, 11).Value = "SALIDAS " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    ws.Cells(11, 13).Value = "SALDO " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
    r = 13
    f0 = r
    For g = 1 To rN
        CopiarFilas wsT, 13, 13, ws, r
        ws.Cells(r, 2).Value = g
        ws.Cells(r, 3).Value = rNom(g)
        r = r + 1
        For k = 1 To anCat
            If iRub(k) = g Then
                CopiarFilas wsT, 14, 14, ws, r
                ws.Cells(r, 2).NumberFormat = "@"
                ws.Cells(r, 2).Value = iNum(k)
                ws.Cells(r, 3).Value = aDes(k)
                ws.Cells(r, 4).Value = aUnd(k)
                ws.Cells(r, 5).Value = iTipo(k)
                ws.Cells(r, 6).Value = aPU(k)
                ws.Cells(r, 7).Value = Redondear(aAnt(k), 4)
                ws.Cells(r, 9).Value = Redondear(iIng(k), 4)
                ws.Cells(r, 11).Value = Redondear(iSal(k), 4)
                ws.Cells(r, 13).Formula = "=G" & r & "+I" & r & "-K" & r
                ws.Cells(r, 8).Formula = "=ROUND($F" & r & "*G" & r & ",2)"
                ws.Cells(r, 10).Formula = "=ROUND($F" & r & "*I" & r & ",2)"
                ws.Cells(r, 12).Formula = "=ROUND($F" & r & "*K" & r & ",2)"
                ws.Cells(r, 14).Formula = "=ROUND($F" & r & "*M" & r & ",2)"
                r = r + 1
            End If
        Next k
    Next g
    f1 = r - 1
    ' total y resumen
    CopiarFilas wsT, 15, 20, ws, r
    ws.Cells(r, 2).Value = "Total"
    For Each c2 In Array(8, 10, 12, 14)
        ws.Cells(r, CLng(c2)).Formula = "=SUM(" & ws.Cells(f0, CLng(c2)).Address(False, False) & ":" & _
                                        ws.Cells(f1, CLng(c2)).Address(False, False) & ")"
    Next c2
    For c = 0 To 3
        ws.Cells(r + 2 + c, 12).Value = ws.Cells(11, 7 + 2 * c).Value
        ws.Cells(r + 2 + c, 14).Formula = "=" & ws.Cells(r, 8 + 2 * c).Address(False, False)
    Next c
    BorrarSobrantes ws, r + 6
    On Error Resume Next
    ws.PageSetup.PrintArea = "$B$1:$N$" & (r + 5)
    ws.PageSetup.PrintTitleRows = "$11:$12"
    On Error GoTo 0
End Sub


'========================================================================== INFORME DE EQUIPO MECÁNICO (4.15)
Public Sub GenerarInformeEquipos()
    Dim r As String
    r = InformeEquipos(True)
End Sub

Private Function UnidadCorta(ByVal u As String) As String
    Dim s As String
    s = Normal(u)
    s = Replace(s, " ", "")
    If s = "HM" Or s = "H.M." Or s = "H.M" Or s = "HORAMAQUINA" Then
        UnidadCorta = "H.M."
    ElseIf s = "HE" Or s = "H.E." Or s = "H.E" Or s = "HORAEQUIPO" Then
        UnidadCorta = "H.E."
    ElseIf Left(s, 3) = "DIA" Or s = "D" Then
        UnidadCorta = "DIA"
    ElseIf Left(s, 3) = "VIA" Or Left(s, 3) = "VJE" Then
        UnidadCorta = "VJE"
    ElseIf s = "" Then
        UnidadCorta = "H.M."
    Else
        UnidadCorta = UCase(Trim(u))
    End If
End Function

Public Function InformeEquipos(ByVal preguntarPdf As Boolean) As String
    Dim cm As Long, wb As Workbook, ws As Worksheet, wsT As Worksheet, ruta As String, rutaPdf As String, msg As String
    Dim hor() As Double, com() As Double, tieneParte() As Boolean, j As Long, i As Long, k As Long, d As Long
    Dim sec As Long, unid() As String, bloques As Collection, u As Variant, r As Long, n As Long, f0 As Long, rB As Long
    Dim v() As Variant, primeros As String, ultimos As String, rTot As Long, ent As String, t As Double
    Dim combCod() As String, nComb As Long, combDia() As Double, cK As Long, x As Long, secNom As String, hayEq As Boolean
    Dim fRows As String, nDet As Long

    InformeEquipos = ""
    Inicio cm, "Armando el informe de equipo mecánico..."
    LeerPartes
    If enEq = 0 Then
        Fin cm
        Aviso "No hay equipos registrados en EQUIPOS.", "INFORME DE EQUIPO MECÁNICO", 48
        Exit Function
    End If
    On Error GoTo Falla
    ReDim hor(1 To enEq, 1 To 31)
    ReDim com(1 To enEq, 1 To 31)
    ReDim tieneParte(1 To enEq)
    ReDim unid(1 To enEq)
    For j = 1 To pnPar
        i = pR(j)
        k = pEq(j)
        t = pFec(j)
        If t >= MesInicio() And t <= MesFin() Then
            d = Day(CDate(t))
            hor(k, d) = hor(k, d) + HorasFila(i)
            com(k, d) = com(k, d) + Num(pV(i, PC_COMB))
            tieneParte(k) = True
        End If
    Next j
    For k = 1 To enEq
        unid(k) = UnidadCorta(eUnd(k))
    Next k

    Set wb = Nothing
    Set ws = CopiarPlantilla("FMT_415", wb, MesCorto(), 14, 25)
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
    ' días del mes (columnas J..AN) para todo el libro
    For d = 1 To 31
        ws.Columns(9 + d).Hidden = (d > DiasMes())
    Next d
    r = 14
    For sec = 1 To 2
        If sec = 1 Then
            secNom = "MAQUINARIAS Y EQUIPOS ALQUILADOS POR SERVICIO (O/S)"
        Else
            secNom = "MAQUINARIAS Y EQUIPOS PROPIOS O PRESTADOS"
        End If
        ' unidades de esta sección, en el orden de EQUIPOS
        Set bloques = New Collection
        hayEq = False
        For k = 1 To enEq
            If (sec = 1) = (eOS(k) <> "") Then
                hayEq = True
                On Error Resume Next
                bloques.Add unid(k), unid(k)
                On Error GoTo Falla
            End If
        Next k
        If hayEq Then
            CopiarFilas wsT, 14, 15, ws, r
            ws.Cells(r, 2).Value = secNom
            r = r + 2
            For Each u In bloques
                rB = r
                CopiarFilas wsT, 16, 18, ws, r
                ws.Cells(r, 6).Value = "COMBUS./ " & u
                If u = "DIA" Then
                    ws.Cells(r, 8).Value = "TOTAL DIAS"
                    ws.Cells(r, 9).Value = "DIAS ACUM."
                ElseIf u = "VJE" Then
                    ws.Cells(r, 8).Value = "TOTAL VIAJES"
                    ws.Cells(r, 9).Value = "VIAJES ACUM."
                Else
                    ws.Cells(r, 8).Value = "TOTAL HRS"
                    ws.Cells(r, 9).Value = "HORAS ACUM."
                End If
                ws.Cells(r, 10).Value = "MES DE " & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
                ws.Cells(r, 41).Value = "TOTAL"
                ws.Cells(r + 1, 41).Value = "COMB."
                ws.Cells(r + 1, 42).Value = IIf(u = "DIA", "DIAS", u)
                CabeceraDias ws, r + 1, r + 2, 10
                r = r + 3
                f0 = r
                n = 0
                For k = 1 To enEq
                    If (sec = 1) = (eOS(k) <> "") And unid(k) = u Then
                        n = n + 1
                        CopiarFilas wsT, 19, 20, ws, r
                        ws.Cells(r, 2).Value = n
                        ws.Cells(r, 3).Value = eDes(k) & IIf(eSerie(k) <> "", " – " & eSerie(k), "")
                        ws.Cells(r, 6).Value = "GAL"
                        ws.Cells(r + 1, 6).Value = IIf(u = "DIA", "DIA", u)
                        ReDim v(1 To 2, 1 To 31)
                        For d = 1 To 31
                            If com(k, d) <> 0 Then v(1, d) = Redondear(com(k, d), 2)
                            If hor(k, d) <> 0 Then v(2, d) = Redondear(hor(k, d), 2)
                        Next d
                        ws.Range(ws.Cells(r, 10), ws.Cells(r + 1, 40)).Value = v
                        ws.Cells(r, 41).Formula = SumaFila(ws, r, 10, 40)
                        ws.Cells(r + 1, 42).Formula = SumaFila(ws, r + 1, 10, 40)
                        ws.Cells(r, 9).Formula = "=" & Trim(Str(Num(Hoja("EQUIPOS").Cells(EQ0 + EquipoFila(k), EC_ANT).Value))) & _
                                                 "+" & ws.Cells(r + 1, 42).Address(False, False)
                        ws.Cells(r, 43).Value = eOper(k)
                        r = r + 2
                    End If
                Next k
                CopiarFilas wsT, 21, 22, ws, r
                ws.Cells(r, 31).Value = "TOTAL: COMBUSTIBLE/ " & IIf(u = "DIA", "DIAS", u)
                ws.Cells(r, 41).Formula = "=SUM(" & ws.Cells(f0, 41).Address(False, False) & ":" & _
                                          ws.Cells(r - 1, 41).Address(False, False) & ")"
                ws.Cells(r, 42).Formula = "=SUM(" & ws.Cells(f0, 42).Address(False, False) & ":" & _
                                          ws.Cells(r - 1, 42).Address(False, False) & ")"
                Domingos ws, 10, f0, r - 1
                r = r + 2
            Next u
        End If
    Next sec
    ' resumen de combustible por tipo (filas del día)
    ReDim combCod(1 To enEq)
    ReDim combDia(1 To enEq, 1 To 31)
    nComb = 0
    For k = 1 To enEq
        If eComb(k) <> "" Then
            cK = 0
            For x = 1 To nComb
                If UCase(combCod(x)) = UCase(eComb(k)) Then cK = x
            Next x
            If cK = 0 Then
                nComb = nComb + 1
                combCod(nComb) = eComb(k)
                cK = nComb
            End If
            For d = 1 To 31
                combDia(cK, d) = combDia(cK, d) + com(k, d)
            Next d
        End If
    Next k
    If nComb > 0 Then
        CopiarFilas wsT, 23, 24, ws, r
        CabeceraDias ws, r, r + 1, 10
        ws.Cells(r, 41).Value = "COMB."
        r = r + 2
        For x = 1 To nComb
            CopiarFilas wsT, 25, 25, ws, r
            ws.Cells(r, 2).Value = DescCatalogoR(combCod(x))
            ReDim v(1 To 1, 1 To 31)
            For d = 1 To 31
                If combDia(x, d) <> 0 Then v(1, d) = Redondear(combDia(x, d), 2)
            Next d
            ws.Range(ws.Cells(r, 10), ws.Cells(r, 40)).Value = v
            ws.Cells(r, 41).Formula = SumaFila(ws, r, 10, 40)
            r = r + 1
        Next x
    End If
    BorrarSobrantes ws, r
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$AP$" & (r - 1)
    ws.PageSetup.PrintTitleRows = "$1:$12"
    On Error GoTo Falla
    ' detalle diario por equipo
    For k = 1 To enEq
        If tieneParte(k) Then
            HojaDetalle wb, k
            nDet = nDet + 1
        End If
    Next k
    Application.CutCopyMode = False
    wb.Worksheets(1).Activate
    wb.Worksheets(1).Range("A1").Select
    Fin cm
    ruta = GuardarLibro(wb, "CONTROL DE EQUIPO MECANICO", rutaPdf, preguntarPdf)
    msg = "Informe 4.15 de " & MesTexto() & ": control de equipo mecánico y combustible (" & enEq & " equipos) y " & nDet & _
          " hoja(s) de detalle diario."
    If ruta = "" Then
        msg = msg & vbCrLf & vbCrLf & "El libro quedó abierto sin guardar (se canceló «Guardar como»)."
    Else
        msg = msg & vbCrLf & vbCrLf & "Excel: " & ruta
        If rutaPdf <> "" Then msg = msg & vbCrLf & "PDF: " & rutaPdf
        Mensaje D_MSGINF, "Informe de equipo mecánico (" & Format(Now, "dd/mm/yyyy hh:mm") & "): " & ruta
    End If
    InformeEquipos = ruta
    Aviso msg, "INFORME DE EQUIPO MECÁNICO"
    Exit Function
Falla:
    Fin cm
    Aviso "No se pudo armar el informe de equipo mecánico: " & Err.Description, "INFORME DE EQUIPO MECÁNICO", 16
End Function

Private Function EquipoFila(ByVal k As Long) As Long
    ' fila (desde 0) del equipo k en la hoja EQUIPOS
    Dim ws As Worksheet, f As Variant
    Set ws = Hoja("EQUIPOS")
    EquipoFila = 0
    f = Application.Match(eCod(k), ws.Range(ws.Cells(EQ0, EC_COD), ws.Cells(EQMAX, EC_COD)), 0)
    If Not IsError(f) Then EquipoFila = CLng(f) - 1
End Function

Private Function DescCatalogoR(ByVal cod As String) As String
    Dim ws As Worksheet, f As Variant
    Set ws = Hoja("CATALOGO")
    DescCatalogoR = cod
    f = Application.Match(cod, ws.Range(ws.Cells(CT0, CC_COD), ws.Cells(CTMAX, CC_COD)), 0)
    If Not IsError(f) Then DescCatalogoR = UCase(txt(ws.Cells(CT0 + CLng(f) - 1, CC_DES).Value))
End Function

Private Sub HojaDetalle(ByVal wb As Workbook, ByVal k As Long)
    ' detalle diario del equipo: horarios de mañana y tarde, horas y observaciones (formato «Detalle»)
    Dim ws As Worksheet, wsT As Worksheet, r As Long, j As Long, i As Long, t As Double, d As Long, f0 As Long
    Dim obs As String, h As Double, a1 As Variant, b1 As Variant, a2 As Variant, b2 As Variant, tr As String
    Set wsT = Hoja("FMT_DET")
    Set ws = CopiarPlantilla("FMT_DET", wb, Left("Detalle " & NombreSeguro(eCod(k) & " " & eDes(k), 40), 31), 4, 5)
    ws.Cells(1, 1).Value = "DETALLE DIARIO DE HORAS – " & eCod(k) & " " & eDes(k) & IIf(eSerie(k) <> "", " – " & eSerie(k), "") & _
                           "   (" & NombreMes(MesInicio()) & " " & Year(CDate(MesInicio())) & ")"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    r = 4
    f0 = r
    j = 1
    Do While j <= pnPar
        If pEq(pOrden(j)) = k Then
            t = pFec(pOrden(j))
            ' todas las filas de ese día
            a1 = Empty: b1 = Empty: a2 = Empty: b2 = Empty: h = 0: obs = "": tr = ""
            Do While j <= pnPar
                If pEq(pOrden(j)) <> k Or pFec(pOrden(j)) <> t Then Exit Do
                i = pR(pOrden(j))
                If IsEmpty(a1) And EsNum(pV(i, PC_MDE)) Then a1 = pV(i, PC_MDE): b1 = pV(i, PC_MA)
                If IsEmpty(a2) And EsNum(pV(i, PC_TDE)) Then a2 = pV(i, PC_TDE): b2 = pV(i, PC_TA)
                h = h + HorasFila(i)
                If txt(pV(i, PC_PART)) <> "" Then
                    If InStr(tr, txt(pV(i, PC_PART))) = 0 Then tr = tr & IIf(tr = "", "", ", ") & txt(pV(i, PC_PART))
                End If
                If Limpio(pV(i, PC_OBS)) <> "" Then
                    If InStr(obs, Limpio(pV(i, PC_OBS))) = 0 Then obs = obs & IIf(obs = "", "", "; ") & Limpio(pV(i, PC_OBS))
                End If
                j = j + 1
            Loop
            CopiarFilas wsT, 4, 4, ws, r
            ws.Cells(r, 1).Value = Day(CDate(t))
            If Not IsEmpty(a1) Then
                ws.Cells(r, 2).Value = Num(a1) - Int(Num(a1))
                ws.Cells(r, 3).Value = Num(b1) - Int(Num(b1))
            End If
            If Not IsEmpty(a2) Then
                ws.Cells(r, 5).Value = Num(a2) - Int(Num(a2))
                ws.Cells(r, 6).Value = Num(b2) - Int(Num(b2))
            End If
            ws.Cells(r, 4).Formula = "=IF(OR(B" & r & "="""",C" & r & "=""""),0,C" & r & "-B" & r & ")"
            ws.Cells(r, 7).Formula = "=IF(OR(E" & r & "="""",F" & r & "=""""),0,F" & r & "-E" & r & ")"
            ws.Cells(r, 10).Value = 0
            If IsEmpty(a1) And IsEmpty(a2) Then
                ' sin horario: horas a valorizar (reales o A.C.U.)
                ws.Cells(r, 11).Value = h / 24
                obs = Trim("Horas a valorizar sin horario (reales o A.C.U. x metrado). " & obs)
            Else
                ws.Cells(r, 11).Formula = "=SUM(D" & r & ",G" & r & ",J" & r & ")"
            End If
            ws.Cells(r, 12).Formula = "=K" & r & "*24"
            If tr <> "" Then obs = Trim("Partidas: " & tr & ". " & obs)
            ws.Cells(r, 13).Value = obs
            r = r + 1
        Else
            j = j + 1
        End If
    Loop
    CopiarFilas wsT, 5, 5, ws, r
    ws.Cells(r, 11).Formula = "=SUM(K" & f0 & ":K" & (r - 1) & ")"
    ws.Cells(r, 12).Formula = "=K" & r & "*24"
    BorrarSobrantes ws, r + 1
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$M$" & r
    ws.PageSetup.Orientation = 2
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    On Error GoTo 0
End Sub
