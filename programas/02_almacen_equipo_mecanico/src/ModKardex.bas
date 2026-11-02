Attribute VB_Name = "ModKardex"
Option Explicit

'==========================================================================================
'  KARDEX Y BINCARD: lectura del almacén y armado de las tarjetas (hojas + PDF)
'  Formatos iguales a los del informe mensual (IMP. KARDEX e IMP. BINCARD):
'  hojas FMT_KARDEX y FMT_BINCARD, área de impresión B:P, con logo y recuadros redondeados.
'==========================================================================================

' ---- geometría de los formatos (la misma de las hojas FMT_KARDEX y FMT_BINCARD)
Public Const KX_ALTO As Long = 41             ' filas de una hoja de kardex (1..41)
Public Const KX_L0 As Long = 10               ' primera fila de movimientos
Public Const KX_N As Long = 22                ' movimientos por hoja (filas 10..31)
Public Const KX_COLS As Long = 16             ' columnas A..P
Public Const BC_ALTO As Long = 37             ' filas de una tarjeta bincard (1..37)
Public Const BC_L0 As Long = 10               ' primera fila de movimientos
Public Const BC_N As Long = 18                ' movimientos por hoja (filas 10..27)
Public Const BC_COLS As Long = 16             ' columnas A..P

' ---- catálogo leído
Public anCat As Long
Public aCod() As String, aDes() As String, aUnd() As String, aGru() As String
Public aPU() As Double, aAnt() As Double
Public aPrim() As Long, aUlt() As Long, aCntMov() As Long
' ---- movimientos válidos (vOrden: orden cronológico)
Public anMov As Long, anMal As Long
Public vFec() As Double, vTipo() As String, vDoc() As String, vNDoc() As String, vOC() As String
Public vProc() As String, vPart() As String, vPed() As String, vObs() As String
Public vCant() As Double, vItem() As Long, vFila() As Long, vSig() As Long, vOrden() As Long
Public aCol As Collection


Public Sub LeerAlmacenCatalogo()
    ' (formato anterior) lee el CATALOGO y los MOVIMIENTOS válidos, ordenados por fecha (ingresos antes que salidas del mismo día)
    Dim ws As Worksheet, v As Variant, i As Long, n As Long, k As Long, last As Long, j As Long
    Dim cod As String, t As Double, tp As String, q As Double
    Dim llave() As Double

    Set aCol = New Collection
    anCat = 0
    anMov = 0
    anMal = 0
    Set ws = Hoja("CATALOGO")
    last = UltimaFila(ws, CC_COD, CT0, CTMAX)
    n = last - CT0 + 1
    If n < 1 Then n = 1
    ReDim aCod(1 To n)
    ReDim aDes(1 To n)
    ReDim aUnd(1 To n)
    ReDim aGru(1 To n)
    ReDim aPU(1 To n)
    ReDim aAnt(1 To n)
    ReDim aPrim(1 To n)
    ReDim aUlt(1 To n)
    ReDim aCntMov(1 To n)
    If last >= CT0 Then
        v = ws.Range(ws.Cells(CT0, 1), ws.Cells(last, CC_ULT)).Value
        For i = 1 To UBound(v, 1)
            cod = txt(v(i, CC_COD))
            If cod <> "" Then
                If Buscar(aCol, UCase(cod)) = 0 Then
                    anCat = anCat + 1
                    aCod(anCat) = cod
                    aDes(anCat) = Limpio(v(i, CC_DES))
                    aUnd(anCat) = txt(v(i, CC_UND))
                    aGru(anCat) = txt(v(i, CC_GRU))
                    aPU(anCat) = Num(v(i, CC_PU))
                    aAnt(anCat) = Num(v(i, CC_SANT))
                    aCol.Add anCat, UCase(cod)
                End If
            End If
        Next i
    End If

    Set ws = Hoja("MOVIMIENTOS")
    last = UltimaFilaVarias(ws, MC_FEC, MC_CANT, MV0, MVMAX)
    n = last - MV0 + 1
    If n < 1 Then n = 1
    ReDim vFec(1 To n)
    ReDim vTipo(1 To n)
    ReDim vDoc(1 To n)
    ReDim vNDoc(1 To n)
    ReDim vOC(1 To n)
    ReDim vProc(1 To n)
    ReDim vPart(1 To n)
    ReDim vPed(1 To n)
    ReDim vObs(1 To n)
    ReDim vCant(1 To n)
    ReDim vItem(1 To n)
    ReDim vFila(1 To n)
    ReDim vSig(1 To n)
    ReDim vOrden(1 To n)
    ReDim llave(1 To n)
    If last >= MV0 Then
        v = ws.Range(ws.Cells(MV0, 1), ws.Cells(last, MC_ULT)).Value
        For i = 1 To UBound(v, 1)
            If txt(v(i, MC_COD)) <> "" Or txt(v(i, MC_CANT)) <> "" Or txt(v(i, MC_FEC)) <> "" Then
                t = Fecha(v(i, MC_FEC))
                tp = UCase(txt(v(i, MC_TIPO)))
                k = Buscar(aCol, UCase(txt(v(i, MC_COD))))
                q = Num(v(i, MC_CANT))
                If t > 0 And k > 0 And q <> 0 And (tp = T_ING Or tp = T_SAL) Then
                    anMov = anMov + 1
                    vFec(anMov) = t
                    vTipo(anMov) = tp
                    vDoc(anMov) = Limpio(v(i, MC_DOC))
                    vNDoc(anMov) = Limpio(v(i, MC_NDOC))
                    vOC(anMov) = Limpio(v(i, MC_OC))
                    vProc(anMov) = Limpio(v(i, MC_PROC))
                    vPart(anMov) = txt(v(i, MC_PART))
                    vPed(anMov) = Limpio(v(i, MC_PED))
                    vObs(anMov) = Limpio(v(i, MC_OBS))
                    vCant(anMov) = q
                    vItem(anMov) = k
                    vFila(anMov) = MV0 + i - 1
                    If tp = T_SAL Then
                        llave(anMov) = (t * 10# + 1#) * 100000# + i
                    Else
                        llave(anMov) = (t * 10#) * 100000# + i
                    End If
                    vOrden(anMov) = anMov
                Else
                    anMal = anMal + 1
                End If
            End If
        Next i
    End If
    If anMov > 1 Then Ordenar llave, vOrden, anMov
    For j = 1 To anMov
        i = vOrden(j)
        k = vItem(i)
        vSig(i) = 0
        If aPrim(k) = 0 Then
            aPrim(k) = i
        Else
            vSig(aUlt(k)) = i
        End If
        aUlt(k) = i
        aCntMov(k) = aCntMov(k) + 1
    Next j
End Sub


Public Sub LeerAlmacen()
    ' insumos y movimientos del mes desde el INV. y el cuadro 4.10
    LeerCuadrosAlmacen
End Sub


Public Sub GenerarKardex()
    Dim n As Long, s As String
    s = ActualizarCuadros(True)
    n = ArmarTarjetas(True, True, Pregunta("Se arma un PDF con todos los kardex del mes." & vbCrLf & vbCrLf & _
                      "¿Desea además un PDF por cada insumo (carpeta KARDEX\POR_INSUMO)?", "KARDEX", False))
End Sub

Public Sub GenerarBincard()
    Dim n As Long, s As String
    s = ActualizarCuadros(True)
    n = ArmarTarjetas(False, True, Pregunta("Se arma un PDF con todas las tarjetas BINCARD del mes." & vbCrLf & vbCrLf & _
                      "¿Desea además un PDF por cada insumo (carpeta BINCARD\POR_INSUMO)?", "BINCARD", False))
End Sub


Private Function Juntar(ByVal a As String, ByVal b As String, ByVal sep As String) As String
    If a = "" Then
        Juntar = b
    ElseIf b = "" Then
        Juntar = a
    Else
        Juntar = a & sep & b
    End If
End Function

Private Sub PonerColumna(ByVal ws As Worksheet, ByVal F0 As Long, ByVal col As Long, ByRef lin() As Variant, _
                         ByVal desde As Long, ByVal n As Long, ByVal idx As Long)
    ' escribe una columna de la tabla (una sola columna: así no hay problemas con las celdas combinadas)
    Dim b() As Variant, i As Long
    ReDim b(1 To n, 1 To 1)
    For i = 1 To n
        b(i, 1) = lin(desde + i - 1, idx)
    Next i
    ws.Range(ws.Cells(F0, col), ws.Cells(F0 + n - 1, col)).Value = b
End Sub


Public Function ArmarTarjetas(ByVal esKardex As Boolean, ByVal conPdf As Boolean, ByVal porInsumo As Boolean) As Long
    ' arma una tarjeta (una o más hojas) por insumo con movimiento o saldo; devuelve el número de tarjetas
    Dim cm As Long, wsF As Worksheet, wsK As Worksheet, alto As Long, l0 As Long, nl As Long, ncol As Long, ucol As Long
    Dim k As Long, i As Long, p As Long, nPag As Long, nLin As Long, r0 As Long, totPag As Long, nTar As Long
    Dim lin() As Variant, li As Long, stock As Double, ingA As Double, salA As Double, m As Long, c As Long
    Dim desde() As Long, hasta() As Long, carpeta As String, nombre As String, archivo As String, ok As Boolean, nPdf As Long
    Dim area As String, msg As String, cols As Variant, conAnt As Boolean, hojaTxt As String, oc As String
    Dim ent As String, ger As String, sub1 As String, proy As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim copObj As Boolean, sub2 As String

    ArmarTarjetas = 0
    If esKardex Then
        nombre = "KARDEX"
    Else
        nombre = "BINCARD"
    End If
    Inicio cm, "Armando " & nombre & "..."
    LeerAlmacen
    If anCat = 0 Then
        Fin cm
        Aviso "El INV. está vacío: registre los ingresos del almacén en la hoja INV. 2026.", nombre, 48
        Exit Function
    End If
    If esKardex Then
        Set wsF = Hoja("FMT_KARDEX")
        Set wsK = Hoja("KARDEX")
        alto = KX_ALTO
        l0 = KX_L0
        nl = KX_N
        ' N°, FECHA, ORD. COMPRA, INGRESO, SALIDA, STOCK, ING. ACUM., SAL. ACUM., PEDIDO POR, FIRMA, OBSERVACIONES
        cols = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14)
        ucol = KX_COLS
    Else
        Set wsF = Hoja("FMT_BINCARD")
        Set wsK = Hoja("BINCARD")
        alto = BC_ALTO
        l0 = BC_L0
        nl = BC_N
        ' FECHA, CLASE, N°, ENTRADA, SALIDA, SALDOS
        cols = Array(3, 5, 7, 10, 12, 14)
        ucol = BC_COLS
    End If
    ncol = UBound(cols) - LBound(cols) + 1
    ent = txt(Hoja("DATOS").Range(D_ENTIDAD).Value)
    ger = txt(Hoja("DATOS").Range(D_GEREN).Value)
    sub1 = txt(Hoja("DATOS").Range(D_SUBG).Value)
    proy = Limpio(Hoja("DATOS").Range(D_PROY).Value)
    f1 = txt(Hoja("DATOS").Range(D_F1).Value)
    f2 = txt(Hoja("DATOS").Range(D_F2).Value)
    f3 = txt(Hoja("DATOS").Range(D_F3).Value)
    f4 = txt(Hoja("DATOS").Range(D_F4).Value)

    copObj = True
    On Error Resume Next
    copObj = Application.CopyObjectsWithCells
    Application.CopyObjectsWithCells = True
    wsK.ResetAllPageBreaks
    BorrarFormas wsK
    wsK.Cells.Delete
    On Error GoTo 0
    wsK.Cells.Clear
    For c = 1 To ucol
        wsK.Columns(c).ColumnWidth = wsF.Columns(c).ColumnWidth
    Next c
    CopiarPagina wsF, wsK

    ReDim desde(1 To anCat)
    ReDim hasta(1 To anCat)
    r0 = 1
    For k = 1 To anCat
        If aCntMov(k) > 0 Or Abs(aAnt(k)) > 0.0000001 Then
            nTar = nTar + 1
            If nTar Mod 10 = 0 Then Estado "Armando " & nombre & ": " & nTar & " ..."
            conAnt = (Abs(aAnt(k)) > 0.0000001)
            nLin = aCntMov(k)
            If conAnt Then nLin = nLin + 1
            If nLin < 1 Then nLin = 1
            nPag = (nLin + nl - 1) \ nl
            ReDim lin(1 To nPag * nl, 1 To ncol)
            stock = aAnt(k)
            ingA = 0
            salA = 0
            li = 0
            If conAnt Then
                li = 1
                If esKardex Then
                    lin(1, 2) = CDate(MesInicio())
                    lin(1, 3) = "SALDO ANTERIOR"
                    lin(1, 6) = Redondear(stock, 4)
                    lin(1, 11) = "Saldo al cierre del mes anterior"
                Else
                    lin(1, 1) = CDate(MesInicio())
                    lin(1, 2) = "SALDO ANTERIOR"
                    lin(1, 6) = Redondear(stock, 4)
                End If
            End If
            m = aPrim(k)
            Do While m > 0
                li = li + 1
                If vTipo(m) = T_ING Then
                    stock = stock + vCant(m)
                    ingA = ingA + vCant(m)
                Else
                    stock = stock - vCant(m)
                    salA = salA + vCant(m)
                End If
                If esKardex Then
                    oc = vOC(m)
                    If oc = "" Then oc = Juntar(vDoc(m), vNDoc(m), " ")
                    lin(li, 2) = CDate(vFec(m))
                    lin(li, 3) = oc
                    If vTipo(m) = T_ING Then
                        lin(li, 4) = vCant(m)
                    Else
                        lin(li, 5) = vCant(m)
                    End If
                    lin(li, 6) = Redondear(stock, 4)
                    lin(li, 7) = Redondear(ingA, 4)
                    lin(li, 8) = Redondear(salA, 4)
                    If vPed(m) <> "" Then
                        lin(li, 9) = vPed(m)
                        lin(li, 11) = Juntar(vProc(m), vObs(m), " – ")
                    Else
                        lin(li, 9) = vProc(m)
                        lin(li, 11) = vObs(m)
                    End If
                    If vOC(m) <> "" And vDoc(m) <> "" Then lin(li, 11) = Juntar(Juntar(vDoc(m), vNDoc(m), " "), _
                                                                         CStr(lin(li, 11)), " – ")
                Else
                    lin(li, 1) = CDate(vFec(m))
                    lin(li, 2) = vDoc(m)
                    If vNDoc(m) <> "" Then
                        lin(li, 3) = vNDoc(m)
                    Else
                        lin(li, 3) = vOC(m)
                    End If
                    If vTipo(m) = T_ING Then
                        lin(li, 4) = vCant(m)
                    Else
                        lin(li, 5) = vCant(m)
                    End If
                    lin(li, 6) = Redondear(stock, 4)
                End If
                m = vSig(m)
            Loop

            desde(k) = r0
            For p = 1 To nPag
                ' alto de filas antes de pegar: así los recuadros y el logo quedan del tamaño del formato
                For i = 1 To alto
                    wsK.Rows(r0 + i - 1).RowHeight = wsF.Rows(i).RowHeight
                Next i
                wsF.Range(wsF.Cells(1, 1), wsF.Cells(alto, ucol)).Copy Destination:=wsK.Cells(r0, 1)
                If nPag > 1 Then
                    hojaTxt = "   –   HOJA " & p & " DE " & nPag
                Else
                    hojaTxt = ""
                End If
                If esKardex Then
                    wsK.Cells(r0 + 1, 5).Value = ent
                    wsK.Cells(r0 + 2, 5).Value = ger
                    wsK.Cells(r0 + 3, 5).Value = sub1
                    wsK.Cells(r0 + 1, 9).Value = proy
                    wsK.Cells(r0 + 4, 11).Value = "FECHA: del " & FechaTxt(MesInicio()) & " al " & FechaTxt(FechaCorte()) & hojaTxt
                    wsK.Cells(r0 + 5, 5).Value = aDes(k)
                    wsK.Cells(r0 + 6, 14).Value = aUnd(k)
                    wsK.Cells(r0 + 6, 15).Value = aCod(k)
                    wsK.Cells(r0 + 31, 3).Value = f1
                    wsK.Cells(r0 + 31, 8).Value = f2
                    wsK.Cells(r0 + 31, 11).Value = f3
                    wsK.Cells(r0 + 31, 13).Value = f4
                    ' numeración correlativa entre hojas
                    For i = 1 To nl
                        wsK.Cells(r0 + l0 + i - 2, 3).Value = (p - 1) * nl + i
                    Next i
                Else
                    wsK.Cells(r0 + 2, 7).Value = aCod(k) & hojaTxt
                    wsK.Cells(r0 + 5, 4).Value = aDes(k) & vbLf & "UNIDAD: " & aUnd(k) & "     –     MES: " & MesTexto()
                End If
                For c = 1 To ncol
                    If esKardex And c = 1 Then
                        ' la columna N° ya está numerada
                    Else
                        PonerColumna wsK, r0 + l0 - 1, CLng(cols(LBound(cols) + c - 1)), lin, (p - 1) * nl + 1, nl, c
                    End If
                Next c
                ' textos largos: dos renglones con letra menor
                For i = 1 To nl
                    If (p - 1) * nl + i > li Then Exit For
                    If esKardex Then
                        AjustarTexto wsK.Cells(r0 + l0 + i - 2, 5), 15, 40, 9
                        AjustarTexto wsK.Cells(r0 + l0 + i - 2, 11), 44, 110, 9
                        AjustarTexto wsK.Cells(r0 + l0 + i - 2, 14), 36, 90, 9
                    Else
                        AjustarTexto wsK.Cells(r0 + l0 + i - 2, 5), 18, 60, 10
                        AjustarTexto wsK.Cells(r0 + l0 + i - 2, 7), 28, 90, 10
                    End If
                Next i
                If r0 > 1 Then
                    On Error Resume Next
                    wsK.HPageBreaks.Add Before:=wsK.Cells(r0, 1)
                    On Error GoTo 0
                End If
                r0 = r0 + alto
                totPag = totPag + 1
            Next p
            hasta(k) = r0 - 1
        End If
    Next k
    On Error Resume Next
    Application.CutCopyMode = False
    Application.CopyObjectsWithCells = copObj
    On Error GoTo 0

    If r0 > 1 Then
        area = "$B$1:$P$" & (r0 - 1)
        On Error Resume Next
        wsK.PageSetup.PrintArea = area
        On Error GoTo 0
    End If
    Fin cm
    ArmarTarjetas = nTar

    msg = "Se armaron " & nTar & " tarjetas de " & nombre & " (" & totPag & " hojas) del mes de " & MesTexto() & "."
    If conPdf And nTar > 0 Then
        Estado "Exportando " & nombre & " a PDF..."
        archivo = GuardarComo(RutaSugerida(nombre, nombre & "_" & MesArchivo() & "_TODOS.pdf"), _
                              "Guardar el PDF de " & nombre & " (todas las tarjetas) como", "PDF")
        If archivo = "" Then
            msg = msg & vbCrLf & vbCrLf & "No se generó el PDF (se canceló «Guardar como»). Las tarjetas están en la hoja " & nombre & "."
        Else
            carpeta = CarpetaDe(archivo)
            ok = ExportarPDF(wsK, archivo)
            If ok Then
                nPdf = 1
                If porInsumo Then
                    sub2 = SubcarpetaPor(carpeta, nombre, "POR_INSUMO")
                    If CrearCarpeta(sub2) Then
                        For k = 1 To anCat
                            If desde(k) > 0 Then
                                Estado "PDF por insumo: " & aCod(k)
                                On Error Resume Next
                                wsK.PageSetup.PrintArea = "$B$" & desde(k) & ":$P$" & hasta(k)
                                On Error GoTo 0
                                If ExportarPDF(wsK, sub2 & sep() & nombre & "_" & _
                                               NombreSeguro(aCod(k) & " " & aDes(k), 70) & ".pdf") Then nPdf = nPdf + 1
                            End If
                        Next k
                        On Error Resume Next
                        wsK.PageSetup.PrintArea = area
                        On Error GoTo 0
                    End If
                End If
                msg = msg & vbCrLf & vbCrLf & "PDF generados: " & nPdf & vbCrLf & archivo
                Mensaje D_MSGPDF, nombre & ": " & nTar & " tarjetas, " & nPdf & " PDF (" & Format(Now, "dd/mm/yyyy hh:mm") & _
                        ") en " & carpeta
            Else
                msg = msg & vbCrLf & vbCrLf & "No se pudo crear el PDF en:" & vbCrLf & archivo & vbCrLf & _
                      "(¿el archivo está abierto en otro programa?)"
            End If
        End If
        Estado ""
    End If
    If Not gSilencio Then
        wsK.Activate
        wsK.Range("A1").Select
    End If
    Aviso msg, nombre
End Function
