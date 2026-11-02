Attribute VB_Name = "ModBase"
Option Explicit

'==========================================================================================
'  PROGRAMA DE ALMACÉN Y EQUIPO MECÁNICO - CONSTANTES Y UTILIDADES COMUNES
'==========================================================================================

' ---- filas de datos de cada registro
Public Const CT0 As Long = 8                  ' CATALOGO
Public Const CTMAX As Long = 1007
Public Const MV0 As Long = 8                  ' MOVIMIENTOS
Public Const MVMAX As Long = 5007
Public Const EQ0 As Long = 8                  ' EQUIPOS
Public Const EQMAX As Long = 207
Public Const PT0 As Long = 8                  ' PARTES
Public Const PTMAX As Long = 3007

' ---- columnas del CATALOGO
Public Const CC_COD As Long = 2
Public Const CC_DES As Long = 3
Public Const CC_UND As Long = 4
Public Const CC_GRU As Long = 5
Public Const CC_PU As Long = 6
Public Const CC_FUE As Long = 7
Public Const CC_SANT As Long = 8
Public Const CC_SALDO As Long = 11
Public Const CC_ACU As Long = 15
Public Const CC_FAC As Long = 16
Public Const CC_ULT As Long = 17
' ---- columnas de MOVIMIENTOS
Public Const MC_FEC As Long = 2
Public Const MC_TIPO As Long = 3
Public Const MC_DOC As Long = 4
Public Const MC_NDOC As Long = 5
Public Const MC_OC As Long = 6
Public Const MC_PROC As Long = 7
Public Const MC_PART As Long = 8
Public Const MC_COD As Long = 9
Public Const MC_CANT As Long = 12
Public Const MC_PED As Long = 13
Public Const MC_OBS As Long = 14
Public Const MC_ORIG As Long = 15
Public Const MC_SALDO As Long = 17
Public Const MC_ULT As Long = 17
' ---- columnas de EQUIPOS
Public Const EC_COD As Long = 2
Public Const EC_DES As Long = 3
Public Const EC_MARCA As Long = 4
Public Const EC_MODELO As Long = 5
Public Const EC_SERIE As Long = 6
Public Const EC_ANIO As Long = 7
Public Const EC_PROV As Long = 8
Public Const EC_OS As Long = 9
Public Const EC_UND As Long = 10
Public Const EC_CANT As Long = 11
Public Const EC_CU As Long = 12
Public Const EC_COND As Long = 14
Public Const EC_COMB As Long = 15
Public Const EC_OPER As Long = 16
Public Const EC_ANT As Long = 17
Public Const EC_SALDO As Long = 20
Public Const EC_ACU As Long = 24
Public Const EC_ULT As Long = 25
' ---- columnas de PARTES
Public Const PC_FEC As Long = 2
Public Const PC_EQ As Long = 3
Public Const PC_NPARTE As Long = 5
Public Const PC_OPER As Long = 6
Public Const PC_MDE As Long = 7
Public Const PC_MA As Long = 8
Public Const PC_TDE As Long = 9
Public Const PC_TA As Long = 10
Public Const PC_HINI As Long = 11
Public Const PC_HFIN As Long = 12
Public Const PC_PART As Long = 13
Public Const PC_TRAB As Long = 14
Public Const PC_MET As Long = 15
Public Const PC_UND As Long = 16
Public Const PC_REND As Long = 17
Public Const PC_HACU As Long = 18
Public Const PC_HREAL As Long = 19
Public Const PC_HVAL As Long = 20
Public Const PC_COMB As Long = 21
Public Const PC_OBS As Long = 22
Public Const PC_ORIG As Long = 23
Public Const PC_CTRL As Long = 25
Public Const PC_ULT As Long = 25

' ---- celdas de la hoja DATOS
Public Const D_ENTIDAD As String = "C5"
Public Const D_GEREN As String = "C6"
Public Const D_SUBG As String = "C7"
Public Const D_PROY As String = "C8"
Public Const D_CUI As String = "C9"
Public Const D_META As String = "C10"
Public Const D_UBIC As String = "C11"
Public Const D_MODAL As String = "C12"
Public Const D_MES As String = "C14"
Public Const D_CORTE As String = "C15"
Public Const D_NINF As String = "C16"
Public Const D_F1 As String = "C18"
Public Const D_F2 As String = "C19"
Public Const D_F3 As String = "C20"
Public Const D_F4 As String = "C21"
Public Const D_CARPETA As String = "C23"
Public Const D_CTRL As String = "C24"
Public Const D_INF410 As String = "C25"
Public Const D_INF415 As String = "C26"
Public Const D_MSG As String = "C29"
Public Const D_MSGPDF As String = "C30"
Public Const D_MSGINF As String = "C31"
Public Const D_RESID As String = "C40"           ' nombre del residente (informes 4.10 - 4.15)
Public Const D_SUPER As String = "C41"           ' nombre del supervisor / inspector

' ---- textos fijos
Public Const T_ING As String = "INGRESO"
Public Const T_SAL As String = "SALIDA"
Public Const OR_ACU As String = "A.C.U."
Public Const OR_ING As String = "INGRESOS DEL CONTROL"

' ---- constantes de Excel
Public Const XL_MANUAL As Long = -4135
Public Const XL_CENTER As Long = -4108

Public gSilencio As Boolean                   ' verdadero en las pruebas automáticas (sin ventanas)
Public gSinPDF As Boolean                     ' verdadero en las pruebas: no exporta PDF
Public gAcum As Boolean                       ' junta los avisos en gTexto (GENERAR TODO)
Public gSiATodo As Boolean                    ' pruebas: responde Sí a todas las preguntas
Public gTexto As String
Public gCarpetaFija As String                 ' GENERAR TODO: carpeta elegida una sola vez para todos los archivos


'========================================================================== LECTURA SEGURA
Public Function ParseNum(ByVal s As String, ByRef ok As Boolean) As Double
    Dim t As String
    ok = False
    ParseNum = 0
    t = Replace(Replace(Replace(s, ChrW(160), ""), " ", ""), vbTab, "")
    If t = "" Then Exit Function
    If InStr(t, ",") > 0 And InStr(t, ".") > 0 Then
        If InStrRev(t, ",") > InStrRev(t, ".") Then
            t = Replace(t, ".", "")
            t = Replace(t, ",", ".")
        Else
            t = Replace(t, ",", "")
        End If
    ElseIf InStr(t, ",") > 0 Then
        If Len(t) - Len(Replace(t, ",", "")) = 1 Then
            t = Replace(t, ",", ".")
        Else
            t = Replace(t, ",", "")
        End If
    End If
    If EsNumeroPunto(t) Then
        ok = True
        ParseNum = Val(t)
    End If
End Function

Public Function EsNumeroPunto(ByVal t As String) As Boolean
    Dim i As Long, c As String, pts As Long, dig As Long
    EsNumeroPunto = False
    If t = "" Then Exit Function
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If c >= "0" And c <= "9" Then
            dig = dig + 1
        ElseIf c = "." Then
            pts = pts + 1
        ElseIf (c = "-" Or c = "+") And i = 1 Then
            ' signo
        Else
            Exit Function
        End If
    Next i
    EsNumeroPunto = (dig > 0 And pts <= 1)
End Function

Public Function Num(ByVal v As Variant) As Double
    Dim ok As Boolean, x As Double
    Num = 0
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        x = ParseNum(CStr(v), ok)
        If ok Then
            Num = x
        ElseIf IsDate(v) Then
            Num = CDbl(CDate(v))
        End If
    ElseIf VarType(v) = vbBoolean Then
        Num = 0
    ElseIf IsNumeric(v) Then
        Num = CDbl(v)
    ElseIf IsDate(v) Then
        Num = CDbl(CDate(v))
    End If
End Function

Public Function EsNum(ByVal v As Variant) As Boolean
    Dim ok As Boolean, x As Double
    EsNum = False
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        x = ParseNum(CStr(v), ok)
        EsNum = ok
    ElseIf VarType(v) = vbBoolean Then
        EsNum = False
    ElseIf IsNumeric(v) Then
        EsNum = True
    ElseIf IsDate(v) Then
        EsNum = True
    End If
End Function

Public Function Fecha(ByVal v As Variant) As Double
    ' fecha como número de serie (0 si no es fecha); acepta texto dd/mm/aaaa
    Dim p() As String
    Fecha = 0
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then
        p = Split(Trim(CStr(v)), "/")
        If UBound(p) = 2 Then
            If EsNumeroPunto(p(0)) And EsNumeroPunto(p(1)) And EsNumeroPunto(p(2)) Then
                If Val(p(2)) < 100 Then p(2) = CStr(2000 + Val(p(2)))
                On Error Resume Next
                Fecha = CDbl(DateSerial(CInt(Val(p(2))), CInt(Val(p(1))), CInt(Val(p(0)))))
                On Error GoTo 0
            End If
        End If
    ElseIf EsNum(v) Then
        Fecha = Int(Num(v))
    End If
    If Fecha < 0 Then Fecha = 0
End Function

Public Function txt(ByVal v As Variant) As String
    txt = ""
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    txt = Trim(CStr(v))
End Function

Public Function Limpio(ByVal v As Variant) As String
    Dim s As String
    s = txt(v)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Limpio = Trim(s)
End Function

Public Function Normal(ByVal v As Variant) As String
    ' para comparar textos: mayúsculas, sin tildes y sin espacios dobles
    Dim s As String
    s = UCase(Limpio(v))
    s = Replace(s, "Á", "A")
    s = Replace(s, "É", "E")
    s = Replace(s, "Í", "I")
    s = Replace(s, "Ó", "O")
    s = Replace(s, "Ú", "U")
    s = Replace(s, "Ü", "U")
    s = Replace(s, "À", "A")
    s = Replace(s, "È", "E")
    s = Replace(s, "Ì", "I")
    s = Replace(s, "Ò", "O")
    s = Replace(s, "Ù", "U")
    Normal = s
End Function

Public Function EsDigitos(ByVal t As String) As Boolean
    Dim i As Long, c As String
    EsDigitos = False
    If t = "" Then Exit Function
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    EsDigitos = True
End Function

Public Function Clave(ByVal s As String) As String
    ' clave para comparar partidas: "01.02.10" = "1.2.10"
    Dim p() As String, i As Long, t As String, r As String
    s = UCase(Trim(s))
    Clave = ""
    If s = "" Then Exit Function
    p = Split(s, ".")
    r = ""
    For i = LBound(p) To UBound(p)
        t = Trim(p(i))
        If t <> "" Then
            If EsDigitos(t) Then
                Do While Len(t) > 1 And Left(t, 1) = "0"
                    t = Mid(t, 2)
                Loop
            End If
            If r = "" Then
                r = t
            Else
                r = r & "." & t
            End If
        End If
    Next i
    Clave = r
End Function


'========================================================================== AUXILIARES
Public Function Buscar(ByVal c As Collection, ByVal k As String) As Long
    On Error GoTo NoHay
    Buscar = CLng(c.Item(k))
    Exit Function
NoHay:
    Buscar = 0
End Function

Public Function Redondear(ByVal x As Double, ByVal n As Long) As Double
    Dim f As Double
    On Error GoTo Alterno
    Redondear = Application.WorksheetFunction.Round(x, n)
    Exit Function
Alterno:
    f = 10 ^ n
    If x >= 0 Then
        Redondear = Int(x * f + 0.5 + 0.000000001) / f
    Else
        Redondear = -Int(-x * f + 0.5 + 0.000000001) / f
    End If
End Function

Public Function NombreMes(ByVal t As Double) As String
    NombreMes = UCase(Choose(Month(CDate(t)), "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", _
                             "agosto", "setiembre", "octubre", "noviembre", "diciembre"))
End Function

Public Function Hoja(ByVal nombre As String) As Worksheet
    Set Hoja = ThisWorkbook.Worksheets(nombre)
End Function

Public Function UltimaFila(ByVal ws As Worksheet, ByVal col As Long, ByVal F0 As Long, ByVal f1 As Long) As Long
    ' última fila con dato en la columna (entre f0 y f1); f0 - 1 si no hay datos
    Dim v As Variant, i As Long
    UltimaFila = F0 - 1
    v = ws.Range(ws.Cells(F0, col), ws.Cells(f1, col)).Value
    For i = UBound(v, 1) To 1 Step -1
        If txt(v(i, 1)) <> "" Then
            UltimaFila = F0 + i - 1
            Exit Function
        End If
    Next i
End Function

Public Function UltimaFilaVarias(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, ByVal F0 As Long, _
                                 ByVal f1 As Long) As Long
    ' última fila con algún dato entre las columnas c1 y c2
    Dim v As Variant, i As Long, j As Long
    UltimaFilaVarias = F0 - 1
    v = ws.Range(ws.Cells(F0, c1), ws.Cells(f1, c2)).Value
    For i = UBound(v, 1) To 1 Step -1
        For j = 1 To UBound(v, 2)
            If txt(v(i, j)) <> "" Then
                UltimaFilaVarias = F0 + i - 1
                Exit Function
            End If
        Next j
    Next i
End Function

Public Sub Inicio(ByRef calcMode As Long, ByVal msg As String)
    calcMode = 0
    On Error Resume Next
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    Application.Calculation = XL_MANUAL
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Public Sub Fin(ByVal calcMode As Long)
    On Error Resume Next
    Application.StatusBar = False
    If calcMode <> 0 Then Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculate
End Sub

Public Sub Estado(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
End Sub

Public Sub Aviso(ByVal msg As String, ByVal titulo As String, Optional ByVal icono As Long = 64)
    If gAcum Then
        gTexto = gTexto & msg & vbCrLf & vbCrLf
        Exit Sub
    End If
    If gSilencio Then Exit Sub
    MsgBox msg, icono, titulo
End Sub

Public Function Pregunta(ByVal msg As String, ByVal titulo As String, Optional ByVal porDefecto As Boolean = True) As Boolean
    ' Sí / No; en las pruebas devuelve la respuesta por defecto
    If gSilencio Then
        Pregunta = porDefecto Or gSiATodo
        Exit Function
    End If
    Pregunta = (MsgBox(msg, vbYesNo + vbQuestion, titulo) = vbYes)
End Function


'========================================================================== FECHAS DEL MES
Public Function MesInicio() As Double
    Dim t As Double
    t = Fecha(Hoja("DATOS").Range(D_MES).Value)
    If t = 0 Then t = CDbl(Date)
    MesInicio = CDbl(DateSerial(Year(CDate(t)), Month(CDate(t)), 1))
End Function

Public Function MesFin() As Double
    Dim t As Double
    t = MesInicio()
    MesFin = CDbl(DateSerial(Year(CDate(t)), Month(CDate(t)) + 1, 0))
End Function

Public Function FechaCorte() As Double
    Dim t As Double
    t = Fecha(Hoja("DATOS").Range(D_CORTE).Value)
    If t = 0 Then t = MesFin()
    FechaCorte = t
End Function

Public Function MesTexto() As String
    MesTexto = NombreMes(MesInicio()) & " " & Year(CDate(MesInicio()))
End Function

Public Function MesArchivo() As String
    ' año_mes para nombres de archivo (2026_10)
    MesArchivo = CStr(Year(CDate(MesInicio()))) & "_" & Right("0" & CStr(Month(CDate(MesInicio()))), 2)
End Function

Public Function FechaTxt(ByVal t As Double) As String
    If t <= 0 Then
        FechaTxt = ""
    Else
        FechaTxt = Format(CDate(t), "dd/mm/yyyy")
    End If
End Function


'========================================================================== CARPETAS Y PDF
Public Function sep() As String
    sep = "\"
    On Error Resume Next
    sep = Application.PathSeparator
End Function

Public Function RutaLocal(ByVal p As String) As String
    ' Excel devuelve una dirección web cuando el libro está en OneDrive: se convierte a la carpeta local
    Dim k As Long, od As String, resto As String
    RutaLocal = p
    If LCase(Left(p, 4)) <> "http" Then Exit Function
    resto = ""
    k = InStr(1, p, "d.docs.live.net/", vbTextCompare)
    If k > 0 Then
        resto = Mid(p, k + Len("d.docs.live.net/"))
        k = InStr(resto, "/")
        If k > 0 Then
            resto = Mid(resto, k + 1)
        Else
            resto = ""
        End If
        od = Environ("OneDriveConsumer")
        If od = "" Then od = Environ("OneDrive")
    Else
        k = InStr(1, p, "/Documents/", vbTextCompare)
        If k > 0 Then resto = Mid(p, k + Len("/Documents/"))
        od = Environ("OneDriveCommercial")
        If od = "" Then od = Environ("OneDrive")
    End If
    If od = "" Then Exit Function
    resto = Replace(Replace(resto, "/", "\"), "%20", " ")
    If resto = "" Then
        RutaLocal = od
    Else
        RutaLocal = od & "\" & resto
    End If
End Function

Public Function ExisteCarpeta(ByVal p As String) As Boolean
    ExisteCarpeta = False
    If p = "" Then Exit Function
    On Error Resume Next
    ExisteCarpeta = ((GetAttr(p) And vbDirectory) = vbDirectory)
End Function

Public Function CrearCarpeta(ByVal p As String) As Boolean
    Dim partes() As String, i As Long, acc As String, s As String
    CrearCarpeta = False
    If p = "" Then Exit Function
    s = sep()
    If Right(p, 1) = s Then p = Left(p, Len(p) - 1)
    If ExisteCarpeta(p) Then
        CrearCarpeta = True
        Exit Function
    End If
    On Error GoTo Falla
    partes = Split(p, s)
    acc = partes(0)
    For i = 1 To UBound(partes)
        acc = acc & s & partes(i)
        If partes(i) <> "" Then
            If Not ExisteCarpeta(acc) Then MkDir acc
        End If
    Next i
    CrearCarpeta = ExisteCarpeta(p)
    Exit Function
Falla:
    CrearCarpeta = False
End Function

Public Function ElegirCarpeta(ByVal titulo As String) As String
    Dim fd As Object
    ElegirCarpeta = ""
    If gSilencio Then Exit Function
    On Error GoTo Falla
    Set fd = Application.FileDialog(4)
    fd.Title = titulo
    If fd.Show = -1 Then ElegirCarpeta = fd.SelectedItems(1)
    Exit Function
Falla:
    ElegirCarpeta = ""
End Function

Public Function CarpetaSalida(ByVal subc As String) As String
    ' carpeta de los PDF: la de DATOS o, si está vacía, SALIDAS_AAAA_MM junto al programa
    Dim raiz As String, p As String
    CarpetaSalida = ""
    raiz = txt(Hoja("DATOS").Range(D_CARPETA).Value)
    If raiz = "" Then
        raiz = RutaLocal(ThisWorkbook.Path)
        If raiz <> "" Then raiz = raiz & sep() & "SALIDAS_" & MesArchivo()
    End If
    p = raiz
    If subc <> "" Then p = raiz & sep() & subc
    If Not CrearCarpeta(p) Then
        raiz = ElegirCarpeta("Elija la carpeta donde se guardarán los PDF")
        If raiz = "" Then Exit Function
        p = raiz
        If subc <> "" Then p = raiz & sep() & subc
        If Not CrearCarpeta(p) Then Exit Function
    End If
    CarpetaSalida = p
End Function

Public Function CarpetaDe(ByVal ruta As String) As String
    Dim k As Long
    k = InStrRev(ruta, sep())
    If k > 0 Then CarpetaDe = Left(ruta, k - 1) Else CarpetaDe = ""
End Function

Public Function RutaSugerida(ByVal subc As String, ByVal archivo As String) As String
    ' ruta propuesta para «Guardar como»: carpeta de GENERAR TODO o SALIDAS_AAAA_MM\subc
    Dim c As String
    If gCarpetaFija <> "" Then
        RutaSugerida = gCarpetaFija & sep() & archivo
        Exit Function
    End If
    c = CarpetaSalida(subc)
    If c = "" Then c = RutaLocal(ThisWorkbook.Path)
    RutaSugerida = c & sep() & archivo
End Function

Public Function GuardarComo(ByVal sugerido As String, ByVal titulo As String, ByVal tipo As String) As String
    ' ventana «Guardar como» (PDF o XLSX); devuelve la ruta elegida o "" si se cancela
    Dim f As Variant, filtro As String
    GuardarComo = ""
    If gCarpetaFija <> "" Or gSilencio Then
        GuardarComo = sugerido
        Exit Function
    End If
    If tipo = "PDF" Then
        filtro = "Archivo PDF (*.pdf),*.pdf"
    Else
        filtro = "Libro de Excel (*.xlsx),*.xlsx"
    End If
    On Error GoTo Falla
    f = Application.GetSaveAsFilename(InitialFileName:=sugerido, FileFilter:=filtro, Title:=titulo)
    If VarType(f) = vbBoolean Then Exit Function
    GuardarComo = CStr(f)
    If tipo = "PDF" And LCase(Right(GuardarComo, 4)) <> ".pdf" Then GuardarComo = GuardarComo & ".pdf"
    Exit Function
Falla:
    GuardarComo = sugerido
End Function

Public Function SubcarpetaPor(ByVal carpeta As String, ByVal nombre As String, ByVal sufijo As String) As String
    ' carpeta de los PDF individuales: ...\KARDEX\POR_INSUMO o ...\KARDEX_POR_INSUMO
    Dim ult As String
    ult = Mid(carpeta, InStrRev(carpeta, sep()) + 1)
    If UCase(ult) = UCase(nombre) Then
        SubcarpetaPor = carpeta & sep() & sufijo
    Else
        SubcarpetaPor = carpeta & sep() & nombre & "_" & sufijo
    End If
End Function

Public Function ElegirCarpetaDesde(ByVal titulo As String, ByVal inicial As String) As String
    Dim fd As Object
    ElegirCarpetaDesde = ""
    If gSilencio Then
        ElegirCarpetaDesde = inicial
        Exit Function
    End If
    On Error GoTo Falla
    Set fd = Application.FileDialog(4)
    fd.Title = titulo
    If inicial <> "" Then fd.InitialFileName = inicial & sep()
    If fd.Show = -1 Then ElegirCarpetaDesde = fd.SelectedItems(1)
    Exit Function
Falla:
    ElegirCarpetaDesde = ""
End Function

Public Function NombreSeguro(ByVal s As String, ByVal lim As Long) As String
    Dim i As Long, c As String, r As String
    r = ""
    s = Limpio(s)
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case "\", "/", ":", "*", "?", """", "<", ">", "|", "."
                r = r & " "
            Case Else
                r = r & c
        End Select
    Next i
    r = Limpio(r)
    If Len(r) > lim Then r = Trim(Left(r, lim))
    NombreSeguro = r
End Function

Public Function ExportarPDF(ByVal ws As Worksheet, ByVal archivo As String) As Boolean
    ExportarPDF = False
    If gSinPDF Then
        ExportarPDF = True
        Exit Function
    End If
    On Error GoTo Falla
    ws.ExportAsFixedFormat Type:=0, Filename:=archivo, Quality:=0, IncludeDocProperties:=False, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=False
    ExportarPDF = True
    Exit Function
Falla:
    ExportarPDF = False
End Function

Public Sub BorrarFormas(ByVal ws As Worksheet)
    ' borra recuadros, logos y botones de una hoja de salida (KARDEX, BINCARD, PARTE_DIARIO)
    Dim i As Long
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Public Sub CopiarPagina(ByVal wsF As Worksheet, ByVal wsK As Worksheet)
    ' la hoja de salida se imprime igual que su formato (papel, orientación, escala, márgenes y centrado)
    On Error Resume Next
    Application.PrintCommunication = False
    With wsK.PageSetup
        .PaperSize = wsF.PageSetup.PaperSize
        .Orientation = wsF.PageSetup.Orientation
        .Zoom = wsF.PageSetup.Zoom
        .LeftMargin = wsF.PageSetup.LeftMargin
        .RightMargin = wsF.PageSetup.RightMargin
        .TopMargin = wsF.PageSetup.TopMargin
        .BottomMargin = wsF.PageSetup.BottomMargin
        .HeaderMargin = wsF.PageSetup.HeaderMargin
        .FooterMargin = wsF.PageSetup.FooterMargin
        .CenterHorizontally = wsF.PageSetup.CenterHorizontally
        .CenterVertically = wsF.PageSetup.CenterVertically
    End With
    Application.PrintCommunication = True
End Sub

Public Sub AjustarTexto(ByVal c As Range, ByVal cap1 As Long, ByVal cap2 As Long, ByVal tam As Double)
    ' textos largos en celdas angostas: hasta cap1 caracteres se dejan igual; hasta cap2 pasan a dos
    ' renglones con letra tam; más largos se reducen para que entren en la celda
    Dim n As Long
    If IsError(c.Value) Then Exit Sub
    n = Len(CStr(c.Value))
    If n <= cap1 Then Exit Sub
    On Error Resume Next
    With c.MergeArea
        .Font.Size = tam
        If n <= cap2 Then
            .ShrinkToFit = False
            .WrapText = True
        Else
            .WrapText = False
            .ShrinkToFit = True
        End If
    End With
End Sub

'========================================================================== LOGO DE LOS FORMATOS
Public Sub CambiarLogo()
    ' inserta o cambia el logo del KARDEX y del PARTE DIARIO (hojas FMT_KARDEX y FMT_PARTE)
    Dim f As Variant, n As Long, ws As Worksheet
    If gSilencio Then Exit Sub
    f = Application.GetOpenFilename("Imágenes (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.emf;*.wmf)," & _
                                    "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.emf;*.wmf", , "Elija el logo de la entidad")
    If VarType(f) = vbBoolean Then Exit Sub
    If PonerLogo(Hoja("FMT_KARDEX"), CStr(f), "C2:D4") Then n = n + 1
    If PonerLogo(Hoja("FMT_PARTE"), CStr(f), "C2:D5") Then n = n + 1
    If n = 2 Then
        Aviso "Logo colocado en los formatos del KARDEX y del PARTE DIARIO." & vbCrLf & vbCrLf & _
              "Saldrá en las hojas y PDF que se generen desde ahora (botones KARDEX + PDF, PARTES DIARIOS + PDF " & _
              "o GENERAR TODO)." & vbCrLf & vbCrLf & "Guarde el programa para conservarlo.", "LOGO DE LOS FORMATOS"
        Hoja("FMT_KARDEX").Activate
    Else
        Aviso "No se pudo colocar el logo (" & n & " de 2 formatos). Pruebe con otra imagen (PNG o JPG).", _
              "LOGO DE LOS FORMATOS", 48
    End If
End Sub

Public Function PonerLogo(ByVal ws As Worksheet, ByVal archivo As String, ByVal marco As String) As Boolean
    ' reemplaza la imagen LOGO de la hoja por la del archivo, centrada dentro del marco (celdas)
    Dim shp As Object, pic As Object, l As Double, t As Double, w As Double, h As Double, r As Double
    PonerLogo = False
    On Error GoTo Falla
    l = ws.Range(marco).Left + 2
    t = ws.Range(marco).Top + 2
    w = ws.Range(marco).Width - 4
    h = ws.Range(marco).Height - 4
    Set pic = ws.Shapes.AddPicture(archivo, 0, -1, l, t, -1, -1)
    For Each shp In ws.Shapes
        If shp.Name = "LOGO" Then
            shp.Delete
            Exit For
        End If
    Next shp
    pic.LockAspectRatio = -1
    r = w / pic.Width
    If h / pic.Height < r Then r = h / pic.Height
    pic.Width = pic.Width * r
    If pic.Height > h Then pic.Height = h
    pic.Left = l + (w - pic.Width) / 2
    pic.Top = t + (h - pic.Height) / 2
    pic.Placement = 2
    pic.Name = "LOGO"
    PonerLogo = True
    Exit Function
Falla:
    PonerLogo = False
End Function

Public Sub Mensaje(ByVal celda As String, ByVal texto As String)
    On Error Resume Next
    Hoja("DATOS").Range(celda).NumberFormat = "@"
    Hoja("DATOS").Range(celda).Value = texto
End Sub


'========================================================================== ORDEN
Public Sub Ordenar(ByRef llave() As Double, ByRef idx() As Long, ByVal n As Long)
    ' ordena idx(1..n) según llave(idx) de menor a mayor (llaves únicas)
    Dim gap As Long, i As Long, j As Long, t As Long
    gap = 1
    Do While gap < n \ 3
        gap = gap * 3 + 1
    Loop
    Do While gap >= 1
        For i = gap + 1 To n
            t = idx(i)
            j = i
            Do While j > gap
                If llave(idx(j - gap)) <= llave(t) Then Exit Do
                idx(j) = idx(j - gap)
                j = j - gap
            Loop
            idx(j) = t
        Next i
        gap = gap \ 3
    Loop
End Sub

Public Sub ModoPrueba()
    ' pruebas automáticas: sin ventanas y sin exportar PDF
    gSilencio = True
    gSinPDF = True
End Sub

Public Sub ModoPruebaPDF()
    ' pruebas automáticas: sin ventanas, exportando los PDF
    gSilencio = True
    gSinPDF = False
End Sub

Public Sub ModoPruebaSi()
    ' pruebas automáticas: sin ventanas, sin PDF y respondiendo Sí a todo
    gSilencio = True
    gSinPDF = True
    gSiATodo = True
End Sub
