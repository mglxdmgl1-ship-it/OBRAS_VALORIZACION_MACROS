Attribute VB_Name = "ModMenu"
Option Explicit

'==========================================================================================
'  MENÚ: generar todo, abrir la carpeta de salidas, cerrar el mes y nueva obra
'==========================================================================================

Public Sub GenerarTodo()
    ' GENERAR LOS TRES: actualiza los cuadros y arma kardex, bincard y partes en sus hojas (sin PDF)
    Dim n As Long, s As String
    gAcum = True
    gTexto = ""
    On Error GoTo Sale
    s = ActualizarCuadros(True)
    gTexto = "ACTUALIZAR CUADROS" & vbCrLf & s & vbCrLf & vbCrLf
    n = ArmarTarjetas(True, False, False)
    n = ArmarTarjetas(False, False, False)
    n = ArmarPartes(False, False)
Sale:
    gAcum = False
    If Err.Number <> 0 Then gTexto = gTexto & "Error: " & Err.Description
    Aviso gTexto & "Revise las hojas KARDEX, BINCARD y PARTE_DIARIO; después exporte con los botones PDF (o PDF DE TODO).", _
          "GENERAR LOS TRES"
    gTexto = ""
End Sub

Public Sub AbrirCarpeta()
    Dim p As String
    p = CarpetaSalida("")
    If p = "" Then Exit Sub
    On Error Resume Next
    Shell "explorer.exe """ & p & """", 1
End Sub

Public Function Respaldo(ByVal sufijo As String) As String
    Dim raiz As String, nom As String, k As Long, ruta As String
    Respaldo = ""
    raiz = RutaLocal(ThisWorkbook.Path)
    nom = ThisWorkbook.Name
    k = InStrRev(nom, ".")
    If k > 0 Then nom = Left(nom, k - 1)
    ruta = raiz & Sep() & nom & "_" & sufijo & ".xlsm"
    On Error GoTo Falla
    ThisWorkbook.SaveCopyAs ruta
    Respaldo = ruta
    Exit Function
Falla:
    Respaldo = ""
End Function

Private Sub VaciarColumnas(ByVal ws As Worksheet, ByVal f0 As Long, ByVal f1 As Long, ByVal desde As Variant, ByVal hasta As Variant)
    Dim i As Long
    For i = LBound(desde) To UBound(desde)
        ws.Range(ws.Cells(f0, desde(i)), ws.Cells(f1, hasta(i))).ClearContents
    Next i
End Sub

Private Sub VaciarSalidas()
    Dim n As Variant
    For Each n In Array("KARDEX", "BINCARD", "PARTE_DIARIO")
        On Error Resume Next
        Hoja(n).ResetAllPageBreaks
        BorrarFormas Hoja(n)
        Hoja(n).Cells.Clear
        On Error GoTo 0
    Next n
    On Error Resume Next
    ThisWorkbook.Worksheets(HOJA_A4).ResetAllPageBreaks
    BorrarFormas ThisWorkbook.Worksheets(HOJA_A4)
    ThisWorkbook.Worksheets(HOJA_A4).Cells.Clear
    On Error GoTo 0
End Sub

Public Sub CerrarMes()
    Dim cm As Long, ruta As String, msg As String, t As Double, s As String, mesTxt As String
    mesTxt = MesTexto()
    If Not Pregunta("CERRAR EL MES DE " & mesTxt & ":" & vbCrLf & vbCrLf & _
                    "• Se guarda una copia de respaldo del programa." & vbCrLf & _
                    "• En el 4.10, el SALDO de cada ítem pasa a SALDO ANTERIOR y se vacían los días." & vbCrLf & _
                    "• En el 4.15, las horas del mes pasan al ACUMULADO ANTERIOR de EQUIPOS y se vacían los días." & vbCrLf & _
                    "• El INV. 2026 se conserva (es el inventario de todo el año)." & vbCrLf & vbCrLf & _
                    "Antes de cerrar: genere los PDF y el informe mensual. ¿Continuar?", "CERRAR MES", False) Then Exit Sub
    If Not Pregunta("¿Seguro que desea cerrar " & mesTxt & "? (queda la copia de respaldo)", "CERRAR MES", False) Then Exit Sub
    ruta = Respaldo("CIERRE_" & MesArchivo())
    If ruta = "" Then
        If Not Pregunta("No se pudo guardar la copia de respaldo junto al programa. ¿Cerrar el mes de todos modos?", _
                        "CERRAR MES", False) Then Exit Sub
    End If
    s = ActualizarCuadros(True)
    Inicio cm, "Cerrando el mes..."
    CerrarCuadros
    VaciarSalidas
    t = MesInicio()
    Hoja("DATOS").Range(D_MES).Value = DateSerial(Year(CDate(t)), Month(CDate(t)) + 1, 1)
    Hoja("DATOS").Range(D_CORTE).Value = DateSerial(Year(CDate(t)), Month(CDate(t)) + 2, 0)
    s = Txt(Hoja("DATOS").Range(D_NINF).Value)
    If EsDigitos(s) Then
        Hoja("DATOS").Range(D_NINF).NumberFormat = "@"
        Hoja("DATOS").Range(D_NINF).Value = Format(Val(s) + 1, String(Len(s), "0"))
    End If
    Fin cm
    s = ActualizarCuadros(True)
    Mensaje D_MSG, "Mes de " & mesTxt & " cerrado el " & Format(Now, "dd/mm/yyyy hh:mm") & "."
    Mensaje D_MSGPDF, ""
    Mensaje D_MSGINF, ""
    msg = "Se cerró " & mesTxt & ". Ahora el programa está en " & MesTexto() & "." & vbCrLf & vbCrLf & _
          "Los saldos pasaron a SALDO ANTERIOR (4.10) y las horas al ACUMULADO ANTERIOR (EQUIPOS); los días quedaron vacíos."
    If ruta <> "" Then msg = msg & vbCrLf & vbCrLf & "Copia de respaldo:" & vbCrLf & ruta
    Aviso msg, "CERRAR MES"
End Sub

Public Sub NuevaObra()
    Dim cm As Long, ruta As String, c As Variant
    If Not Pregunta("NUEVA OBRA: se borran los datos de la obra, el INV. 2026, los cuadros 4.10 al 4.15 y los EQUIPOS " & _
                    "(antes se guarda una copia de respaldo)." & vbCrLf & vbCrLf & "¿Continuar?", "NUEVA OBRA", False) Then Exit Sub
    If Not Pregunta("¿Seguro? Se vaciará todo el programa para empezar otra obra.", "NUEVA OBRA", False) Then Exit Sub
    ruta = Respaldo("RESPALDO_" & Format(Now, "yyyymmdd_hhmm"))
    Inicio cm, "Vaciando el programa..."
    For Each c In Array(D_ENTIDAD, D_GEREN, D_SUBG, D_PROY, D_CUI, D_META, D_UBIC, D_NINF, D_CARPETA, D_CTRL, D_INF410, _
                        D_INF415, D_MSG, D_MSGPDF, D_MSGINF, D_RESID, D_SUPER)
        Hoja("DATOS").Range(c).ClearContents
    Next c
    VaciarColumnas Hoja("CATALOGO"), CT0, CTMAX, Array(CC_COD, CC_ACU), Array(CC_SANT, CC_ULT)
    VaciarColumnas Hoja("MOVIMIENTOS"), MV0, MVMAX, Array(MC_FEC, MC_CANT, MC_SALDO), Array(MC_COD, MC_ORIG, MC_SALDO)
    VaciarColumnas Hoja("EQUIPOS"), EQ0, EQMAX, Array(EC_COD, EC_COND, EC_ACU), Array(EC_CU, EC_ANT, EC_ULT)
    VaciarColumnas Hoja("PARTES"), PT0, PTMAX, Array(PC_FEC, PC_NPARTE, PC_HREAL, PC_COMB, PC_CTRL), _
                   Array(PC_EQ, PC_REND, PC_HREAL, PC_ORIG, PC_CTRL)
    VaciarSalidas
    On Error Resume Next
    VaciarCuadros
    On Error GoTo 0
    Fin cm
    Hoja("DATOS").Activate
    If ruta = "" Then
        Aviso "El programa quedó vacío. Registre los datos de la nueva obra en esta hoja.", "NUEVA OBRA"
    Else
        Aviso "El programa quedó vacío. Registre los datos de la nueva obra en esta hoja." & vbCrLf & vbCrLf & _
              "Copia de respaldo:" & vbCrLf & ruta, "NUEVA OBRA"
    End If
End Sub
