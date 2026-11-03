Attribute VB_Name = "ModUtil"
Option Explicit

'==========================================================================================
'  CONTROL DE METRADOS DIARIOS - UTILIDADES COMUNES
'  Constantes de las hojas, lectura segura de celdas, claves de partida y textos.
'==========================================================================================

Public Const SENT As Double = 73051           ' 31/12/2099 = sin ingreso / sin limite
Public Const ND As Long = 31                  ' dias del mes
' ---- RESUMEN
Public Const RR0 As Long = 10                 ' primera fila de partidas
Public Const RRMAX As Long = 5009             ' ultima fila disponible
Public Const CD1 As Long = 9                  ' columna del dia 1 (I)
' ---- PRESUPUESTO
Public Const PP0 As Long = 8                  ' primera fila
Public Const PPMAX As Long = 5007             ' ultima fila
' ---- ACU
Public Const AC0 As Long = 8
Public Const ACMAX As Long = 40007
' ---- INSUMOS
Public Const IN0 As Long = 8
Public Const INMAX As Long = 3007
' ---- INGRESOS
Public Const IG0 As Long = 8
Public Const IG1 As Long = 1007
' ---- MAQ_COMB y MANO_OBRA
Public Const MQ0 As Long = 8
Public Const MQMAX As Long = 900
Public Const MO0 As Long = 8
Public Const MOMAX As Long = 400
' ---- planillas diarias
Public Const PL0 As Long = 9
Public Const PLMAX As Long = 1500
' ---- constantes de Excel
Public Const XL_MANUAL As Long = -4135
Public Const XL_NONE As Long = -4142
Public Const XL_CONT As Long = 1
Public Const XL_THIN As Long = 2
Public Const XL_HIDDEN As Long = 0
Public Const XL_VISIBLE As Long = -1
Public Const XL_CENTER As Long = -4108
Public Const XL_LEFT As Long = -4131
Public Const XL_RIGHT As Long = -4152
Public Const SEP As String = "|#|"
' ---- etiquetas de los bloques (las mismas que escribe la plantilla)
Public Const LB_ACU As String = "SEGÚN A.C.U. (A VALORIZAR)"
Public Const LB_PARTE As String = "SEGÚN PARTE DIARIO (REAL)"
Public Const LB_VALES As String = "SEGÚN VALES DE DESPACHO (REAL)"
Public Const LB_DIFP As String = "DIFERENCIA (PARTE - A.C.U.)"
Public Const LB_DIFV As String = "DIFERENCIA (VALES - A.C.U.)"
Public Const LB_TOTEQ As String = "TOTAL VALORIZABLE EN EQUIPOS (S/)"
Public Const LB_HHACU As String = "HH SEGÚN A.C.U."
Public Const LB_TRAB As String = "N° DE TRABAJADORES (TAREO)"
Public Const LB_HHTAR As String = "HH DEL TAREO"
Public Const LB_USO As String = "% DE USO (A.C.U. / TAREO)"
Public Const LB_THHACU As String = "TOTAL HH SEGÚN A.C.U."
Public Const LB_THHTAR As String = "TOTAL HH DEL TAREO"
Public Const LB_TUSO As String = "% DE USO TOTAL"
' ---- celdas de la hoja DATOS
Public Const D_ENTIDAD As String = "C5"
Public Const D_DEPEND As String = "C6"
Public Const D_PROY As String = "C7"
Public Const D_CUI As String = "C8"
Public Const D_META As String = "C9"
Public Const D_UBIC As String = "C10"
Public Const D_MODAL As String = "C11"
Public Const D_INICIO As String = "C14"
Public Const D_PLAZO As String = "C15"
Public Const D_MES As String = "C19"
Public Const D_VALN As String = "C21"
Public Const D_UMBRAL As String = "C24"
Public Const D_JORNADA As String = "C25"
Public Const D_USO As String = "C26"
Public Const D_NOCTRL As String = "C27"
Public Const D_COMB As String = "C28"
Public Const D_FIRMA1 As String = "C31"
Public Const D_FIRMA2 As String = "C32"
Public Const D_FIRMA3 As String = "C33"
Public Const D_ULTACT As String = "C36"
Public Const D_HUELLA As String = "C37"
Public Const D_HUELLAHOY As String = "C38"
Public Const D_MSGACT As String = "C39"
Public Const D_ULTPREP As String = "C40"
Public Const D_HUELLAP As String = "C41"
Public Const D_HUELLAPHOY As String = "C42"
Public Const D_MSGPREP As String = "C43"


Public gSilencio As Boolean                   ' verdadero en las pruebas automáticas (sin ventanas)


'========================================================================== LECTURA SEGURA
Public Function ParseNum(ByVal s As String, ByRef ok As Boolean) As Double
    ' número escrito como texto ("1,234.50", "12.5", "0,25"); ok = falso si no es número
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
    ' dígitos con un punto decimal opcional y signo opcional
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
    ElseIf IsNumeric(v) Then
        Num = CDbl(v)
    ElseIf IsDate(v) Then
        Num = CDbl(CDate(v))
    End If
End Function

Public Function EsNum(ByVal v As Variant) As Boolean
    ' verdadero si la celda trae un número (también un número escrito como texto)
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
    End If
End Function

Public Function Txt(ByVal v As Variant) As String
    Txt = ""
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    Txt = Trim(CStr(v))
End Function

Public Function Limpio(ByVal v As Variant) As String
    ' texto sin saltos de línea, tabulaciones ni espacios dobles
    Dim s As String
    s = Txt(v)
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


'========================================================================== CÓDIGOS DE PARTIDA
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
    ' clave para comparar ítems: "01.02.10" = "1.2.10" (sin ceros a la izquierda)
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

Public Function Nivel(ByVal s As String) As Long
    Dim k As String
    k = Clave(s)
    If k = "" Then
        Nivel = 1
    Else
        Nivel = Len(k) - Len(Replace(k, ".", "")) + 1
    End If
End Function

Public Function EsCodigo(ByVal s As String) As Boolean
    ' "01", "01.02", "1.2.3" (solo dígitos y puntos, empieza con dígito)
    Dim i As Long, c As String, t As String
    EsCodigo = False
    t = Trim(s)
    If Right(t, 1) = "." Then t = Left(t, Len(t) - 1)
    If t = "" Then Exit Function
    If Left(t, 1) < "0" Or Left(t, 1) > "9" Then Exit Function
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If Not ((c >= "0" And c <= "9") Or c = ".") Then Exit Function
    Next i
    If InStr(t, "..") > 0 Then Exit Function
    EsCodigo = True
End Function


'========================================================================== AUXILIARES
Public Function Buscar(ByVal c As Collection, ByVal k As String) As Long
    On Error GoTo NoHay
    Buscar = CLng(c.Item(k))
    Exit Function
NoHay:
    Buscar = 0
End Function

Public Function Minimo(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        Minimo = a
    Else
        Minimo = b
    End If
End Function

Public Function Redondear(ByVal x As Double, ByVal n As Long) As Double
    ' mismo redondeo que la función REDONDEAR/ROUND de Excel
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

Public Function Unir(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        Unir = b
    Else
        Unir = a & " | " & b
    End If
End Function

Public Function Agregar(ByVal s As String, ByVal nom As String) As String
    If s = "" Then
        Agregar = nom
    ElseIf InStr(1, SEP & s & SEP, SEP & nom & SEP, vbBinaryCompare) > 0 Then
        Agregar = s
    Else
        Agregar = s & SEP & nom
    End If
End Function

Public Function Cortar(ByVal s As String, ByVal lim As Long) As String
    Dim t As String
    t = Replace(s, SEP, ", ")
    If Len(t) > lim Then t = Left(t, lim) & "…"
    Cortar = t
End Function

Public Function NombreDia(ByVal t As Double) As String
    Select Case Weekday(CDate(t))
        Case 1
            NombreDia = "DOMINGO"
        Case 2
            NombreDia = "LUNES"
        Case 3
            NombreDia = "MARTES"
        Case 4
            NombreDia = "MIÉRCOLES"
        Case 5
            NombreDia = "JUEVES"
        Case 6
            NombreDia = "VIERNES"
        Case Else
            NombreDia = "SÁBADO"
    End Select
End Function

Public Function NombreMes(ByVal t As Double) As String
    NombreMes = UCase(Choose(Month(CDate(t)), "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", _
                             "agosto", "setiembre", "octubre", "noviembre", "diciembre"))
End Function

Public Function ColorNivel(ByVal lv As Long) As Long
    Select Case lv
        Case 1
            ColorNivel = RGB(192, 0, 0)
        Case 2
            ColorNivel = RGB(0, 0, 255)
        Case 3
            ColorNivel = RGB(0, 176, 80)
        Case 4
            ColorNivel = RGB(0, 160, 176)
        Case 5
            ColorNivel = RGB(204, 0, 153)
        Case 6
            ColorNivel = RGB(140, 122, 0)
        Case Else
            ColorNivel = 0
    End Select
End Function

Public Function Hoja(ByVal nombre As String) As Worksheet
    Set Hoja = ThisWorkbook.Worksheets(nombre)
End Function

Public Function UltimaFila(ByVal ws As Worksheet, ByVal col As Long, ByVal f0 As Long, ByVal f1 As Long) As Long
    ' última fila con dato en la columna (entre f0 y f1); f0 - 1 si no hay datos
    Dim v As Variant, i As Long
    UltimaFila = f0 - 1
    v = ws.Range(ws.Cells(f0, col), ws.Cells(f1, col)).Value
    For i = UBound(v, 1) To 1 Step -1
        If Txt(v(i, 1)) <> "" Then
            UltimaFila = f0 + i - 1
            Exit Function
        End If
    Next i
End Function

Public Sub Inicio(ByRef calcMode As Long, ByVal aviso As String)
    calcMode = 0
    On Error Resume Next
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    Application.Calculation = XL_MANUAL
    Application.StatusBar = aviso
    On Error GoTo 0
End Sub

Public Sub Fin(ByVal calcMode As Long)
    On Error Resume Next
    Application.PrintCommunication = True
    Application.StatusBar = False
    If calcMode <> 0 Then Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculate
End Sub

Public Function Silencio() As Boolean
    Silencio = gSilencio
End Function
