Attribute VB_Name = "ModProcesar"
Option Explicit

'==========================================================================================
'  PROCESAR Y VALIDAR: saldo después de cada movimiento, saldos negativos, control diario de
'  horómetro/horario en los partes y saldo de las órdenes de servicio
'==========================================================================================

Public Sub ProcesarTodo()
    Dim msg As String
    msg = Procesar()
    Aviso msg, "PROCESAR Y VALIDAR"
End Sub

Public Function Procesar() As String
    Dim cm As Long, ws As Worksheet, last As Long, j As Long, i As Long, k As Long, x As Long
    Dim stock() As Double, negK() As Boolean, nNeg As Long, neg As String, nObs As Long, v As Variant
    Dim outQ() As Variant, outY() As Variant, g0 As Long, g1 As Long, f As Double, nG As Long, nDif As Long
    Dim hIni As Double, hFin As Double, hHor As Double, totH As Double, t As String, dif As String
    Dim tieneHor As Boolean, nExc As Long, exc As String, msg As String, rowsP As Long

    Inicio cm, "Procesando el almacén..."
    LeerAlmacen
    Set ws = Hoja("MOVIMIENTOS")
    last = UltimaFilaVarias(ws, MC_FEC, MC_CANT, MV0, MVMAX)
    ws.Range(ws.Cells(MV0, MC_SALDO), ws.Cells(MVMAX, MC_SALDO)).ClearContents
    If anCat > 0 Then
        ReDim stock(1 To anCat)
        ReDim negK(1 To anCat)
        For k = 1 To anCat
            stock(k) = aAnt(k)
        Next k
    End If
    If last >= MV0 Then
        ReDim outQ(1 To last - MV0 + 1, 1 To 1)
        For j = 1 To anMov
            i = vOrden(j)
            k = vItem(i)
            If vTipo(i) = T_ING Then
                stock(k) = stock(k) + vCant(i)
            Else
                stock(k) = stock(k) - vCant(i)
            End If
            outQ(vFila(i) - MV0 + 1, 1) = Redondear(stock(k), 4)
            If stock(k) < -0.00001 And Not negK(k) Then
                negK(k) = True
                nNeg = nNeg + 1
                If nNeg <= 12 Then neg = neg & vbCrLf & "   • " & aCod(k) & "  " & Left(aDes(k), 45) & "  (" & _
                                         FechaTxt(vFec(i)) & ": " & Format(stock(k), "#,##0.00") & ")"
            End If
        Next j
        ws.Range(ws.Cells(MV0, MC_SALDO), ws.Cells(last, MC_SALDO)).Value = outQ
        Application.Calculate
        v = ws.Range(ws.Cells(MV0, 16), ws.Cells(last, 16)).Value
        For i = 1 To UBound(v, 1)
            If Txt(v(i, 1)) <> "" And Txt(v(i, 1)) <> "OK" Then nObs = nObs + 1
        Next i
    End If

    ' ---- partes: horómetro y horario frente a las horas a valorizar de cada día
    Estado "Procesando los partes diarios..."
    LeerPartes
    Set ws = Hoja("PARTES")
    ws.Range(ws.Cells(PT0, PC_CTRL), ws.Cells(PTMAX, PC_CTRL)).ClearContents
    If Not IsEmpty(pV) Then
        rowsP = UBound(pV, 1)
        ReDim outY(1 To rowsP, 1 To 1)
        j = 1
        Do While j <= pnPar
            g0 = j
            k = pEq(pOrden(j))
            f = pFec(pOrden(j))
            Do While j <= pnPar
                If pEq(pOrden(j)) <> k Or pFec(pOrden(j)) <> f Then Exit Do
                j = j + 1
            Loop
            g1 = j - 1
            nG = nG + 1
            hIni = 0
            hFin = 0
            hHor = 0
            totH = 0
            tieneHor = False
            For x = g0 To g1
                i = pR(pOrden(x))
                If EsNum(pV(i, PC_HINI)) Then
                    If hIni = 0 Or Num(pV(i, PC_HINI)) < hIni Then hIni = Num(pV(i, PC_HINI))
                End If
                If EsNum(pV(i, PC_HFIN)) Then
                    If Num(pV(i, PC_HFIN)) > hFin Then hFin = Num(pV(i, PC_HFIN))
                End If
                If Not tieneHor Then
                    If EsNum(pV(i, PC_MDE)) Or EsNum(pV(i, PC_TDE)) Then
                        tieneHor = True
                        hHor = Tramo(pV(i, PC_MDE), pV(i, PC_MA)) + Tramo(pV(i, PC_TDE), pV(i, PC_TA))
                    End If
                End If
                totH = totH + HorasFila(i)
            Next x
            If hIni > 0 And hFin > 0 Then
                If Abs((hFin - hIni) - totH) > 0.05 Then
                    t = "REVISAR: horómetro " & Format(hFin - hIni, "0.00") & " h y horas a valorizar " & Format(totH, "0.00") & " h"
                    nDif = nDif + 1
                    If nDif <= 10 Then dif = dif & vbCrLf & "   • " & eCod(k) & " " & FechaTxt(f) & ": horómetro " & _
                                         Format(hFin - hIni, "0.00") & " h / partes " & Format(totH, "0.00") & " h"
                Else
                    t = "OK (horómetro " & Format(hFin - hIni, "0.00") & " h)"
                End If
            ElseIf tieneHor And hHor > 0 Then
                If totH > hHor + 0.05 Then
                    t = "REVISAR: horas a valorizar " & Format(totH, "0.00") & " h mayores que el horario " & Format(hHor, "0.00") & " h"
                    nDif = nDif + 1
                    If nDif <= 10 Then dif = dif & vbCrLf & "   • " & eCod(k) & " " & FechaTxt(f) & ": horario " & _
                                         Format(hHor, "0.00") & " h / partes " & Format(totH, "0.00") & " h"
                Else
                    t = "OK (horario " & Format(hHor, "0.00") & " h)"
                End If
            Else
                t = "Sin horómetro ni horario: " & Format(totH, "0.00") & " h del día"
            End If
            For x = g0 To g1
                outY(pR(pOrden(x)), 1) = t
            Next x
        Loop
        ws.Range(ws.Cells(PT0, PC_CTRL), ws.Cells(PT0 + rowsP - 1, PC_CTRL)).Value = outY
    End If

    ' ---- equipos que superan lo contratado en la O/S
    Application.Calculate
    Set ws = Hoja("EQUIPOS")
    last = UltimaFila(ws, EC_COD, EQ0, EQMAX)
    If last >= EQ0 Then
        v = ws.Range(ws.Cells(EQ0, 1), ws.Cells(last, EC_ULT)).Value
        For i = 1 To UBound(v, 1)
            If Txt(v(i, EC_COD)) <> "" And EsNum(v(i, EC_CANT)) Then
                If Num(v(i, EC_CANT)) > 0 And Num(v(i, EC_SALDO)) < -0.005 Then
                    nExc = nExc + 1
                    exc = exc & vbCrLf & "   • " & Txt(v(i, EC_COD)) & " " & Left(Limpio(v(i, EC_DES)), 40) & ": excede en " & _
                          Format(-Num(v(i, EC_SALDO)), "#,##0.00") & " " & Txt(v(i, EC_UND))
                End If
            End If
        Next i
    End If
    Fin cm

    msg = "ALMACÉN (" & MesTexto() & ")" & vbCrLf & _
          "   Insumos en el catálogo: " & anCat & vbCrLf & _
          "   Movimientos válidos: " & anMov
    If anMal > 0 Then msg = msg & vbCrLf & "   Filas incompletas o con código inexistente: " & anMal
    msg = msg & vbCrLf & "   Filas con observación (columna VERIFICACIÓN): " & nObs
    If nNeg > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Insumos con SALDO NEGATIVO (sale más de lo que hay):" & neg
        If nNeg > 12 Then msg = msg & vbCrLf & "   ... y " & (nNeg - 12) & " más (vea la columna SALDO DESPUÉS)."
    Else
        msg = msg & vbCrLf & "   Ningún insumo queda con saldo negativo."
    End If
    msg = msg & vbCrLf & vbCrLf & "EQUIPO MECÁNICO" & vbCrLf & "   Equipos: " & enEq & vbCrLf & "   Partes (equipo-día): " & nG
    If pnMal > 0 Then msg = msg & vbCrLf & "   Filas de PARTES sin fecha o con equipo no registrado: " & pnMal
    If nDif > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Días con diferencia entre horómetro/horario y horas a valorizar:" & dif
    End If
    If nExc > 0 Then msg = msg & vbCrLf & vbCrLf & "Equipos que superan lo contratado en la O/S:" & exc
    Mensaje D_MSG, "Procesado el " & Format(Now, "dd/mm/yyyy hh:mm") & ": " & anMov & " movimientos (" & nObs & _
            " con observación, " & nNeg & " insumos con saldo negativo); " & nG & " partes (" & nDif & " por revisar)."
    Procesar = msg
End Function
