Option Explicit
Public ValorAnterior As Variant

' Guarda el valor antes de que cambie
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column <= 50 And Target.Row <= 200 Then
        ValorAnterior = Target.Value
    End If
End Sub

' Detecta cambios y ejecuta las dos funciones
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim HojaTrail As Worksheet
    Dim RangoTrail As Range
    Dim NuevaFila As Integer

    ' Evita bucles infinitos
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Desprotege la hoja antes de modificarla
    Me.Unprotect "Takeda--Ing!25"

    ' -------------------- HISTORIAL DE CAMBIOS --------------------
    If Target.Column <= 100 And Target.Row <= 250 Then
        Set HojaTrail = ThisWorkbook.Sheets("Trail")
        Set RangoTrail = HojaTrail.Range("A2").CurrentRegion
        NuevaFila = RangoTrail.Rows.Count + 1

        With HojaTrail
            .Cells(NuevaFila, 1).Value = Application.UserName
            .Cells(NuevaFila, 2).Value = Target.Address
            .Cells(NuevaFila, 3).Value = "INFORME"
            .Cells(NuevaFila, 4).Value = Date & "  " & Time
            .Cells(NuevaFila, 5).Value = ValorAnterior
            .Cells(NuevaFila, 6).Value = Target.Value
        End With
    End If

    ' ------------------ OCULTAR FILAS SI A4 = "HUMEDAD" ------------------
    If Not Intersect(Target, Me.Range("A4")) Is Nothing Then
        If UCase(Me.Range("A4").Value) = "HUMEDAD" Then
            Me.Rows("39:45").Hidden = True  ' Oculta filas
        Else
            Me.Rows("39:45").Hidden = False ' Muestra filas
        End If
    End If

    ' Protege la hoja nuevamente
    Me.Protect "Takeda--Ing!25"

    ' Reactiva eventos y actualización de pantalla
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

