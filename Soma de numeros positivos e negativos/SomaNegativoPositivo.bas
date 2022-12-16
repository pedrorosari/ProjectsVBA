Attribute VB_Name = "SomaNegativoPositivo"
Sub Calculo()
Application.ScreenUpdating = False
Range("B11").Select
negativos = 0
positvos = 0
Total = 0

Do Until ActiveCell.Value = ""
Total = Total + ActiveCell.Value
If ActiveCell.Value < 0 Then
negativos = negativos + ActiveCell.Value
ActiveCell.Offset(1, 0).Select

Else
positvos = positvos + ActiveCell.Value
ActiveCell.Offset(1, 0).Select
End If

Loop

Range("E15").Value = positvos
Range("E16").Value = negativos
Range("E17").Value = Total
Application.ScreenUpdating = True
End Sub
