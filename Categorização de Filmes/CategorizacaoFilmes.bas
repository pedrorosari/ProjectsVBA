Attribute VB_Name = "CategorizacaoFilmes"
Sub DoLoop_Ex3()

'Dim QtdMinuto As Integer
'Dim DuracaoFilme As String

'Application.ScreenUpdating = False

'Organiza��o planilhas de suporte
'Para que todas estejam na ultima linha em branco, antes de inserir dados

Worksheets("Longo").Activate
Range("B8").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select

Worksheets("M�dio").Activate
Range("B8").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select

Worksheets("Curto").Activate
Range("B8").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select


'Volta a planilha original para avalia��o
Worksheets("Ex3").Activate
Range("B11").Select

'Avalia��o de linha a linha, similar a um PROCV
Do Until ActiveCell.Value = ""
QtdMinuto = ActiveCell.Offset(0, 4).Value
If QtdMinuto < 100 Then
DuracaoFilme = "Curto"
ElseIf QtdMinuto < 130 Then
DuracaoFilme = "M�dio"
Else
DuracaoFilme = "Longo"
End If

'Atribui os valores de Longo, M�dio, Curto para a quinta celula a direita da selecionada
ActiveCell.Offset(0, 5).Value = DuracaoFilme

'Copia a linha toda para a devida aba
Range(ActiveCell, ActiveCell.End(xlToRight)).Copy
Worksheets(DuracaoFilme).Activate
ActiveCell.PasteSpecial
ActiveCell.Offset(1, 0).Select

'Volta a planilha original e desce uma linha
Worksheets("Ex3").Activate
ActiveCell.Offset(1, 0).Select

'O loop faz com que o c�digo retorne a linha 'Do Until'
Loop
'Application.ScreenUpdating = True
End Sub


