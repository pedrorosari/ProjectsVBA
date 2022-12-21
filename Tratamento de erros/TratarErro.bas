Attribute VB_Name = "Módulo1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
Dim range1, cell As Range
Dim aba_compilado, aba_notas As String
Dim encontrado As Integer

aba_compilado = "Compilado"
aba_notas = "Notas Alunos"


Sheets(aba_compilado).Activate
Set range1 = Range("A1", Range("A1").End(xlDown))

For Each cell In range1
    encontrado = 1
    Sheets(aba_notas).Activate
    
    On Error GoTo tratamento
    Cells.Find(What:=cell.Offset(0, 1).Value, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    If encontrado = 1 Then
        cell.Offset(0, 2).Value = Selection.Offset(0, 1).Value
    End If
    

Next


Exit Sub

tratamento:
    encontrado = 0
    Resume Next


End Sub
