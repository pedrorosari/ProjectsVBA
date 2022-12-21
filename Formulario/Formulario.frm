VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Formulário de Registro de Compra"
   ClientHeight    =   9630.001
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   6850
   OleObjectBlob   =   "Formulario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
        'cadastrar informações
        Call Cadastrar
        
        
        'limpar formulário
        UserForm1.Hide
        For Each objeto In UserForm1.Controls
            On Error Resume Next
            objeto.Value = ""
        Next
        
End Sub
Sub Cadastrar()
    Dim range1 As Range
        If RefEdit1.Value <> "" Then
            Set range1 = Range(RefEdit1.Value)
        ElseIf Range("A2").Value = "" Then
            Set range1 = Range("A2")
        Else
            Set range1 = Range("A1").End(xlDown).Offset(1, 0)
        End If
        
        range1.Value = UserForm1.ComboBox1.Value
        range1.Offset(0, 1).Value = UserForm1.ListBox1.Value
        range1.Offset(0, 2).Value = UserForm1.ToggleButton1.Value
        range1.Offset(0, 3).Value = UserForm1.CheckBox1.Value
        range1.Offset(0, 4).Value = UserForm1.CheckBox2.Value
        range1.Offset(0, 5).Value = UserForm1.CheckBox3.Value
        range1.Offset(0, 6).Value = UserForm1.CheckBox4.Value

        'cadastro do tipo (produto ou serviço)
        If UserForm1.OptionButton1.Value = True Then
            range1.Offset(0, 7).Value = UserForm1.OptionButton1.Caption
        Else
            range1.Offset(0, 7).Value = UserForm1.OptionButton2.Caption
        End If
                
        'cadastro prazo de pagamento
        If UserForm1.OptionButton3.Value = True Then
            range1.Offset(0, 8).Value = UserForm1.OptionButton3.Caption
        ElseIf UserForm1.OptionButton4.Value = True Then
            range1.Offset(0, 8).Value = UserForm1.OptionButton4.Caption
        Else
            range1.Offset(0, 8).Value = UserForm1.OptionButton5.Caption
        End If
        
        range1.Offset(0, 9).Value = CDbl(UserForm1.TextBox2.Value)
        range1.Offset(0, 9).Style = "Currency"
        range1.Offset(0, 10).Value = UserForm1.TextBox1.Value
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub ToggleButton1_Click()
    Frame2.Visible = ToggleButton1.Value
End Sub

Private Sub UserForm_Initialize()

    With ComboBox1
        .AddItem ("Financeiro")
        .AddItem ("Marketing")
        .AddItem ("Operações")
        .AddItem ("Administrativo")
    End With
    
    UserForm1.Caption = "Registro de Compra"
    
    UserForm1.ToggleButton1.Caption = "Nota Emitida"
    
    Frame2.Visible = False
    Frame2.Caption = "Impostos"
    CheckBox1.Caption = "IR"
    CheckBox2.Caption = "PIS"
    CheckBox3.Caption = "COFINS"
    CheckBox4.Caption = "ISS"
    
    OptionButton1.Caption = "Produto"
    OptionButton2.Caption = "Serviço"
    
    MultiPage1.Pages(0).Caption = "Pagamento"
    MultiPage1.Pages(1).Caption = "Descrição"
    MultiPage1.Pages(2).Caption = "Valor"
    
    
    Frame1.Caption = "Prazo"
    
    OptionButton3.Caption = "Antecipado"
    OptionButton4.Caption = "Na entrega"
    OptionButton5.Caption = "30 dias"
    
    CommandButton1.Caption = "Registrar"
End Sub

