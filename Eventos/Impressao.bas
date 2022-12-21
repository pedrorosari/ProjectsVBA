Attribute VB_Name = "Módulo1"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("Tabela Dinâmica").Select
    Range("C13").Select
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotCache.Refresh
    Sheets("Impressão").Select
    Range("A1:B24").Select
    Range("B24").Activate
    ActiveSheet.PageSetup.PrintArea = "$A$1:$B$24"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\Johnny\Desktop\Cursos Udemy\VBA para Universitários\Planilhas do Curso\Planilhas Prontas\Seção 8\Template Ferramenta Eventos e Impressão.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    Sheets("Impressão").Select
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveWorkbook.RefreshAll
End Sub
