Attribute VB_Name = "M�dulo1"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("Tabela Din�mica").Select
    Range("C13").Select
    ActiveSheet.PivotTables("Tabela din�mica1").PivotCache.Refresh
    Sheets("Impress�o").Select
    Range("A1:B24").Select
    Range("B24").Activate
    ActiveSheet.PageSetup.PrintArea = "$A$1:$B$24"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\Johnny\Desktop\Cursos Udemy\VBA para Universit�rios\Planilhas do Curso\Planilhas Prontas\Se��o 8\Template Ferramenta Eventos e Impress�o.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    Sheets("Impress�o").Select
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveWorkbook.RefreshAll
End Sub
