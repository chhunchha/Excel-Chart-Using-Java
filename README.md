# Excel-Chart-Using-Java

## Usecase : Insert data and prepare charts - pie, bar or other as per requirement.
Apache POI allows only scatter and line chart as of now.

## Solution:
- Created a template file which has macro as below. Use this workbook and worksheet as base for creating new workbook/worksheet.
- Macro run when worksheet activates and it changes chart type to required chart type based on value at cell Z1.

        Private Sub Workbook_SheetActivate(ByVal Sh As Object)
            On Error GoTo ErrHandler:
            
            ChartType = Cells(1, "Z").Value
            If ChartType = "" Then
                Exit Sub
            End If
            
            ActiveSheet.ChartObjects("Chart 1").Activate
            ActiveChart.PlotArea.Select
    
            If ChartType = "pie" Then
                ActiveChart.ChartType = xlPie
            ElseIf ChartType = "column" Or ChartType = "bar" Then
                ActiveChart.ChartType = xlColumnClustered
            End If
        
            'for enabling data lable and font size change
            ActiveChart.Legend.Select
            Selection.Format.TextFrame2.TextRange.Font.Size = 14
        
            'https://msdn.microsoft.com/en-us/library/office/ff195014.aspx
            ActiveChart.SeriesCollection(1).ApplyDataLabels Type:=xlDataLabelsShowValue
            ActiveChart.SeriesCollection(1).DataLabels.Select
            Selection.Format.TextFrame2.TextRange.Font.Size = 14
        
            'remove chart type so it wont run again
            Cells(1, "Z") = ""
            Cells(1, "A").Select
        ErrHandler:
        End Sub
    
    - Rather than creating new sheet, will clone template sheet, insert data and create line chart as default.
    - put required chart type at Z1.
    
