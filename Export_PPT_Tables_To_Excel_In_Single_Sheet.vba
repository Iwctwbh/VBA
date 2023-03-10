'For Grape
Sub Export_PPT_Tables_To_Excel_In_Single_Sheet()
    Dim pptSlide As Slide '创建 PowerPoint 幻灯片对象
    Dim pptShape As Shape '创建 PowerPoint 形状对象
    Dim excelApp As Object '创建 Excel 应用程序对象
    Dim excelWorkbook As Object '创建 Excel 工作簿对象
    Dim excelWorksheet As Object '创建 Excel 工作表对象
    Dim row As Long
    Dim rowTemp As Long
    Dim col As Long
    Dim flag As Boolean

    row = 0

    '创建一个新的 Excel 应用程序
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True '可以看到导出的结果

    '创建一个新的 Excel 工作簿
    Set excelWorkbook = excelApp.Workbooks.Add
    '获取第一个 Excel 工作表
    Set excelWorksheet = excelWorkbook.Sheets(1)

    '循环遍历每个幻灯片
    For Each pptSlide In ActivePresentation.Slides
        '循环遍历每个窗格
        For Each pptShape In pptSlide.Shapes
            If pptShape.HasTable Then
                rowTemp = 1

                excelWorksheet.Cells(rowTemp + row, 1).Value = "Slide" & CStr(pptSlide.SlideIndex)

                '判断原本表格中是否有换行符，有则保留并合并相应单元格
                flag = False
                If InStr(1, pptShape.Table.Cell(rowTemp, 1).Shape.TextFrame.TextRange.Text, Chr(11)) > 0 Then
                    flag = True
                End If
                If flag Then
                    For colTemp = 1 To pptShape.Table.Columns.Count
                        excelWorksheet.Range(excelWorksheet.Cells(rowTemp + row, colTemp + 1), excelWorksheet.Cells(rowTemp + row + 1, colTemp + 1)).Merge
                    Next colTemp
                    row = row + 1
                End If

                '将表格数据复制到 Excel 工作表中
                For rowTemp = 1 To pptShape.Table.Rows.Count
                    For colTemp = 1 To pptShape.Table.Columns.Count
                        excelWorksheet.Cells(rowTemp + row, colTemp + 1).MergeArea.Value = pptShape.Table.Cell(rowTemp, colTemp).Shape.TextFrame.TextRange.Text
                        '颜色
                        excelWorksheet.Cells(rowTemp + row, colTemp + 1).MergeArea.Interior.Color = pptShape.Table.Cell(rowTemp, colTemp).Shape.Fill.ForeColor.RGB
                    Next colTemp
                Next rowTemp
                '累加每个Table的行数，末尾为空行数
                row = row + pptShape.Table.Rows.Count + 2
            End If
        Next pptShape
    Next pptSlide
    '列宽自适应
    excelWorksheet.Columns.AutoFit
    excelWorksheet.Rows.AutoFit
End Sub
