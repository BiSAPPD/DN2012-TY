Attribute VB_Name = "unlinkPivot"


Sub Unlink_Pivot()

    Dim pvtTable As PivotTable
    Dim r As Long, c As Long
    
    NF = ActiveWorkbook.Name
    act_month = InputBox("Month")
         
    myLib.VBA_Start
         
    Set NewBook = Workbooks.Add
    With NewBook
        .Title = "DN_2016-2012_actl2016" & act_month
        .Subject = "Sales"
       ' .SaveAs Filename:="Allsales.xls"
    End With
    NWB = ActiveWorkbook.Name
    
    
    Workbooks(NF).Activate
    sts_pvt = 0
    num_sh = 1
    For Each sh In ActiveWorkbook.Worksheets
    sh.Activate
    nm_sh = sh.Name
       For Each pvtTable In ActiveSheet.PivotTables
    
    
       If Not pvtTable Is Nothing Then
        sts_pvt = 1
        pvtTable.TableRange2.Copy
        Workbooks(NWB).Activate
        Set Sh_new = Worksheets.Add()
        Sh_new.Name = sh.Name & "_"
        ActiveSheet.Range("a1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ActiveSheet.Range("a1").PasteSpecial Paste:=xlPasteColumnWidths


        
            For c = 1 To Selection.Columns.Count
                For r = 1 To Selection.Rows.Count
    
                    Selection.Cells(r, c).Interior.Color = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Interior.Color
    
                    Selection.Cells(r, c).Font.Name = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Font.Name
                    Selection.Cells(r, c).Font.Size = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Font.Size
                    Selection.Cells(r, c).Font.Color = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Font.Color
                    Selection.Cells(r, c).Font.Bold = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Font.Bold
                    Selection.Cells(r, c).Font.Italic = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Font.Italic
    
                    Selection.Cells(r, c).Borders(xlEdgeLeft).Color = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeLeft).Color
                    Selection.Cells(r, c).Borders(xlEdgeRight).Color = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeRight).Color
                    Selection.Cells(r, c).Borders(xlEdgeTop).Color = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeTop).Color
                    Selection.Cells(r, c).Borders(xlEdgeBottom).Color = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeBottom).Color
    
                    Selection.Cells(r, c).Borders(xlEdgeLeft).LineStyle = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeLeft).LineStyle
                    Selection.Cells(r, c).Borders(xlEdgeRight).LineStyle = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeRight).LineStyle
                    Selection.Cells(r, c).Borders(xlEdgeTop).LineStyle = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeTop).LineStyle
                    Selection.Cells(r, c).Borders(xlEdgeBottom).LineStyle = pvtTable.TableRange2.Cells(r, c).DisplayFormat.Borders(xlEdgeBottom).LineStyle
    
                Next r
            Next c
        Else
        
              
        End If
        
    
    
    Next pvtTable
    
        If sts_pvt = 0 Then
        Workbooks(NF).Activate
        Sheets(nm_sh).Select
        Sheets(nm_sh).Copy After:=Workbooks(NWB).Sheets(1)
        
        For Each cell In ActiveSheet.UsedRange
        If Not IsError(cell) Then
        c_cells = cell.Value
        End If
        Next cell
                
        End If
        
        Workbooks(NWB).Activate
        num_p = ActiveWorkbook.Worksheets.Count
        ActiveSheet.Move After:=Sheets(num_p)
        ActiveWindow.DisplayGridlines = False
        num_sh = num_sh + 1

    Next sh
    
    Workbooks(NWB).Activate
    Set wb = ActiveWorkbook
        WorkbookLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
        If IsArray(WorkbookLinks) Then
            For i = LBound(WorkbookLinks) To UBound(WorkbookLinks)
                wb.BreakLink _
                        Name:=WorkbookLinks(i), _
                        Type:=xlLinkTypeExcelLinks
            Next i
        End If
       
    myLib.VBA_End
    
End Sub

