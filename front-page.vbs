Dim bFlag As Boolean
Private Sub ComboBox1_Change()
    If bFlag = False Then Exit Sub
    'MsgBox "ComboBox Changed"
End Sub

Private Sub ComboBox1_Click()

    Dim PublisherName As String
    Dim PublisherFilter As String
    Dim FrontPage As String
    Dim RevPivotTable As String
    'Dim MACPivotTable As String
    
    RevPivotTable = "Revenue Pivot Table"
    MACPivotTable = "MAC Pivot Table"
    MarginalCosts = "Marginal Costs"
    MarginalCostsPivotTable = "Marginal Costs Pivot Table"
    RefundsPivotTable = "Refunds Pivot Table"
    
    
    FrontPage = "FrontPage"
    PublisherFilter = "Publisher"
   
    If bFlag = False Then Exit Sub
   ' MsgBox "ComboBox Clicked" & ComboBox1.Value
    
' Make all sheets Visible

    Sheets(RevPivotTable).Visible = True
    'Sheets(MACPivotTable).Visible = True
    Sheets(MarginalCosts).Visible = True
    Sheets(RefundsPivotTable).Visible = True

    ' Start Changing the Pivot Table
    PublisherName = ComboBox1.Value
 
    'MsgBox " Publisher is" & PublisherName
    
    Sheets(RevPivotTable).Select
    ActiveSheet.PivotTables(RevPivotTable).PivotFields(PublisherFilter).ClearAllFilters
    ActiveSheet.PivotTables(RevPivotTable).PivotFields(PublisherFilter). _
        CurrentPage = PublisherName
        
    ' End Changing the Pivot Table
    
    '    Sheets(MACPivotTable).Select
   ' ActiveSheet.PivotTables(MACPivotTable).PivotFields(PublisherFilter).ClearAllFilters
    'ActiveSheet.PivotTables(MACPivotTable).PivotFields(PublisherFilter). _
    '    CurrentPage = PublisherName
    
    ' Start Changing the Pivot Table
    
    Sheets(MarginalCosts).Select
    ActiveSheet.PivotTables(MarginalCostsPivotTable).PivotFields(PublisherFilter).ClearAllFilters
    ActiveSheet.PivotTables(MarginalCostsPivotTable).PivotFields(PublisherFilter). _
        CurrentPage = PublisherName
        
    ' End Changing the Pivot Table
    ' Start Changing the Pivot Table
    
    'Sheets(RefundsPivotTable).Select
    'ActiveSheet.PivotTables(RefundsPivotTable).PivotFields(PublisherFilter).ClearAllFilters
  ' On Error GoTo ErrHandler:
    'ActiveSheet.PivotTables(RefundsPivotTable).PivotFields(PublisherFilter). _
     '   CurrentPage = PublisherName
         
    ' End Changing the Pivot Table
    
    
    'Sheets("Revenue Pivot Table").Visible = False
    'Sheets("MAC Pivot Table").Visible = False
    'Sheets(MTMOPivotTable).Visible = False
    'Sheets(RefundsPivotTable).Visible = False
    
    Sheets(FrontPage).Select
    Exit Sub
    
ErrHandler:
    ActiveSheet.PivotTables(RefundsPivotTable).PivotFields(PublisherFilter). _
    CurrentPage = "None"
    ' go back to the line following the error
    Resume Next
End Sub



Private Sub ComboBox1_GotFocus()
bFlag = True
End Sub

Private Sub ComboBox1_LostFocus()
   bFlag = False
End Sub
