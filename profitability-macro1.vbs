Sub MerchantProfitability()
'
' MerchantProfitability Macro
'

'

    Dim FinData(100)
    Dim PublisherName As String
    Dim ColOrigin As String
    Dim StartofData As Integer
    Dim EndofData As Integer
    Dim Header As Integer
        StartofData = 0
        EndofData = 15
    Dim PublisherFilter As String
    Dim FrontPage As String
    Dim RevPivotTable As String
    'Dim MACPivotTable As String
    Dim MarginalCosts As String
    Dim MarginalCostsPivotTable As String
    Dim ReportPage As String
    Dim Merchants As Integer
    Dim MerchantRanks As String
    
    MerchantRanks = "Merchant_Ranks"
    RevPivotTable = "Revenue Pivot Table"
    MarginalCosts = "Marginal Costs"
    MarginalCostsPivotTable = "Marginal Costs Pivot Table"
    ReportPage = "Report Page"
    FrontPage = "FrontPage"
    PublisherFilter = "Country"
    'RefundsPivotTable = "Refunds Pivot Table"
    ' MACPivotTable = "MAC Pivot Table"
    'APSDetailPivotTable = "APS Detail PivotTable"
    'TotalAPSPivotTable = "Total APS PivotTable"
    
    Merchants = 100
    ColOrigin = "A2"
   ' MsgBox "ComboBox Clicked" & ComboBox1.Value
    
' Make all sheets Visible

    Sheets(RevPivotTable).Visible = True
    'Sheets(MACPivotTable).Visible = True
    Sheets(MarginalCosts).Visible = True
    'Sheets(RefundsPivotTable).Visible = True

    ' Start Changing the Pivot Table
    
    For x = 1 To Merchants
    ' For x = 60 To Merchants
           PublisherName = Application.WorksheetFunction.VLookup(x, Range(MerchantRanks), 2, False)
        
           'MsgBox " Publisher is" & PublisherName
           
           Sheets(RevPivotTable).Select
           ActiveSheet.PivotTables(RevPivotTable).PivotFields(PublisherFilter).ClearAllFilters
           ActiveSheet.PivotTables(RevPivotTable).PivotFields(PublisherFilter). _
               CurrentPage = PublisherName
               
           ' End Changing the Pivot Table
           
               'Sheets(MACPivotTable).Select
           'ActiveSheet.PivotTables(MACPivotTable).PivotFields(PublisherFilter).ClearAllFilters
           'ActiveSheet.PivotTables(MACPivotTable).PivotFields(PublisherFilter). _
               CurrentPage = PublisherName
           
           ' Start Changing the Pivot Table
           
           Sheets(MarginalCosts).Select
           ActiveSheet.PivotTables(MarginalCostsPivotTable).PivotFields(PublisherFilter).ClearAllFilters
           ActiveSheet.PivotTables(MarginalCostsPivotTable).PivotFields(PublisherFilter). _
               CurrentPage = PublisherName
            
          ' Sheets(APSDetailPivotTable).Select
          ' ActiveSheet.PivotTables(APSDetailPivotTable).PivotFields(PublisherFilter).ClearAllFilters
           'ActiveSheet.PivotTables(APSDetailPivotTable).PivotFields(PublisherFilter). _
            '   CurrentPage = PublisherName
               
          ' Sheets(TotalAPSPivotTable).Select
          ' ActiveSheet.PivotTables(TotalAPSPivotTable).PivotFields(PublisherFilter).ClearAllFilters
           'ActiveSheet.PivotTables(TotalAPSPivotTable).PivotFields(PublisherFilter). _
               CurrentPage = PublisherName
        
           ' End Changing the Pivot Table
           ' Start Changing the Pivot Table
           
          ' Sheets(RefundsPivotTable).Select
          ' ActiveSheet.PivotTables(RefundsPivotTable).PivotFields(PublisherFilter).ClearAllFilters
         '   On Error GoTo ErrHandler:
         '  ActiveSheet.PivotTables(RefundsPivotTable).PivotFields(PublisherFilter). _
               CurrentPage = PublisherName
               
         
         



               
               
               
           ' End Changing the Pivot Table
            Sheets(ReportPage).Select
            
            Range(ColOrigin).Select
             
            For i = StartofData To EndofData
                'MsgBox "Into of Loop"
                FinData(i - StartofData) = ActiveCell.Value
                ActiveCell.Offset(0, 1).Select
            Next i

            Range(ColOrigin).Select
            ActiveCell.Offset(x + 10, 0).Select
            

            For i = StartofData To EndofData
                'MsgBox "Into of Loop"
                ActiveCell.Value = FinData(i - StartofData)
                ActiveCell.Offset(0, 1).Select
            Next i
            
     Next x
    
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
    



