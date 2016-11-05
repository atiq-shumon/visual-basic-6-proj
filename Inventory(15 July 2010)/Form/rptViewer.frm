VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rptViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "rptViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   6765
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7725
      lastProp        =   500
      _cx             =   13626
      _cy             =   11933
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report1   As New CrystalReport1
Dim Report2   As New CrystalReport2
Dim Report3   As New CrystalReport3
Dim Report4   As New CrystalReport4
Dim Report5   As New CrystalReport5
Dim Report6   As New CrystalReport6
Dim Report7   As New CrystalReport7
Dim Report8   As New CrystalReport8
Dim Report9   As New CrystalReport9
Dim Report10   As New CrystalReport10
Dim Report11   As New CrystalReport11
Dim Report12   As New CrystalReport12


Private Sub Form_Load()
On Error GoTo err_desc
Screen.MousePointer = vbHourglass
Screen.MousePointer = vbDefault
CRViewer1.EnableAnimationCtrl = True
CRViewer1.Zoom (100)

Dim Conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim RS As New ADODB.Recordset
Dim Param1 As New Parameter
Dim Param2 As New Parameter
Dim Param3 As New Parameter
Dim Param4 As New Parameter
Dim Param5 As New Parameter
Dim Param6 As New Parameter
Dim Param7 As New Parameter
Dim Param8 As New Parameter

Select Case rptmode
    Case 1
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, optionMode)
        cmd.Parameters.Append Param1 'combo
        
               
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 5, CategoryCode)
        cmd.Parameters.Append Param2 'category
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 5, Get_Code(frmRptItemStatements.Combo1))
        cmd.Parameters.Append Param3 'Type
       
        
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 5, Get_Code(frmRptItemStatements.Combo2))
        cmd.Parameters.Append Param4 'Group
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_item_info (?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report1.Database.SetDataSource RS
        
        ''''''''''for viewing'''''''''''''''
        Report1.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report1
        
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
        
  Case 2
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 5, CatCode)
        cmd.Parameters.Append Param2
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Opening_Banalce_info(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report2.Database.SetDataSource RS
        Report2.Text7.SetText "Category : " + CatName
        ''''''''''for viewing'''''''''''''''
        Report2.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report2
        
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
    Case 3
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 2)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 2, CatCode)
        cmd.Parameters.Append Param2
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Opening_Banalce_info(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report2.Database.SetDataSource RS
        Report2.Text7.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report2.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report2
        
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
    Case 4
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param3
        
        Set Param4 = cmd.CreateParameter("param4", adChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param4
        
        Set Param5 = cmd.CreateParameter("param5", adChar, adParamInput, 10, SuppCode)
        cmd.Parameters.Append Param5
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_PurchaseStatement(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report3.Database.SetDataSource RS
        Report3.Text9.SetText "PURCHASE STATEMENTS FOR THE PERIOD OF " + StDate + " TO " + EdDate
        Report3.Text12.SetText "Supplier : " + SuppName
        Report3.Text10.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report3.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report3
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
    Case 5
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param3
        
        Set Param4 = cmd.CreateParameter("param4", adChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param4
        
        Set Param5 = cmd.CreateParameter("param5", adChar, adParamInput, 10, CustCode)
        cmd.Parameters.Append Param5
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_IssueStatement(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report4.Database.SetDataSource RS
        Report4.Text8.SetText "ISSUE STATEMENTS FOR THE PERIOD OF " + StDate + " TO " + EdDate
        Report4.Text12.SetText "Issue Type : " + CustName
        Report4.Text14.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report4.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report4
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
    Case 6
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param3
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_ExpireDateStatement(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report5.Database.SetDataSource RS
        Report5.Text8.SetText "EXPIRED DATE STATEMENTS AS ON " + StDate
        Report5.Text14.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report5.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report5
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
    Case 7
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        If Len(frmRptStockStatements.Combo1) <> 0 And (Len(frmRptStockStatements.Combo2) = 0 And Len(frmRptStockStatements.Combo3) = 0) Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
                cmd.Parameters.Append Param1
        ElseIf (Len(frmRptStockStatements.Combo1) <> 0 And Len(frmRptStockStatements.Combo2) <> 0) And Len(frmRptStockStatements.Combo3) = 0 Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 2)
                cmd.Parameters.Append Param1
                CatName = CatName + " " + "   Type : " + Get_Description(frmRptStockStatements.Combo2)
         ElseIf Len(frmRptStockStatements.Combo1) <> 0 And Len(frmRptStockStatements.Combo2) <> 0 And Len(frmRptStockStatements.Combo3) <> 0 Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 3)
                cmd.Parameters.Append Param1
                CatName = CatName + " " + "  Type : " + Get_Description(frmRptStockStatements.Combo2) + "  Group : " + Get_Description(frmRptStockStatements.Combo3)
        End If
        
        Set Param2 = cmd.CreateParameter("param2", adSmallInt, adParamInput, 3, stockOrValue)
        cmd.Parameters.Append Param2
      
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param3
        
         Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Get_Code(frmRptStockStatements.Combo2.Text))
        cmd.Parameters.Append Param4
     
        Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, Get_Code(frmRptStockStatements.Combo3.Text))
        cmd.Parameters.Append Param5
     
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param6
        
        Set Param7 = cmd.CreateParameter("param7", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param7
        
        cmd.Properties("PLSQLRSet") = True
        Screen.MousePointer = vbHourglass
        cmd.CommandText = "{CALL Rpt_stock_info(?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report6.Database.SetDataSource RS
        If stockOrValue = 1 Then
          Report6.Text8.SetText "STOCK STATEMENTS FOR THE PERIOD " + StDate + " TO " + EdDate
          Report6.ReportFooterSection1.Suppress = True
          Report6.GroupFooterSection1.Suppress = True
        ElseIf stockOrValue = 2 Then
          Report6.Text8.SetText "STOCK VALUE STATEMENTS FOR THE PERIOD " + StDate + " TO " + EdDate
          Report6.ReportFooterSection1.Suppress = False
          Report6.GroupFooterSection1.Suppress = False
        End If
        
        Report6.Text14.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report6.DiscardSavedData
        
        CRViewer1.ReportSource = Report6
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
    Case 8
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param3
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param4
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_AdjustmentStatement(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report7.Database.SetDataSource RS
        Report7.Text8.SetText "ADJUSTMENT STATEMENTS FOR THE PERIOD " + StDate + " TO " + EdDate
        Report7.Text14.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report7.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report7
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
        
     Case 9
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param3
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param4
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL  Rpt_Item_Ledger(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report8.Database.SetDataSource RS
        Report8.Text9.SetText "ITEM LEDGER FOR THE PERIOD " + StDate + " TO " + EdDate
        Report8.Text12.SetText "Item : " + CatName
'        ''''''''''for viewing''''''''''''
        Report8.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report8
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
    Case 10
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
        If CatCode = "All" Then
          Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
          cmd.Parameters.Append Param1 'combo
        Else
          Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 2)
          cmd.Parameters.Append Param1 'combo
        End If
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 3, CatCode)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 6, "1")
        cmd.Parameters.Append Param3
         
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 6, "1")
        cmd.Parameters.Append Param4
 
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_minimum_Banalce_info(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report9.Database.SetDataSource RS
        Report9.Text7.SetText (CatName)
        ''''''''''for viewing''''''''''''
        Report9.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report9
        
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
        
     Case 11
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        
       
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
          cmd.Parameters.Append Param1 'combo
       
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, frmIssue.txtfields(0).Text)
        cmd.Parameters.Append Param2
        
           
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Issue_info (?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report10.Database.SetDataSource RS
'        Report9.Text7.SetText (CatName)
        ''''''''''for viewing''''''''''''
        Report10.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report10
        
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
        
   Case 12
        Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        If Len(frmRptClosingStockValuation.Combo1) <> 0 And (Len(frmRptClosingStockValuation.Combo2) = 0 And Len(frmRptClosingStockValuation.Combo3) = 0) Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
                cmd.Parameters.Append Param1
        ElseIf (Len(frmRptClosingStockValuation.Combo1) <> 0 And Len(frmRptClosingStockValuation.Combo2) <> 0) And Len(frmRptClosingStockValuation.Combo3) = 0 Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 2)
                cmd.Parameters.Append Param1
                CatName = CatName + " " + "   Type : " + Get_Description(frmRptClosingStockValuation.Combo2)
         ElseIf Len(frmRptClosingStockValuation.Combo1) <> 0 And Len(frmRptClosingStockValuation.Combo2) <> 0 And Len(frmRptClosingStockValuation.Combo3) <> 0 Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 3)
                cmd.Parameters.Append Param1
                CatName = CatName + " " + "  Type : " + Get_Description(frmRptClosingStockValuation.Combo2) + "  Group : " + Get_Description(frmRptClosingStockValuation.Combo3)
        End If
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Get_Code(frmRptClosingStockValuation.Combo2.Text))
        cmd.Parameters.Append Param3
     
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Get_Code(frmRptClosingStockValuation.Combo3.Text))
        cmd.Parameters.Append Param4
     
        
        Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param5

        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param6

        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_stock_valuation(?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report11.Database.SetDataSource RS
        Report11.Text16.SetText "STOCK VALUATION STATEMENTS FOR THE PERIOD " + StDate + " TO " + EdDate
        Report11.Text16.SetText "Category : " + CatName
        Report11.Text4.SetText (EdDate)
'        ''''''''''for viewing''''''''''''
        CRViewer1.EnableExportButton = True
        Report11.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report11
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault
        
 Case 14
      Set cmd.ActiveConnection = objmyCon
        cmd.CommandType = adCmdText
        If Len(frmRptStockStatements.Combo1) <> 0 And (Len(frmRptStockStatements.Combo2) = 0 And Len(frmRptStockStatements.Combo3) = 0) Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 1)
                cmd.Parameters.Append Param1
        ElseIf (Len(frmRptStockStatements.Combo1) <> 0 And Len(frmRptStockStatements.Combo2) <> 0) And Len(frmRptStockStatements.Combo3) = 0 Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 2)
                cmd.Parameters.Append Param1
                CatName = CatName + " " + "   Type : " + Get_Description(frmRptStockStatements.Combo2)
         ElseIf Len(frmRptStockStatements.Combo1) <> 0 And Len(frmRptStockStatements.Combo2) <> 0 And Len(frmRptStockStatements.Combo3) <> 0 Then
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 1, 3)
                cmd.Parameters.Append Param1
                CatName = CatName + " " + "  Type : " + Get_Description(frmRptStockStatements.Combo2) + "  Group : " + Get_Description(frmRptStockStatements.Combo3)
        End If
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, CatCode)
        cmd.Parameters.Append Param2
        
         Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Get_Code(frmRptStockStatements.Combo2.Text))
        cmd.Parameters.Append Param3
     
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Get_Code(frmRptStockStatements.Combo3.Text))
        cmd.Parameters.Append Param4
     
        
        Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, StDate)
        cmd.Parameters.Append Param5
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, EdDate)
        cmd.Parameters.Append Param6
         Screen.MousePointer = vbHourglass
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_YEARLY_REQUISITION(?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report12.Database.SetDataSource RS
        Report12.Text6.SetText "STOCK STATEMENTS FOR THE PERIOD: " + StDate + " TO  " + EdDate + " AND REQUIREMENT FOR THE PERIOD FROM ...........TO .............."
         Report12.Text9.SetText (StDate)
        Report12.Text10.SetText "Category : " + CatName
        ''''''''''for viewing''''''''''''
        Report12.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report12
        CRViewer1.ViewReport
       
        Screen.MousePointer = vbDefault

   End Select
   Exit Sub
err_desc:
        MsgBox Err.Description, vbInformation, cmp
End Sub
Private Sub Form_Resize()
   CRViewer1.Top = 0
   CRViewer1.Left = 0
   CRViewer1.Zoom 100
   CRViewer1.Height = ScaleHeight
   CRViewer1.Width = ScaleWidth
End Sub
