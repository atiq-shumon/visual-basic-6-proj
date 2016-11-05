VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form CRViewer1 
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   5565
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   5295
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7035
      lastProp        =   500
      _cx             =   12409
      _cy             =   9340
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "CRViewer1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report1   As New CrystalReport1
Dim Report2   As New CrystalReport2
Dim Report3   As New CrystalReport3
Dim Report4   As New CrystalReport4
Dim Report5   As New CrystalReport13
Dim Report13   As New CrystalReport5
Dim Report6   As New CrystalReport6
Dim Report7   As New CrystalReport7
Dim Report8   As New CrystalReport8
Dim Report9   As New CrystalReport9
Dim Report12   As New CrystalReport12
Dim Report10   As New CrystalReport10
Dim Report14   As New CrystalReport14
Dim Report15  As New CrystalReport15
Dim Report16  As New CrystalReport16
Dim Report17  As New CrystalReport17
Dim Report18  As New CrystalReport18
Dim Report19  As New CrystalReport19
Dim Report20  As New CrystalReport20
Dim Report21  As New CrystalReport21
Dim Report22  As New CrystalReport22



Private Sub CRViewer1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo err_desc
    CRViewer91.Zoom 80
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
Select Case rptMode

    Case 1
            Conn.Open strcn.Connection_String
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
    
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL RptAcct}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report1.Database.SetDataSource RS
                
            Report1.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report1
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
     Case 100
            Conn.Open strcn.Connection_String
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
    
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL  Rpt_Opening_Balance}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report22.Database.SetDataSource RS
                
            Report22.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report22
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
  
    Case 2
            If Form9.strCase = "Vou_No" Then

                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                 Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, 2)
                  cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(Form9.txtVOU_NO.Text))
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 15, Trim(Form9.txtVouType.Text))
                cmd.Parameters.Append Param3
              
              
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Trim(Form9.dtst_dt.Value))
                cmd.Parameters.Append Param4
                
                Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 15, Trim(Form9.dted_dt.Value))
                cmd.Parameters.Append Param5
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptVou_all(?,?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report4.Database.SetDataSource RS
                Report4.Text5.SetText (Form9.cboReportType.Text)
                Report4.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report4
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
            ElseIf Form9.strCase = "Vou_Date" Then

                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                              
                Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, 3)
                  cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(Form9.txtVOU_NO.Text))
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 15, Trim(Form9.txtVouType.Text))
                cmd.Parameters.Append Param3
              
              
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Trim(Form9.dtst_dt.Value))
                cmd.Parameters.Append Param4
                
                Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 15, Trim(Form9.dted_dt.Value))
                cmd.Parameters.Append Param5
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptVou_all(?,?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report4.Database.SetDataSource RS
                Report4.Text5.SetText (Form9.cboReportType.Text)
                Report4.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report4
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                                
            ElseIf Form9.strCase = "All" Then

                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                                          
                                          
                If Form9.cboReportType.Text = "All Transaction" Then
                    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, 4)
                  cmd.Parameters.Append Param1
                Else
                   Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, 1)
                   cmd.Parameters.Append Param1
               End If
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(Form9.txtVOU_NO.Text))
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 15, Trim(Form9.txtVouType.Text))
                cmd.Parameters.Append Param3
              
              
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Trim(Form9.dtst_dt.Value))
                cmd.Parameters.Append Param4
                
                Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 15, Trim(Form9.dted_dt.Value))
                cmd.Parameters.Append Param5
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptVou_all(?,?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report4.Database.SetDataSource RS
                Report4.Text5.SetText (Form9.cboReportType.Text)
                Report4.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report4
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
            End If
    Case 3

                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
              
                Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Trim(Form6.txtVOU_NO.Text))
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, Trim(Form6.cboVou_Type))
                cmd.Parameters.Append Param2
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptVou(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report2.Database.SetDataSource RS
                    
                Report2.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report2
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
    Case 4
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
              
                Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 20, Form10.cboStUserCode.Text)
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Format(Form10.dtst_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param2
              
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 8, Form10.dted_dt.Value)
                cmd.Parameters.Append Param3
                
                cmd.Properties("PLSQLRSet") = True
                
                cmd.CommandText = "{CALL RptLedger(?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report5.Text17.SetText (Form10.cboStUserCode)
                Report5.Text19.SetText (Form10.cboStAccName)
                Report5.Text10.SetText (Form10.dtst_dt)
                Report5.Text12.SetText (Form10.dted_dt)
                
                Report5.Database.SetDataSource RS
''                Report5.FormulaFields.Item(5).Text = Chr(34) & Form10.dtst_dt.Value & Chr(34)
'                Report5.FormulaFields.Item(6).Text = Chr(34) & Form10.dted_dt.Value & Chr(34)
                Report5.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report5
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
    Case 5
        Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                 
                Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, Trim(Form11.dtst_dt.Value))
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Trim(Form11.dted_dt.Value))
                cmd.Parameters.Append Param2
                
                Report6.date1.SetText (Form11.dtst_dt.Value)
                Report6.date2.SetText (Form11.dted_dt.Value)
                              
                
                
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL trial_bal(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report6.Database.SetDataSource RS
                    
                Report6.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report6
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault

       
    Case 6
       
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 20, 1)
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form15.cboStUserCode.Text)
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Format(Form15.dtst_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param3
              
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Format(Form15.dted_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param4
                
                Report12.Text18.SetText (Form15.txtacc_head.Text)
                Report12.Text20.SetText (Form15.dtst_dt.Value)
                Report12.Text21.SetText (Form15.dted_dt.Value)
                Report12.Text22.SetText (Form15.cboStAccName.Text)
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptCash_bank(?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report12.Database.SetDataSource RS
                Report12.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report12
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
    Case 8
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
              
                Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 20, Form12.cboUserHead)
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Format(Form12.DTPicker1.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param2
              
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 8, Format(Form12.DTPicker2.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param3
                
                cmd.Properties("PLSQLRSet") = True
                
                
                cmd.CommandText = "{CALL Rpt_Schedule(?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report8.Text17.SetText (Form12.cboUserHead)
                Report8.Text20.SetText (Form12.cboHeadName)
                Report8.Text9.SetText (Form12.DTPicker1.Value)
                Report8.Text10.SetText (Form12.DTPicker2.Value)
'
                Report8.Database.SetDataSource RS
''                Report5.FormulaFields.Item(5).Text = Chr(34) & Form10.dtst_dt.Value & Chr(34)
'                Report5.FormulaFields.Item(6).Text = Chr(34) & Form10.dted_dt.Value & Chr(34)
                Report8.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report8
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault

    Case 11
          Conn.Open strcn.Connection_String
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
    
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL RptAcct}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report7.Database.SetDataSource RS
                
            Report7.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report7
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
    
    Case 13
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
              
                Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 10, Format(Trim(Form13.dtpStart.Value), "DD-MMM-YY"))
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Format(Trim(Form13.dtpEnd), "DD-MMM-YY"))
                cmd.Parameters.Append Param2
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RPT_INCOME_STATEMENT(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report9.Database.SetDataSource RS
                    
                Report9.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report9
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
    Case 14
            
                 Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                 
                Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, Trim(Form16.dtst_dt.Value))
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Trim(Form16.dted_dt.Value))
                cmd.Parameters.Append Param2
                
                Report13.Text13.SetText (Form16.dtst_dt.Value)
                Report13.Text15.SetText (Form16.dted_dt.Value)
                
                
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL rpt_receipt_payment(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report13.Database.SetDataSource RS
                    
                Report13.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report13
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
   Case 15
            Conn.Open strcn.Connection_String
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
    
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_budget}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report10.Database.SetDataSource RS
                
            Report10.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report10
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
    Case 16
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 15, 0)  '''option
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Trim(Form17.dtst_dt.Value))
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, Trim(Form17.dted_dt.Value))
                cmd.Parameters.Append Param3
                
                Report2.Text19.SetText (Form17.dtst_dt.Value)
                Report2.Text20.SetText (Form17.dted_dt.Value)

                
                
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL rpt_asset_schedule(?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report2.Database.SetDataSource RS
                    
                Report2.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report2
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
       Case 17
              Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 20, 1)
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form18.cboStUserCode.Text)
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Format(Form18.dtst_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param3
              
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Format(Form18.dted_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param4
                
                Report14.Text20.SetText (Form18.dtst_dt.Value)
                Report14.Text21.SetText (Form18.dted_dt.Value)
                If Form18.cboStUserCode = "2102" Then
                    Report14.Text18.SetText ("Bank Book")
                ElseIf Form18.cboStUserCode = "2103" Then
                    Report14.Text18.SetText ("Cash Book")
                End If
'                Report14.Text22.SetText (Form18.cboStAccName.Text)
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptCash_book(?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report14.Database.SetDataSource RS
                Report14.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report14
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
            

 Case 18
 
 
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 20, 1)
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form19.cboStUserCode.Text)
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Format(Form19.dtst_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param3
              
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Format(Form19.dted_dt.Value, "DD-MMM-YY"))
                cmd.Parameters.Append Param4
                
                Report15.Text29.SetText (Format(Form19.dtst_dt, "DD/MM/YY"))
                Report15.Text20.SetText (Form19.dtst_dt)
                Report15.Text21.SetText (Form19.dted_dt)
'                If Form18.cboStUserCode = "2102" Then
'                    Report14.Text18.SetText ("Bank Book")
'                ElseIf Form18.cboStUserCode = "2103" Then
'                    Report14.Text18.SetText ("Cash Book")
'                End If
'                Report14.Text22.SetText (Form18.cboStAccName.Text)
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptCash_book_at_a_glnc(?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report15.Database.SetDataSource RS
                Report15.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report15
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
                
      Case 19
      
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 15, 2)  '''option
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Trim(Form20.dtst_dt.Value))
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, Trim(Form20.dted_dt.Value))
                cmd.Parameters.Append Param3
                
                Report16.Text19.SetText (Form20.dtst_dt.Value)
                Report16.Text20.SetText (Form20.dted_dt.Value)

                
                
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL rpt_asset_schedule(?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report16.Database.SetDataSource RS
                    
                Report16.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report16
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
            
 Case 20
 
               Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
              
                Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 10, Format(Trim(Form13.dtpStart.Value), "DD-MMM-YY"))
                cmd.Parameters.Append Param1
                
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Format(Trim(Form13.dtpEnd), "DD-MMM-YY"))
                cmd.Parameters.Append Param2
              
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RPT_INCOME_STATEMENT(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                Report18.Text2.SetText (Form13.dtpStart.Value)
                Report18.Text5.SetText (Form13.dtpEnd.Value)
                Report18.Database.SetDataSource RS
                    
                Report18.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report18
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
   Case 21
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                If Form23.Option1(3) = True Then
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 1)
                      cmd.Parameters.Append Param1
                End If
                If Form23.Option1(2) = True Then  '''----unpaid
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 2)
                      cmd.Parameters.Append Param1
                End If
                If Form23.Option1(1) = True Then  '''----paid
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 3)
                      cmd.Parameters.Append Param1
                End If
                If Form23.Option1(4) = True Then  '''----received
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 4)
                      cmd.Parameters.Append Param1
                End If
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form23.cboStUserCode.Text)
                cmd.Parameters.Append Param2

                Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, Form23.cboStUserCode.Text)
                cmd.Parameters.Append Param3

                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Format(Trim(Form23.dtst_dt.Value), "DD-MMM-YYYY"))
                cmd.Parameters.Append Param4

                Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, Format(Trim(Form23.dted_dt.Value), "DD-MMM-YYYY"))
                cmd.Parameters.Append Param5
                
                 Report3.Text10.SetText (Form23.cboStAccName.Text)
                 Report3.Text20.SetText (Form23.dtst_dt.Value)
                 Report3.Text22.SetText (Form23.dted_dt.Value)
'                Report12.Text22.SetText (Form15.cboStAccName.Text)
'
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptCheque_status(?,?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report3.Database.SetDataSource RS
                    
                Report3.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report3
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
               
               
               
   Case 22
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
                If Form25.Option1(3) = True Then
                    Report17.Text18.SetText ("Income Tax Schedule")
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 1)
                      cmd.Parameters.Append Param1
                End If
                 If Form25.Option1(0) = True Then
                    Report17.Text18.SetText ("VAT Schedule")
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 2)
                      cmd.Parameters.Append Param1
                End If
                
                If Form25.Option1(2) = True Then
                    Report17.Text18.SetText ("Security Schedule")
                    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 3)
                      cmd.Parameters.Append Param1
                End If
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form25.cboStUserCode.Text)
                cmd.Parameters.Append Param2
                 
                 
'                 If Form25.Check1(0).Value = 1 Then
                        Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 20, 1)
                        cmd.Parameters.Append Param3
''                 ElseIf Form25.Check1(1).Value = 1 Then
'                     Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 20, 2)
'                        cmd.Parameters.Append Param3
'                 End If
'
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Format(Trim(Form25.dtst_dt.Value), "DD-MMM-YYYY"))
                cmd.Parameters.Append Param4

                Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, Format(Trim(Form25.dted_dt.Value), "DD-MMM-YYYY"))
                cmd.Parameters.Append Param5
                
                Report17.Text6.SetText (Form25.cboStAccName.Text)
                Report17.Text7.SetText (Form25.dtst_dt.Value)
                Report17.Text19.SetText (Form25.dted_dt.Value)
'                Report12.Text22.SetText (Form15.cboStAccName.Text)
'
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL RptTax_vat_sch(?,?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report17.Database.SetDataSource RS
                    
                Report17.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report17
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
    Case 23
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
             
                  Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 1)
                      cmd.Parameters.Append Param1
         
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form26.cboStUserCode.Text)
                cmd.Parameters.Append Param2
                Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 20, 1)
                    cmd.Parameters.Append Param3
              
                  
                Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Format(Trim(Form26.dtst_dt.Value), "DD-MMM-YYYY"))
                cmd.Parameters.Append Param4

                Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 10, Format(Trim(Form26.dted_dt.Value), "DD-MMM-YYYY"))
                cmd.Parameters.Append Param5
                
                Report19.Text23.SetText (Form26.cboStAccName.Text)
                Report19.Text24.SetText (Form26.dtst_dt.Value)
                Report19.Text26.SetText (Form26.dted_dt.Value)
'                Report12.Text22.SetText (Form15.cboStAccName.Text)
'
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL Rpt_party_sch(?,?,?,?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report19.Database.SetDataSource RS
                    
                Report19.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report19
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
                
               
 
 Case 24
                Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                
               If Form28.Option1(0).Value = True Then
                  Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 1)
                      cmd.Parameters.Append Param1
                  Report20.Text4.SetText ("Income Budget Estimation")
               ElseIf Form28.Option1(1).Value = True Then
                   Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 2)
                      cmd.Parameters.Append Param1
                    Report20.Text4.SetText ("Expense Budget Estimation")
              End If
         
                Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Form28.Combo1)
                cmd.Parameters.Append Param2
                Report20.Text14.SetText (Form28.txtField(1))
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL rpt_budget(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report20.Database.SetDataSource RS
                    
                Report20.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report20
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
    Case 26
               Conn.Open strcn.Connection_String
                Set cmd.ActiveConnection = Conn
                cmd.CommandType = adCmdText
                 
                Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, Trim(Form29.dtst_dt.Value))
                cmd.Parameters.Append Param2
                
                Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, Trim(Form29.dted_dt.Value))
                cmd.Parameters.Append Param3
                
                Report21.Text7.SetText (Form29.dtst_dt.Value)
                Report21.Text9.SetText (Form29.dted_dt.Value)

                
                
                cmd.Properties("PLSQLRSet") = True
                cmd.CommandText = "{CALL rptbl(?,?)}"
                Set RS = cmd.Execute
                cmd.Properties("PLSQLRSet") = False
                
                Report21.Database.SetDataSource RS
                    
                Report21.DiscardSavedData
                Screen.MousePointer = vbHourglass
                CRViewer91.ReportSource = Report21
                CRViewer91.ViewReport
                Screen.MousePointer = vbDefault
             
                
    
 
 End Select
 Exit Sub
err_desc:
     Screen.MousePointer = vbNormal
     MsgBox Err.Description, vbCritical, "IT Division,DNMIH"
     
 End Sub
Private Sub Form_Resize()
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.Width = ScaleWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    RS.Close
End Sub
