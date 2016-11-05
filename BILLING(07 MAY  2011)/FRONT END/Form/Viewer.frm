VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Viewer 
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   Icon            =   "Viewer.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8010
      lastProp        =   500
      _cx             =   14129
      _cy             =   3889
      DisplayGroupTree=   0   'False
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
      EnableAnimationControl=   -1  'True
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
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Report1   As New CrystalReport1
Dim Report2   As New CrystalReport2
Dim Report3   As New CrysOutdoorPat
Dim Report5   As New CrystalReport5
Dim Report4   As New CrystalReport4
Dim Report4indoor As New CrystalReport4indoor
Dim Report6   As New CrystalReporttran
Dim Report7   As New CrystalReport7
Dim Report8   As New CrystalReporttran
Dim Report10   As New CrystalReport3
Dim report11 As New CrystalReport10
Dim Report9   As New CrystalReport9
Dim report12 As New CrystalReport11
Dim reportstaff As New CrystalReport12
Dim report13 As New CrystalReport13
Dim report14 As New CrystalReport14
Dim report15 As New CrystalReport15
Dim report17 As New CrystalReport17
Dim report18 As New CrystalReport18
Dim report19 As New CrystalReport19
Dim report20 As New CrystalReport20
Dim report16 As New CrystalReport16
Dim report21 As New CrystalReport21
Dim report22 As New CrystalReport22
Dim report23 As New CrystalReport23
Dim report24 As New CrystalReport24
Dim report25 As New CrystalReport25
Dim report26 As New CrystalReport26
Dim report27 As New CrystalReport27
Dim report28 As New CrystalReport28
Dim report29 As New CrystalReport29
Dim report30 As New CrystalReport30
Dim report31 As New CrystalReport31
Dim report33 As New CrystalReport33
Dim Reporttran   As New CrystalReporttran
Dim Reporto_diag   As New CrystalReport_diag
Dim Report_admission_summ As New CrystalReport8
Private Sub Form_Load()
On Error GoTo ERR_DESC
CRViewer91.Zoom 100
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

Select Case rptMode
    Case 7
        Screen.MousePointer = vbHourglass
    If Conn.State = 0 Then
        Conn.Open strcn.Connection_String
    End If
        
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 20, frmRelease.txtRegNoRelease.Text)
        cmd.Parameters.Append Param1 'combo
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL RptPatientRelease(?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report5.Database.SetDataSource RS
        
  ''''''''''for viewing'''''''''''''''
        Report5.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report5
        
        CRViewer91.ViewReport
        Report5.PrintOutEx
        Screen.MousePointer = vbDefault
        
 '''''''''''''''for printing''''''''''''
            Report5.PrintOut
            RS.Close
        
    Case 1
         If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, IntOption)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Trim(Rpt_doc_info.Combo1.Text))
        cmd.Parameters.Append Param2 'refer_code

        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL rptdoc_info(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report1.Database.SetDataSource RS
 ''''''''''''''for viewing''''''''''''''''
        Report1.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report1
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
'''''''''''''''for printing''''''''''''
            Report1.PrintOut
            RS.Close
 

    Case 2
        If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, 4)
        cmd.Parameters.Append Param1 'OPTION BUTTON
         Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Rpt_test_info.cboDeptCode.Text)
        cmd.Parameters.Append Param2 'combo
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, Trim(Rpt_test_info.Text1.Text))
        cmd.Parameters.Append Param3 'combo
       
        
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 20, "")
        cmd.Parameters.Append Param4 'combo
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Ret_Test_info_sub(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report2.Database.SetDataSource RS
   '''''''''''''''''for viewing
        Report2.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report2
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report5.PrintOut
'            RS.Close
'
            
    Case 3
        Dim var_val_chk As Integer
        var_val_chk = 0
        
        If Rpt_out_door_info.Check1.Value = 1 Then
        var_val_chk = 1
        
        Else
        var_val_chk = 0
        
        End If
        
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, IntOption)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Trim(Rpt_out_door_info.rptOutCombo.Text))
        cmd.Parameters.Append Param2 'on text department,main_name
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Rpt_out_door_info.DTPicker1.Value)
        cmd.Parameters.Append Param3 'date1
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Rpt_out_door_info.DTPicker2.Value)
        cmd.Parameters.Append Param4 'date2
        
        Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 4, var_val_chk)
        cmd.Parameters.Append Param5 'date1
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Rpt_out_door_info.DTPicker3.Value)
        cmd.Parameters.Append Param6 'date specific
        
        Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 4, Rpt_out_door_info.Check2.Value)
        cmd.Parameters.Append Param7 'shift_specific
        
        Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, Rpt_out_door_info.cboShift.Text)
        cmd.Parameters.Append Param8 'date specific
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptout_door_info(?,?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report3.Database.SetDataSource RS
        
        Report3.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report3
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
            
    Case 4
        Screen.MousePointer = vbHourglass
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
       
      Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 2)
      cmd.Parameters.Append Param1 'MODE
                  
      Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 15, Trim(Pat_Info_out.TXTRECEIPT_NO))
      cmd.Parameters.Append Param2 'RECEIPT NO
      
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptout_door_info_print(?,?)}"
        Set RS = cmd.Execute
        
        cmd.Properties("PLSQLRSet") = False
        
        Report4.Database.SetDataSource RS
 '''''''''''for viewing'''''''''''''''''''''
        Report4.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report4
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report4.PrintOut
'            RS.Close
'
    
 Case 24
        Screen.MousePointer = vbHourglass
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
       
      Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 2)
      cmd.Parameters.Append Param1 'comment
            
      Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 100, Trim(Pat_Info_out_for_indoor_test.TXTRECEIPT_NO))
      cmd.Parameters.Append Param2 'comment
                
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptout_door_info_print_indoor(?,?)}"
        Set RS = cmd.Execute
        
        cmd.Properties("PLSQLRSet") = False
        
        Report4indoor.Database.SetDataSource RS
 '''''''''''for viewing'''''''''''''''''''''
        Report4indoor.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report4indoor
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report4.PrintOut
'            RS.Close
'
   
    Case 5          ''Indoor pat
        Dim var_val_chk_1 As Integer
        var_val_chk_1 = 0
        
'        If Rpt_Indoor_door_info.Check1.Value = 1 Then
'        var_val_chk = 1
'
'        Else
'        var_val_chk_1 = 0
'
'        End If
        
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        If IntOption = 5 Then
            Report7.Text24.SetText ("Bed Wise")
            Report7.Text46.SetText ("Bed Type :")
            Report7.Text47.SetText (Rpt_Indoor_door_info.rptOutCombo.Text)
        Else
         Report7.Text24.SetText ("Department Wise income ")
          Report7.Text46.SetText ("Department Name :")
           Report7.Text47.SetText (Rpt_Indoor_door_info.rptOutCombo.Text)
        End If
        If var_val_chk = 1 Then
'          Report7.Text43.SetText (Rpt_Indoor_door_info.DTPicker3.Value)
          Report7.Text44.SetText ("")
          Report7.Text45.SetText ("")

        Else
          Report7.Text43.SetText (Rpt_Indoor_door_info.DTPicker1.Value)
          Report7.Text44.SetText ("To")
          Report7.Text45.SetText (Rpt_Indoor_door_info.DTPicker2.Value)
      End If
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 10, Rpt_Indoor_door_info.DTPicker1.Value)
        cmd.Parameters.Append Param1 'date1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Rpt_Indoor_door_info.DTPicker2.Value)
        cmd.Parameters.Append Param2 'date2
        
        Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 4, 0)
        cmd.Parameters.Append Param3 'date specific
        
         Set Param4 = cmd.CreateParameter("param4", adSmallInt, adParamInput, 5, 1)
        cmd.Parameters.Append Param4 'OPTION BUTTON
                  
        Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 30, Trim(Rpt_Indoor_door_info.rptOutCombo.Text))
        cmd.Parameters.Append Param5 'on text department,main_name
      
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_indoor_coll_sum(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report7.Database.SetDataSource RS
        
        Report7.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report7
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
   Case 6
           Screen.MousePointer = vbHourglass
           If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmd.Parameters.Append Param1 '
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
           cmd.Properties("PLSQLRSet") = False
          
          Report6.Database.SetDataSource RS
          Report6.Text4.Width = 1650
          Report6.Text4.SetText ("Admission")
        ''''''''''''''for viewing'''''''''''''''''''''''
            Report6.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report6
             CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            Set Report6 = Nothing
 '''''''''''''''for printing''''''''''''
            '''Report6.PrintOut
            '''RS.Close
            
                        
       Case 9
            Screen.MousePointer = vbHourglass
           If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
       
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
          Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, frm_reAdvance.TXTREC_NO)
            cmd.Parameters.Append Param1 '
            
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
           cmd.Properties("PLSQLRSet") = False
          
            Report8.Database.SetDataSource RS
            Report8.Text4.Width = 1650
            Report8.Text4.SetText ("Re-Advance")
            
            Report8.Text2.SetText ("Re-Advance")
            Report8.Text2.Font.Bold = True
            
            Report8.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report8
             CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
            Set Report8 = Nothing
            
  '''''''''''''''for printing''''''''''''
           ''' Report8.PrintOut
            ''''RS.Close
            
      Case 8
        Screen.MousePointer = vbHourglass
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 20, frmUlitity_release.txtRegNoRelease.Text)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, frmUlitity_release.CBOYRCODE.Text)
        cmd.Parameters.Append Param2 'combo
       
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL RptPatientRelease(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report5.Database.SetDataSource RS
  '''''''''''for viewing''''''''''''''''
        Report5.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report5
        
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
          
  '''''''''''''''for printing''''''''''''
            Report5.PrintOut
            RS.Close
           
 Case 10
        Dim var_val_chk1 As Integer
        var_val_chk1 = 0
        
        If Rpt_out_door_info.Check1.Value = 1 Then
        var_val_chk1 = 1
        
        Else
        var_val_chk1 = 0
        
        End If
        
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, IntOption)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Trim(Rpt_out_door_info.rptOutCombo.Text))
        cmd.Parameters.Append Param2 'on text department,main_name
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Rpt_out_door_info.DTPicker1.Value)
        cmd.Parameters.Append Param3 'date1
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Rpt_out_door_info.DTPicker2.Value)
        cmd.Parameters.Append Param4 'date2
        
        Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 4, var_val_chk1)
        cmd.Parameters.Append Param5 'date1
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Rpt_out_door_info.DTPicker3.Value)
        cmd.Parameters.Append Param6 'date specific
        
        Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 4, Rpt_out_door_info.Check2.Value)
        cmd.Parameters.Append Param7 'shift_specific
        
        Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, Rpt_out_door_info.cboShift.Text)
        cmd.Parameters.Append Param8 'date specific
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptout_door_info_summary(?,?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report10.Database.SetDataSource RS
        
        Report10.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report10
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
  Case 11
      On Error Resume Next
        Dim var_val_chk2 As Integer
        var_val_chk2 = 0
        
        If Rpt_IN_out_door_info_RECEIPT.Check1.Value = 1 Then
        var_val_chk2 = 1
        
        Else
        var_val_chk2 = 0
        
        End If
        
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, IntOption)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Trim(Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Text))
        cmd.Parameters.Append Param2 'on text department,main_name
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Rpt_IN_out_door_info_RECEIPT.DTPicker1.Value)
        cmd.Parameters.Append Param3 'date1
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Rpt_IN_out_door_info_RECEIPT.DTPicker2.Value)
        cmd.Parameters.Append Param4 'date2
        
        Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 4, var_val_chk2)
        cmd.Parameters.Append Param5 'flag
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Rpt_IN_out_door_info_RECEIPT.DTPicker3.Value)
        cmd.Parameters.Append Param6 'date specific
        
        Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 4, Rpt_IN_out_door_info_RECEIPT.Check2.Value)
        cmd.Parameters.Append Param7 'shift_specific
        
        Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, Rpt_IN_out_door_info_RECEIPT.cboShift.Text)
        cmd.Parameters.Append Param8 'date specific
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptin_out_door_info_receipt(?,?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        Report9.Database.SetDataSource RS
        If Rpt_IN_out_door_info_RECEIPT.cboShift.Enabled = True Then
                Report9.Text16.SetText (Rpt_IN_out_door_info_RECEIPT.cboShift)
'                 Report9.Text23.SetText ("")
         Else
            Report9.Text16.SetText ("")
'             Report9.Text23.SetText ("")
        End If
               Report9.Text17.SetText (Rpt_IN_out_door_info_RECEIPT.DTPicker3)
'                Report9.Text23.SetText ("")
        If Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Enabled = True Then
              Report9.Text18.SetText (Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Text + " - " + user_name)
               
         Else
            Report9.Text18.SetText ("")
'            Report9.Text23.SetText ("")

        End If
        If IntOption = "1" And var_val_chk2 = "0" And Rpt_IN_out_door_info_RECEIPT.Check2.Value = "0" Then
             Report9.Text17.SetText (Rpt_IN_out_door_info_RECEIPT.DTPicker1.Value)
            Report9.Text26.SetText (Rpt_IN_out_door_info_RECEIPT.DTPicker2.Value)
            Report9.Text25.SetText ("To")
            
        Else
           Report9.Text25.SetText (" ")
           Report9.Text26.SetText (" ")
        End If
        
        Report9.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report9
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
               
         'Report9.Text16.SetText = Nothing
 Case 12  '''''''''Month wise
       On Error Resume Next
        Dim Date_String
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
         Date_String = "1" + "-" + Mid(Rpt_IN_out_door_info_RECEIPT.cboshift1.Text, 1, 3) + "-" + Rpt_IN_out_door_info_RECEIPT.cboShift
         report11.Text8.SetText ("Monthly Income Statement: " + Rpt_IN_out_door_info_RECEIPT.cboshift1.Text + "," + Rpt_IN_out_door_info_RECEIPT.cboShift)
         
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 30, Trim(Date_String))
        cmd.Parameters.Append Param1 'on text department,main_name
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL rpt_month_wise(?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report11.Database.SetDataSource RS
        report11.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report11
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
               
         'Report9.Text16.SetText = Nothing
  Case 13
       On Error Resume Next
          '''''''''report on discount
       ' Dim Date_String
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
         'Date_String = "1" + "-" + Mid(Rpt_IN_out_door_info_RECEIPT.cboshift1.Text, 1, 3) + "-" + Rpt_IN_out_door_info_RECEIPT.cboShift
         report12.Text9.SetText (Rpt_discount.DTPicker1.Value)
         report12.Text11.SetText (Rpt_discount.DTPicker2.Value)
        
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 30, Trim(Rpt_discount.DTPicker1))
        cmd.Parameters.Append Param1 'on text department,main_name
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 30, Trim(Rpt_discount.DTPicker2))
        cmd.Parameters.Append Param2 'on text department,main_name
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptdiscount_info_receipt(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report12.Database.SetDataSource RS
        report12.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report12
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
               
         'Report9.Text16.SetText = Nothing
           
            
  Case 14
              
          If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
         
          report19.Text18.SetText (Rpt_discount_staff.DTPicker1.Value)
          report19.Text20.SetText (Rpt_discount_staff.DTPicker2.Value)
       
            report19.Text23.SetText (Rpt_discount_staff.TXTSTF_NAME)
         
        
           
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, "5")
        cmd.Parameters.Append Param1 'Option
       Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Rpt_discount_staff.txtstaffId)
        cmd.Parameters.Append Param2 'P_id
       
       Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 30, Trim(Rpt_discount_staff.DTPicker1.Value))
        cmd.Parameters.Append Param3 'on text department,main_name
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 30, Trim(Rpt_discount_staff.DTPicker2.Value))
        cmd.Parameters.Append Param4 'on text department,main_name
       
        
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptdiscount_receipt_DETAIL(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
       report19.Database.SetDataSource RS
        report19.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report19
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 ''''''
                     
       ''' Set cmd = Nothing
        
    Case 15
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        If frmpatient_history.Chk_date.Value = 1 Then
            report15.Text13.SetText (frmpatient_history.DTPicker1.Value)
            report15.Text25.Suppress = True
            report15.Text26.Suppress = True
        End If
        If frmpatient_history.Chk_date_to_date.Value = 1 Then
            report15.Text26.Suppress = False
            report15.Text13.SetText (frmpatient_history.DTPicker2.Value)
            report15.Text26.SetText (frmpatient_history.DTPicker3.Value)
            report15.Text25.Suppress = False
            
        Else
              report15.Text13.SetText (frmpatient_history.DTPicker1.Value)
              report15.Text25.Suppress = True
              report15.Text26.Suppress = True
         End If
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, optionbuttonval)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, frmpatient_history.TxtName)
        cmd.Parameters.Append Param2 'combo
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 20, frmpatient_history.DTPicker1)
        cmd.Parameters.Append Param3 'combo
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 20, frmpatient_history.DTPicker2)
        cmd.Parameters.Append Param4 'combo
        Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 20, frmpatient_history.DTPicker3)
        cmd.Parameters.Append Param5 'combo
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_patient_history(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report15.Database.SetDataSource RS
   '''''''''''''''''for viewing
        report15.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report15
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report5.PrintOut
'            RS.Close
'
 Case 16
         
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
       
       Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, optionbuttonval)
        cmd.Parameters.Append Param1 'OPTION BUTTON
       If optionbuttonval = 1 Then
'        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, frmpatient_history.txtName)
'        cmd.Parameters.Append Param2 'combo
         report17.Text14.SetText (frmDaily_collection_statement.DTPicker1.Value)
         report17.Text15.SetText ("")
         report17.Text16.SetText (" ")

       Else
           report17.Text14.SetText (frmDaily_collection_statement.DTPicker2.Value)
           report17.Text15.SetText ("TO")
           report17.Text16.SetText (frmDaily_collection_statement.DTPicker3.Value)
  End If
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 20, Format(frmDaily_collection_statement.DTPicker1.Value, "DD-mmm-YY"))
        cmd.Parameters.Append Param3 'combo
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 20, Format(frmDaily_collection_statement.DTPicker2.Value, "DD-MMM-YY"))
        cmd.Parameters.Append Param4 'combo
       Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 20, Format(frmDaily_collection_statement.DTPicker3.Value, "DD-MMM-YY"))
        cmd.Parameters.Append Param5 'combo
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_head_wise_coll(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report17.Database.SetDataSource RS
   '''''''''''''''''for viewing
        report17.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report17
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report5.PrintOut
'            RS.Close
'
 Case 17
          Screen.MousePointer = vbHourglass
           If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmd.Parameters.Append Param1 '
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
          
            Reporttran.Database.SetDataSource RS
  ''''''''''''''for viewing'''''''''''''''''''''''
            Reporttran.Text4.Width = 3100
            Reporttran.Text4.SetText ("Department Transfer")
            Reporttran.DoctorsDepartment1.Font.Bold = True
      
            Reporttran.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Reporttran
             CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
            Set Reporttran = Nothing
            
    Case 19
           If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        If optionbuttonval = 3 Then
            report18.Text14.SetText (frmSTATISTICS.DTPicker1.Value)
            report18.Text15.SetText ("")
            report18.Text12.SetText ("")
        End If
        
        If optionbuttonval = 4 Then
            report18.Text14.SetText (frmSTATISTICS.DTPicker2.Value)
            report18.Text15.SetText (frmSTATISTICS.DTPicker3.Value)
            report18.Text12.SetText ("To")
        End If
            
            
        report18.Text5.SetText (frmSTATISTICS.dept.Text)
        report18.Text41.SetText (frmSTATISTICS.dept.Text)

'        If frmpatient_history.Chk_date.Value = 1 Then
'            report15.Text13.SetText (frmpatient_history.DTPicker1.Value)
'            report15.Text25.Suppress = True
'            report15.Text26.Suppress = True
'        End If
'        If frmpatient_history.Chk_date_to_date.Value = 1 Then
'            report15.Text26.Suppress = False
'            report15.Text13.SetText (frmpatient_history.DTPicker2.Value)
'            report15.Text26.SetText (frmpatient_history.DTPicker3.Value)
'            report15.Text25.Suppress = False
'
'        Else
'              report15.Text13.SetText (frmpatient_history.DTPicker1.Value)
'              report15.Text25.Suppress = True
'              report15.Text26.Suppress = True
'         End If
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, optionbuttonval)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, frmSTATISTICS.dept)
        cmd.Parameters.Append Param2 'combo
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 20, frmSTATISTICS.DTPicker1)
        cmd.Parameters.Append Param3 'combo
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 20, frmSTATISTICS.DTPicker2)
        cmd.Parameters.Append Param4 'combo
       Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 20, frmSTATISTICS.DTPicker3)
        cmd.Parameters.Append Param5 'combo
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_patient_STATISTICS(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report18.Database.SetDataSource RS
   '''''''''''''''''for viewing
        report18.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report18
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 Case 20
         If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
          
          report19.Text18.SetText (Rptdiscount_detail.DTPicker1.Value)
          report19.Text20.SetText (Rptdiscount_detail.DTPicker2.Value)
        If Option_discount = 1 Then
            report19.Text23.SetText ("Hospital Staff")
         End If
          If Option_discount = 2 Then
             report19.Text23.SetText ("College Staff")
          End If
          
        If Option_discount = 3 Then
             report19.Text23.SetText ("Poor Patient(FREE-BED ONLY)")
          End If
        If Option_discount = 4 Then
             report19.Text23.SetText ("Committee Member")
        End If
        If Option_discount = 5 Then
             report19.Text23.SetText ("Poor Patient(OTHER THAN FREE-BED)")
          End If
          
           
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, Trim(Option_discount))
        cmd.Parameters.Append Param1 'on text department,main_name
       Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, "0")
        cmd.Parameters.Append Param2 'P_id
       
       Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 30, Trim(Rptdiscount_detail.DTPicker1.Value))
        cmd.Parameters.Append Param3 'on text department,main_name
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 30, Trim(Rptdiscount_detail.DTPicker2.Value))
        cmd.Parameters.Append Param4 'on text department,main_name
       
        
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptdiscount_receipt_DETAIL(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
       report19.Database.SetDataSource RS
        report19.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report19
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 ''''''
                     
        Set cmd = Nothing
        
        
   Case 21
          If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
            report20.Text17.SetText (frmSTATISTICS.DTPicker2.Value)
            report20.Text19.SetText (frmSTATISTICS.DTPicker3.Value)
       Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 10, frmSTATISTICS.DTPicker2.Value)
        cmd.Parameters.Append Param1 'combo
       Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, frmSTATISTICS.DTPicker3.Value)
        cmd.Parameters.Append Param2 'combo
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL all_depart_stat(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report20.Database.SetDataSource RS
   '''''''''''''''''for viewing
         report20.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report20
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 Case 30
        
            If Conn.State = 1 Then
               Conn.Close
            Else
                Conn.Open strcn.Connection_String
            End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            
            
            ''''''Dim Report8   As New CrystalReport8
           ''''' Dim Param1 As New Parameter
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 30, frmadm_pat_cancel.TXTREC_NO)
            cmd.Parameters.Append Param1 'IN_REG_NO
            Report8.Text4.Width = 4000
            Report8.Text4.SetText ("Admission Cancellation")
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report8.Database.SetDataSource RS


        
  ''''''''''''''for viewing'''''''''''''''''''''''
            Report8.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Report8
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
          
   Case 41
      On Error Resume Next
        ''Dim var_val_chk2 As Integer
        var_val_chk2 = 0
        
        If Rpt__REC_DET.Check1.Value = 1 Then
        var_val_chk2 = 1
        
        Else
        var_val_chk2 = 0
        
        End If
        
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, IntOption)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Trim(Rpt__REC_DET.rptOutCombo.Text))
        cmd.Parameters.Append Param2 'on text department,main_name
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Rpt__REC_DET.DTPicker1.Value)
        cmd.Parameters.Append Param3 'date1
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Rpt__REC_DET.DTPicker2.Value)
        cmd.Parameters.Append Param4 'date2
        
        Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 4, var_val_chk2)
        cmd.Parameters.Append Param5 'date1____-flag
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Rpt__REC_DET.DTPicker3.Value)
        cmd.Parameters.Append Param6 'date specific
        
        Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 4, Rpt__REC_DET.Check2.Value)
        cmd.Parameters.Append Param7 'shift_specific
        
        Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, Rpt__REC_DET.cboShift.Text)
        cmd.Parameters.Append Param8 'date specific
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rptinout_door_inf_rec_Detail(?,?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report21.Database.SetDataSource RS
        report21.txtrptdate.SetText (Rpt__REC_DET.DTPicker3.Value)

'        If Rpt_IN_out_door_info_RECEIPT.cboShift.Enabled = True Then
'                Report9.Text16.SetText (Rpt_IN_out_door_info_RECEIPT.cboShift)
'                 Report9.Text23.SetText ("")
'         Else
'            Report9.Text16.SetText ("")
'             Report9.Text23.SetText ("")
'        End If
'               Report9.Text17.SetText (Rpt_IN_out_door_info_RECEIPT.DTPicker3)
'                Report9.Text23.SetText ("")
'        If Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Enabled = True Then
'              Report9.Text18.SetText (Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Text)
'               Report9.Text23.SetText (user_name)
'         Else
'            Report9.Text18.SetText ("")
'            Report9.Text23.SetText ("")
'
'        End If
        
        report21.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report21
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
    Case 42
         Screen.MousePointer = vbHourglass
           If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmd.Parameters.Append Param1 '
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_Others_money(?)}"
            Set RS = cmd.Execute
           cmd.Properties("PLSQLRSet") = False
          
          report13.Database.SetDataSource RS
  ''''''''''''''for viewing'''''''''''''''''''''''
            report13.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = report13
             CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
            '''Report6.PrintOut
            '''RS.Close
            
   Case 43
        If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, Trim(frm_Diagnostic_refund.TXTPRINT_diag.Text))
        cmd.Parameters.Append Param1 'comment
  
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_diag_refund(?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Reporto_diag.Database.SetDataSource RS
 ''''''''''''''for viewing''''''''''''''''
        Reporto_diag.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Reporto_diag
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
'''''''''''''''for printing''''''''''''
'            Report1.PrintOut
'            RS.Close
 
    Case 44
        If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
       
       Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, optionbuttonval)
        cmd.Parameters.Append Param1 'OPTION BUTTON
       If optionbuttonval = 1 Then
'        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, frmpatient_history.txtName)
'        cmd.Parameters.Append Param2 'combo
         report22.Text8.SetText (frmDaily_collection_stat.DTPicker1.Value)
         report22.Text9.SetText ("")
         report22.Text10.SetText (" ")

       Else
           report22.Text8.SetText (frmDaily_collection_stat.DTPicker2.Value)
           report22.Text9.SetText ("TO")
           report22.Text10.SetText (frmDaily_collection_stat.DTPicker3.Value)
  End If
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 20, Format(frmDaily_collection_stat.DTPicker1.Value, "DD-mmm-YY"))
        cmd.Parameters.Append Param3 'combo
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 20, Format(frmDaily_collection_stat.DTPicker2.Value, "DD-MMM-YY"))
        cmd.Parameters.Append Param4 'combo
       Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 20, Format(frmDaily_collection_stat.DTPicker3.Value, "DD-MMM-YY"))
        cmd.Parameters.Append Param5 'combo
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_head_wise_statistics(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report22.Database.SetDataSource RS
   '''''''''''''''''for viewing
        report22.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report22
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report5.PrintOut
'            RS.Close



   Case 46
        On Error Resume Next
        Dim var_val_chk3 As Integer
        var_val_chk3 = 0
        
        If Rpt_receipt_group.Check1.Value = 1 Then
            var_val_chk3 = 1
        
        Else
           var_val_chk3 = 0
        
        End If
        
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'OPTION BUTTON
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, Trim(Rpt_receipt_group.rptOutCombo.Text))
        cmd.Parameters.Append Param2 'on text department,main_name
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Rpt_receipt_group.DTPicker1.Value)
        cmd.Parameters.Append Param3 'date1
        
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Rpt_receipt_group.DTPicker2.Value)
        cmd.Parameters.Append Param4 'date2
        
        Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 4, var_val_chk3)
        cmd.Parameters.Append Param5 'date1____-flag
        
        Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, Rpt_receipt_group.DTPicker3.Value)
        cmd.Parameters.Append Param6 'date specific
        
        Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 4, Rpt_receipt_group.Check2.Value)
        cmd.Parameters.Append Param7 'shift_specific
        
        Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 10, Rpt_receipt_group.cboShift.Text)
        cmd.Parameters.Append Param8 'date specific
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL  Rpt_receipt_group(?,?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report23.Database.SetDataSource RS
'        If Rpt_IN_out_door_info_RECEIPT.cboShift.Enabled = True Then
'                Report9.Text16.SetText (Rpt_IN_out_door_info_RECEIPT.cboShift)
'                 Report9.Text23.SetText ("")
'         Else
'            Report9.Text16.SetText ("")
'             Report9.Text23.SetText ("")
'        End If
'               Report9.Text17.SetText (Rpt_IN_out_door_info_RECEIPT.DTPicker3)
'                Report9.Text23.SetText ("")
'        If Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Enabled = True Then
'              Report9.Text18.SetText (Rpt_IN_out_door_info_RECEIPT.rptOutCombo.Text)
'               Report9.Text23.SetText (user_name)
'         Else
'            Report9.Text18.SetText ("")
'            Report9.Text23.SetText ("")
'
'        End If
        report23.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report23
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
        
        
        
    Case 47 '''''group wise collection statement
         
       If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
       
       Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, optionbuttonval)
        cmd.Parameters.Append Param1 'OPTION BUTTON
       If optionbuttonval = 1 Then
'        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, frmpatient_history.txtName)
'        cmd.Parameters.Append Param2 'combo
         report24.Text9.SetText (frmgroup_statement.DTPicker1.Value)
         report24.Text12.SetText ("")
         report24.Text14.SetText (" ")

       Else
           report24.Text9.SetText (frmgroup_statement.DTPicker2.Value)
           report24.Text12.SetText ("TO")
           report24.Text14.SetText (frmgroup_statement.DTPicker3.Value)
  End If
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 20, Format(frmgroup_statement.DTPicker1.Value, "DD-mmm-YY"))
        cmd.Parameters.Append Param3 'combo
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 20, Format(frmgroup_statement.DTPicker2.Value, "DD-MMM-YY"))
        cmd.Parameters.Append Param4 'combo
       Set Param5 = cmd.CreateParameter("param5", adDate, adParamInput, 20, Format(frmgroup_statement.DTPicker3.Value, "DD-MMM-YY"))
        cmd.Parameters.Append Param5 'combo
       
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_group_wise_state(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report24.Database.SetDataSource RS
   '''''''''''''''''for viewing
        report17.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report24
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
 '''''''''''''''for printing''''''''''''
'            Report5.PrintOut
'            RS.Close

Case 48
        On Error GoTo ERR_DESC
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
         
        report25.Text10.SetText (frm_abscond.DTPicker1.Value)
        report25.Text14.SetText (frm_abscond.DTPicker2.Value)
'
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 30, 1) ''OPTION
        cmd.Parameters.Append Param1
       
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 30, Trim(frm_abscond.DTPicker1))
        cmd.Parameters.Append Param2
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 30, Trim(frm_abscond.DTPicker2))
        cmd.Parameters.Append Param3
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_FLED_PAT(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report25.Database.SetDataSource RS
        report25.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report25
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
  Case 313
        On Error Resume Next
          '''''''''report on discount
       ' Dim Date_String
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
'         'Date_String = "1" + "-" + Mid(Rpt_IN_out_door_info_RECEIPT.cboshift1.Text, 1, 3) + "-" + Rpt_IN_out_door_info_RECEIPT.cboShift
         report26.Text12.SetText (Rpt_advance_reg.DTPicker1.Value)
         report26.Text15.SetText (Rpt_advance_reg.DTPicker2.Value)
       
        If Rpt_advance_reg.Option1(0).Value = True Then
           Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
           cmd.Parameters.Append Param1 'on text department,main_name
         report26.Text20.SetText ("All")
         report26.Text21.SetText ("All User")
       
        ElseIf Rpt_advance_reg.Option1(1).Value = True Then
              Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 2)
              cmd.Parameters.Append Param1 'on text department,main_name
              
         report26.Text20.SetText (Rpt_advance_reg.Combo1.Text)
         report26.Text21.SetText (Rpt_advance_reg.Text1.Text)
       
        End If
        
      
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 30, Trim(Rpt_advance_reg.DTPicker1))
        cmd.Parameters.Append Param2 'on text department,main_name
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 30, Trim(Rpt_advance_reg.DTPicker2))
        cmd.Parameters.Append Param3 'on text department,main_name
        
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Trim(Rpt_advance_reg.Combo1.Text))
        cmd.Parameters.Append Param4 'on text department,main_name
       
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL advance_reg(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report26.Database.SetDataSource RS
        report26.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report26
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
   
   Case 413
        On Error Resume Next
          '''''''''report on discount
       ' Dim Date_String
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
'         'Date_String = "1" + "-" + Mid(Rpt_IN_out_door_info_RECEIPT.cboshift1.Text, 1, 3) + "-" + Rpt_IN_out_door_info_RECEIPT.cboShift
'         report27.Text12.SetText (Rpt_advance_reg_REG.DTPicker1.Value)
'         report27.Text15.SetText (Rpt_advance_reg_REG.DTPicker2.Value)
''
           Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
           cmd.Parameters.Append Param1 'on text department,main_name
     
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 30, Trim(Rpt_advance_reg_REG.DTPicker1))
        cmd.Parameters.Append Param2 'on text department,main_name
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 30, Trim(Rpt_advance_reg_REG.DTPicker2))
        cmd.Parameters.Append Param3 'on text department,main_name
        
          
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL advance_reg_reGnowISE(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report27.Database.SetDataSource RS
        report27.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report27
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
   Case 500
        On Error Resume Next
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        report28.Text8.SetText (frmDiagnostic_Income.MaskEdBox1.Text + "     To    " + frmDiagnostic_Income.MaskEdBox2)
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'mode
     
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox1))
        cmd.Parameters.Append Param2 'date 1
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox2))
        cmd.Parameters.Append Param3 'date 2
          
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_diag_income_dept_Wise(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report28.Database.SetDataSource RS
        frmDiagnostic_Income.Label2.Caption = " "
        report28.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report28
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
    Case 501
        On Error Resume Next
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        report16.Text9.SetText (frmDiagnostic_Income.MaskEdBox1.Text + "     To    " + frmDiagnostic_Income.MaskEdBox2)
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, 2)
        cmd.Parameters.Append Param1 'mode
     
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox1))
        cmd.Parameters.Append Param2 'date 1
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox2))
        cmd.Parameters.Append Param3 'date 2
          
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_diag_income_dept_Wise(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report16.Database.SetDataSource RS
        frmDiagnostic_Income.Label2.Caption = " "
        report16.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report16
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
  Case 502
        If Conn.State = 0 Then
           Conn.Open strcn.Connection_String
        End If
        
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        
         report29.Text24.SetText (CStr(Rpt_Indoor_door_info.DTPicker1) + "     To    " + CStr(Rpt_Indoor_door_info.DTPicker2))
         If IntOption = 1 Then
             report29.Text46.SetText ("Indoor Income Only")
         ElseIf IntOption = 2 Then
             report29.Text46.SetText ("Outdoor Income Only")
         Else
             report29.Text46.SetText ("Both Indoor and Outdoor Income")
         End If
         Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 2, IntOption)
         cmd.Parameters.Append Param1 'param mode
      
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Rpt_Indoor_door_info.DTPicker1.Value)
        cmd.Parameters.Append Param2 'date1
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Rpt_Indoor_door_info.DTPicker2.Value)
        cmd.Parameters.Append Param3 'date2
        
         
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL ALL_DEPT_INCOME(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        
        report29.Database.SetDataSource RS
        
        report29.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report29
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
   Case 504
        On Error Resume Next
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        Report_admission_summ.Text11.SetText (RPT_PAT_ADMISSION.MaskEdBox1.Text + "     To    " + RPT_PAT_ADMISSION.MaskEdBox2)
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'mode
     
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Trim(RPT_PAT_ADMISSION.MaskEdBox1))
        cmd.Parameters.Append Param2 'date 1
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Trim(RPT_PAT_ADMISSION.MaskEdBox2))
        cmd.Parameters.Append Param3 'date 2
          
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL RptAdmission_Summary(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        Report_admission_summ.Database.SetDataSource RS
        frmDiagnostic_Income.Label2.Caption = " "
        Report_admission_summ.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = Report_admission_summ
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
   Case 505
        On Error Resume Next
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        report30.Text8.SetText (frmDiagnostic_Income.MaskEdBox1.Text + "     To    " + frmDiagnostic_Income.MaskEdBox2)
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'mode
     
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox1))
        cmd.Parameters.Append Param2 'date 1
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox2))
        cmd.Parameters.Append Param3 'date 2
          
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Diag_Income_Test_Wise(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report30.Database.SetDataSource RS
       frmDiagnostic_Income.Label2.Caption = " "
        report30.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report30
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        Set Conn = Nothing
        Set cmd = Nothing
        Set RS = Nothing
        Set report30 = Nothing
 Case 506
          On Error Resume Next
      If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        report30.Text8.SetText (frmDiagnostic_Income.MaskEdBox1.Text + "     To    " + frmDiagnostic_Income.MaskEdBox2)
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'mode
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 5, Trim(frmDiagnostic_Income.cboMainCode))
        cmd.Parameters.Append Param2 'Head code
     
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox1))
        cmd.Parameters.Append Param3 'date 1
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 10, Trim(frmDiagnostic_Income.MaskEdBox2))
        cmd.Parameters.Append Param4 'date 2
          
       
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Diag_Income_Test_Head_Wise(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report14.Database.SetDataSource RS
        frmDiagnostic_Income.Label2.Caption = " "
        
       
        report14.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report14
         Dim dateString As String
         dateString = Param3 & "    TO    " & Param4
         report14.Text14.SetText ("Test Head Title:    " & frmDiagnostic_Income.txtMainName)
         report14.Text5.SetText (dateString)
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        Set Conn = Nothing
        Set cmd = Nothing
        Set RS = Nothing
        Set report14 = Nothing
      
  Case 507
          On Error Resume Next
         If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 10, Rpt_SHIFTWISE_PAT_ADMISSION.DTPicker1.Value)
        cmd.Parameters.Append Param1 'date
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(Rpt_SHIFTWISE_PAT_ADMISSION.cboShift))
        cmd.Parameters.Append Param2 'Shift
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(Rpt_SHIFTWISE_PAT_ADMISSION.CBOYRCODE))
        cmd.Parameters.Append Param3 'yEAR CODE
     
         
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL RPT_SHIFTWISE_ADMISSION(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report31.Database.SetDataSource RS
       
        
       
        report31.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report31
        report31.Text11.SetText (Param1 & " ( " & Param2 & ")")
        
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        Set Conn = Nothing
        Set cmd = Nothing
        Set RS = Nothing
        Set report31 = Nothing
          
  Case 508
        On Error Resume Next
        If Conn.State = 0 Then
               Conn.Open strcn.Connection_String
         End If
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        Set Param1 = cmd.CreateParameter("param1", adSmallInt, adParamInput, 10, paramMode)
        cmd.Parameters.Append Param1 'option
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(CurrentPatientUI.BedTypeCombo.Text))
        cmd.Parameters.Append Param2 'bed type
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(CurrentPatientUI.CabinOrWardCombo.Text))
        cmd.Parameters.Append Param3 'cabin or ward type
        
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, Trim(CurrentPatientUI.DepartmentCombo.Text))
        cmd.Parameters.Append Param4 '''department
        
        Set Param5 = cmd.CreateParameter("param5", adInteger, adParamInput, 3, PatientStatus)
        cmd.Parameters.Append Param5 '''PATEINT TYPE
         
        Set Param6 = cmd.CreateParameter("param6", adInteger, adParamInput, 4, Trim(CurrentPatientUI.DaysAboveCombo.Text))
        cmd.Parameters.Append Param6 ''' DAYS ABOVE
        
        Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, Trim(CurrentPatientUI.FiscalYearsCombo.Text))
        cmd.Parameters.Append Param7
         
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_CUR_ADM_PAT_ADVANCE_COLL(?,?,?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
       
        report33.Database.SetDataSource RS
       
        report33.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = report33
        report33.Text10.SetText (CurrentPatientUI.Option(paramMode).Caption)
        Dim helper As String
        If paramMode = 0 Then
            helper = "Bed Type :  " + CurrentPatientUI.BedTypeCombo.Text
          
        ElseIf paramMode = 1 Then
           helper = "Bed Type : " + CurrentPatientUI.BedTypeCombo.Text + "   " + "Ward/Cabin  : " + CurrentPatientUI.CabinOrWardCombo.Text
        ElseIf paramMode = 2 Then
           helper = "Bed Type : " + CurrentPatientUI.BedTypeCombo.Text + "   " + "Department : " + CurrentPatientUI.DepartmentCombo.Text
        ElseIf paramMode = 4 Then
          helper = "Bed Type : " + CurrentPatientUI.BedTypeCombo.Text + "  Patient List of Staying More than" + CurrentPatientUI.DaysAboveCombo.Text + "  Day(s)"
        End If
        report33.Text6.SetText (helper)

        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        Set Conn = Nothing
        Set cmd = Nothing
        Set RS = Nothing
        Set report33 = Nothing
  End Select
                

                
    Exit Sub
ERR_DESC:
       MsgBox Err.Description, vbCritical, "DNMIH."
           
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
