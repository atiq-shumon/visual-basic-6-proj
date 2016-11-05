VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Form20 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Viewer"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "frmReport_Viewer.frx":0000
   LinkTopic       =   "Form20"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CRViewer1 
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   11775
      TabIndex        =   1
      Top             =   0
      Width           =   11835
      Begin CRVIEWERLibCtl.CRViewer CRViewer2 
         Height          =   8250
         Left            =   135
         TabIndex        =   2
         Top             =   90
         Width           =   11835
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
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer preference"
      ForeColor       =   &H000000C0&
      Height          =   645
      Left            =   8865
      TabIndex        =   0
      Top             =   45
      Width           =   1500
      Begin VB.Image Image1 
         Height          =   480
         Left            =   495
         Picture         =   "frmReport_Viewer.frx":08CA
         Stretch         =   -1  'True
         Top             =   180
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rept_rs As New ADODB.Recordset
Dim objRpt As New clsReport
Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cmd As New Command
Dim report1 As New CrystalReport1       '' Employee Information
Dim report2 As New CrysEmployeeInfo
'Dim report3 As New CrysSalaryPreparation1
Dim report3 As New CrysSalaryPreparation
Dim report4 As New CrysEmployee_PF
'Dim Report5 As New CrysOvertimePre
Dim Report5 As New CrysOvertimePreparation
Dim Report6 As New CrysJobStatusInfo
Dim reportIT6 As New CrystalReport6
Dim report7 As New CrystLoanLedgerInfo
Dim report8 As New CrysPayslip
Dim report9 As New CrysLeaveApplication
Dim report10 As New CrysIncrementReport
'Dim report10 As New CrysIncrementRpt
Dim report11 As New CrysSalaryStatmentToBank
Dim report12 As New CrysProvidentYrrlRpt
Dim report13 As New CrysEmpLoanStatatus
Dim Report14 As New CrysSalaryPreparation1
'Dim report15  As New RptToWhomitmayConcern
Dim report15  As New CrysToWhomitmeConcern1
Dim report16 As New CrysLaonStatment
Dim report17 As New CrysLoanScheduleforAll
Dim report18 As New CrysPromotionRecord
Dim REPORT19 As New CrysAllEmpID
Dim REPORT20 As New CrysRetiredEmpReport
Dim report21 As New CrysIndividualSalaryPre
Dim report22 As New RptBonusPreparation
Dim report23 As New CrysBonusSendtoBank
Dim report24 As New CrysBonusPrePayslip
Dim report25 As New CrysDeptwiseSal
'Dim report26 As New CrysLoanSummaryAll
Dim report27 As New CrysEmpPFEach
Dim report28 As New CrysGratuityIncomeExpence
Dim report29 As New CrysGratuityCashBook
Dim report30 As New CrysGratuityOpeningClosingEndofYear
Dim report31 As New CrysBankAccountStatement
Dim report32 As New CrysGratuityBalanceSheet
Dim report33 As New CrysPFIncomeExpence
Dim report34 As New CrysPFCashBook
Dim report35 As New CrysGraSourceofGratuityWise
Dim report36 As New CrysPFBalanceSheet
Dim report37 As New CrysPFBankAccountStatement
Dim report38 As New CrysGraPurposeofPaymentWise
Dim report39 As New CrysGraReceiveBankWise
Dim report40 As New CrysGraPaymentBankWise
Dim report41 As New CrysPFReceiveSourceofFund
Dim report42 As New CrysPFPurposeofPaymentWise
Dim report43 As New CrystalReport2
Dim report44 As New CrystalReport3
Dim report45 As New CrystalReport4
Dim report48 As New CrysPayrollStatistics
Dim report49 As New CrystalReport5


Dim Param1 As New ADODB.Parameter
Dim Param2 As New ADODB.Parameter
Dim Param3 As New ADODB.Parameter
Dim Param4 As New ADODB.Parameter


Dim Param5 As New ADODB.Parameter
Dim Param6 As New ADODB.Parameter
Dim Param7 As New ADODB.Parameter
Dim SourceId As New ADODB.Parameter
Dim BankCode As New ADODB.Parameter
Dim PurposeId As New ADODB.Parameter









Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth


CRViewer2.Top = 0
CRViewer2.Left = 0
CRViewer2.Height = ScaleHeight
CRViewer2.Width = ScaleWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
     con.Close
    Set cmd = Nothing
    Set RS = Nothing
    
    Form20.MousePointer = 1
    Unload Me
    Destroy Me
End Sub
Private Sub Form_Load()
'On Error GoTo Errdes
CRViewer2.Zoom (100)
'CRViewer1.Zoom (100)
Select Case rptmode
    Case 3
    
        With objRpt
            .Connstring = strCN.Connection_String
           ' Set rept_rs = .Employee_Information(Form21.Text1)
        End With

        report1.Database.SetDataSource rept_rs
        report1.DiscardSavedData


        Screen.MousePointer = vbHourglass

        CRViewer2.ReportSource = report1
        CRViewer2.ViewReport

        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
         
    Case 4
    

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Form21.Text1)
        cmd.Parameters.Append Param1 'combo
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL RPT_EMPINFO_ALL(?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report2.Database.SetDataSource RS
        
        report2.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report2
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
       
   Case 5
   
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = " SELECT EMP_ID, PAY_MONTH, PAY_YEAR, BASIC, H_RENT, MED, " + _
                       " CONV, TFN, DA, R_STAMP, LN_AMOUNT, ADV_AMOUNT, " + _
                        " PF_DEDUCTION, OTHERS_ALLOWANCE,OTHERS_DEDUCTION , PAY_STAT, Adv_Id From SALARY "

        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report3.Database.SetDataSource RS
        
        report3.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report3
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
Case 6

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_EmpInfo_ExpendedSalary}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report4.Database.SetDataSource RS
        
        report4.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report4
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing

Case 7

        Screen.MousePointer = vbHourglass
     
        con.Open strCN.Connection_String
        
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        
        
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, Emp_ID_Value)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 20, EnddateforReport)
        cmd.Parameters.Append Param2 'combo
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Overtime_Preparation(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report5.Database.SetDataSource RS
        
        Report5.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = Report5
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        Emp_ID_Value = ""
        EnddateforReport = ""
        
Case 8

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_EmpJobStatus_Information}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        Report6.Database.SetDataSource RS
        
        Report6.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = Report6
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        
Case 9

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, GetMonthOftheYear)
        cmd.Parameters.Append Param1 'combo


        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 20, Trim(GetSalaryPreparationYaer))
        cmd.Parameters.Append Param2 'combo

        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, CheckStatusofEmployee)
        cmd.Parameters.Append Param3 'combo

        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 100, ComboValue_Dept)
        cmd.Parameters.Append Param4 'combo
        
        Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 1, srlType)
        cmd.Parameters.Append Param5 'Salary Type
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Salry_Preparation_deptwise(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report25.Database.SetDataSource RS
        
        report25.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report25
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
      

Case 10

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, Emp_ID_Value)
        cmd.Parameters.Append Param1 'combo
        
       

        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL RPT_PF_EACH_EMP (?)}" '
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report27.Database.SetDataSource RS '
        
        report27.DiscardSavedData '
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report27 '
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
      
 Case 11
        
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Emp_ID_Value_ForLoan)
        cmd.Parameters.Append Param1 'combo
        
        cmd.Properties("PLSQLRSet") = True
        
   
        
        cmd.CommandText = "{CALL Rpt_Ln_Register(?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report7.Database.SetDataSource RS

        report7.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report7
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
Case 12
        
       
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, GetMonthOftheYear)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("Param2", adVarChar, adParamInput, 100, Rpt_Year)
        cmd.Parameters.Append Param2 'combo
        
        Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 100, Emp_ID_Value)
        cmd.Parameters.Append Param3 'combo
        
        
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_PaySlip_Preparation(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        report8.Database.SetDataSource RS
        
        report8.Text103.SetText "Pay Slip For the Month of " & GetMonthOftheYear & " " & Rpt_Year & " :- "
        report8.Text88.SetText "Pay Slip For the Month of " & GetMonthOftheYear & " " & Rpt_Year & " :- "
        
        report8.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report8
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing

Case 13

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("Param2", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 15, Emp_IDforLeave)
        cmd.Parameters.Append Param3 'combo
        
        
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_Leave_Application(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False


                
        report9.Database.SetDataSource RS

        report9.Text25.SetText "-" & Mid(Format(Date$, "dd-mm-yyyy"), 7, 10)
        
        report9.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report9
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
Case 14

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        
        Set Param1 = cmd.CreateParameter("param2", adChar, adParamInput, 10, Emp_ID_Value)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("param2", adChar, adParamInput, 100, DEPARMENTNAMEFORTPT)
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 15, CheckStatusofEmployee)
        cmd.Parameters.Append Param3 'combo

        
                
        Set Param4 = cmd.CreateParameter("param4", adDate, adParamInput, 15, BeginDateOfIncremnt)
        cmd.Parameters.Append Param4 'combo
        
        Set Param5 = cmd.CreateParameter("Param5", adDate, adParamInput, 15, EndDateOfIncremnt)
        cmd.Parameters.Append Param5
                
                
                
                
                
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_Next_Increment_Report(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False


                
        report10.Database.SetDataSource RS

        report10.Text16.SetText "(From " & BeginDateOfIncremnt
        report10.Text17.SetText " To " & EndDateOfIncremnt & ")"
        
     
        If CheckStatusofEmployee <> 0 And Len(DEPARMENTNAMEFORTPT) = 0 Then
            If CheckStatusofEmployee = "1" Then
                report10.Text30.SetText "Yearly Increment of First Class Employee(s)"
            ElseIf CheckStatusofEmployee = "2" Then
                report10.Text30.SetText "Yearly Increment of Second Class Employee(s)"
            ElseIf CheckStatusofEmployee = "3" Then
                report10.Text30.SetText "Yearly Increment of Third Class Employee(s)"
            Else
                report10.Text30.SetText "Yearly Increment of Forth Class Employee(s)"
            End If
         End If
       
        If CheckStatusofEmployee = 0 And Len(DEPARMENTNAMEFORTPT) <> 0 Then
                report10.Text30.SetText "Yearly Increment of the Employee(s) in the Department of :" & DEPARMENTNAMEFORTPT
        End If
        
         
        If CheckStatusofEmployee = 0 And Len(Trim(DEPARMENTNAMEFORTPT)) = 0 Then
                report10.Text30.SetText "Increment Information of Specific Employee"
        End If
        
        report10.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report10
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        Emp_ID_Value = ""
        CheckStatusofEmployee = 0
        DEPARMENTNAMEFORTPT = ""
        BeginDateOfIncremnt = ""
        EndDateOfIncremnt = ""
        
Case 15

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        Dim bankSalaryText As String
         Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, GetMonthOftheYear)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("Param2", adVarChar, adParamInput, 15, GetSalaryPreparationYaer)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 15, Trim(StatusofEmployee))
        cmd.Parameters.Append Param3
        
        Set Param4 = cmd.CreateParameter("Param4", adVarChar, adParamInput, 1, localSalaryType)
        cmd.Parameters.Append Param4
       
                
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_Bank_Payment_Salary(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False


                
        report11.Database.SetDataSource RS

        bankSalaryText = ""
       
        report11.Text9.SetText "(" & GetMonthOftheYear & "-" & GetSalaryPreparationYaer & " )"
        
        If localSalaryType = "R" Then
             bankSalaryText = "Details Salary Disbursement Statement of the Employees of "
        ElseIf localSalaryType = "S" Then
             bankSalaryText = "Details Supplementary Salary Disbursement Statement of the Employees of "
        ElseIf localSalaryType = "B" Then
            bankSalaryText = "Festival Bonus Disbursement Statement of the Employees of "
         ElseIf localSalaryType = "D" Then
            bankSalaryText = "Dress Allowance Disbursement Statement of the Employees of "
        End If
        
        If StatusofEmployee = "1" Then
           bankSalaryText = bankSalaryText + ":First Class"
        ElseIf StatusofEmployee = "2" Then
             bankSalaryText = bankSalaryText + ":Second Class"
        ElseIf StatusofEmployee = "3" Then
             bankSalaryText = bankSalaryText + " :Third Class"
        ElseIf StatusofEmployee = "4" Then
             bankSalaryText = bankSalaryText + ":Fourth Class"
        Else
             bankSalaryText = bankSalaryText + ":ALL Class"
        End If
        
        report11.Text1.SetText (bankSalaryText)
        
        report11.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report11
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
Case 16

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
         Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, Emp_ID_Value)
        cmd.Parameters.Append Param1 'combo
        
                      
        cmd.Properties("PLSQLRSet") = True
        
        
       ' cmd.CommandText = "{CALL Rpt_PF_BalanceEndofYear(?)}"
        
        cmd.CommandText = "{CALL Rpt_PF_Bal_EndofYear(?)}"
       
        
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        report12.Database.SetDataSource RS

        report12.Text2.SetText Format(Date, "MMM") & ",  " & Format(Date, "yyyy")
        
        
        
        
        report12.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report12
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        
Case 17

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, Emp_ID_Value_ForLoan)
        cmd.Parameters.Append Param1 'combo
        
        
                
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_Ln_Register_PerEmp(?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        report13.Database.SetDataSource RS

                   
        report13.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report13
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing

Case 18
        Dim salaryText As String
        
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, GetMonthOftheYear)
        cmd.Parameters.Append Param1 'combo
        
'        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, Trim(frmSalaryReportForm.cboYear))
'        cmd.Parameters.Append Param2 'combo
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, Trim(GetSalaryPreparationYaer))
        cmd.Parameters.Append Param2 'combo
        
        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 100, ComboValue_Dept)
        cmd.Parameters.Append Param3 'combo
       

        Set Param4 = cmd.CreateParameter("param4", adChar, adParamInput, 1, localSalaryType)
        cmd.Parameters.Append Param4 'Salary Type

        
    
        
                
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_Salry_Preparation_Final(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        If localSalaryType = "D" Then '' dress allowance only
           report45.Database.SetDataSource RS
        Else
          Report14.Database.SetDataSource RS
        End If
        
        
        salaryText = ""
        
        If Param3 = 1 Then
           salaryText = "(First Clsss Employee :"
        ElseIf Param3 = 2 Then
            salaryText = "(Second Clsss Employee :"
        ElseIf Param3 = 3 Then
            salaryText = "(Third Clsss Employee :"
        ElseIf Param3 = 4 Then
            salaryText = "(Fourth Clsss Employee: "
        Else
            salaryText = "(All Clsss Employee: "
        End If
        
        
        
        If localSalaryType = "R" Then
          salaryText = salaryText + " Employee(s) Salary For the Month of "
           Report14.FakeBasic1.Suppress = True
          Report14.BASIC1.Suppress = False

        ElseIf localSalaryType = "S" Then
          salaryText = salaryText + " Employee(s) Supplementary Salary For the Month of "
          Report14.FakeBasic1.Suppress = True
          Report14.BASIC1.Suppress = False
          
        ElseIf localSalaryType = "D" Then
          salaryText = salaryText + " Employee(s) Dress Allowance For the Month of "
       Else
          salaryText = salaryText + " Employee(s) Festival Bonus For the Month of "
          Report14.FakeBasic1.Suppress = False
          Report14.BASIC1.Suppress = True
       End If
       salaryText = salaryText + GetMonthOftheYear & ")" & "-" & GetSalaryPreparationYaer
       
       
       If localSalaryType = "D" Then
       
          report45.Text28.SetText salaryText
          report45.DiscardSavedData
          Screen.MousePointer = vbHourglass
          CRViewer2.ReportSource = report45
       Else
          Report14.Text28.SetText salaryText
          Report14.DiscardSavedData
          Screen.MousePointer = vbHourglass
          CRViewer2.ReportSource = Report14
       End If
       CRViewer2.ViewReport
       Screen.MousePointer = vbDefault
        
        Set cmd = Nothing
        Set RS = Nothing

Case 19

    
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, EmpIDForTowhom)
        cmd.Parameters.Append Param1 'combo
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, BEGINYEARFORWHOM)
        cmd.Parameters.Append Param2 'combo
        
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, ENDDATEFORWHOM)
        cmd.Parameters.Append Param3 'combo
                  
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_ToWhomitMayConcern(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
       
           
        Dim Department As String
        Dim he As String
        Dim his As String
        
        
        If sGender = "M" Then
          he = "He"
          his = "his"
        Else
          he = "She"
          his = "her"
        End If
        
        Dim helper As String
       If GetFromMonthtoWhom = GetToMonthtoWhom Then
           helper = " has drawn salary and allowances from this hospital for the Month of  " & GetFromMonthtoWhom & " under the following manner :"
       Else
            helper = " has drawn monthly salary and allowances from this hospital during the period of  " & GetFromMonthtoWhom & " to " & GetToMonthtoWhom & " under the following manner :"
       End If
       
       
       If InStr(UCase(DepartemntOfEmp), "SECTION") > 0 Or InStr(UCase(DepartemntOfEmp), "DIVISION") Then
        
                Department = DepartemntOfEmp
        Else
           Department = DepartemntOfEmp + " Department"
        End If

      Dim firstPart As String
      
'      firstPart = "This is to  certify  that, " & EmployeeName & ", " & DesignationOfEmp & ", " & Department & ", Dhaka National Medical Institute Hospital"
           
       firstPart = "This is to  certify  that, " & EmployeeName & " has been working in Dhaka National Medical Institute Hospital, 53/1, Johnson Road, Dhaka since " & Format(DatofJoin, "dd/mm/yyyy") & "."
       If underDepartmentorNot = 0 Then
       
          firstPart = firstPart + "  " & he & " is a " & JobType & " employee  " + " and at present " & LCase(he) & " is working as  " & DesignationOfEmp & " under   " & Department & " in this hospital."
          
       Else
          firstPart = firstPart + "  " & he & " is a " & JobType & " employee  " + " and at present " & LCase(he) & " is working as  " & DesignationOfEmp & " in this hospital."
       
       End If
      If currentFormat = 0 Then
         report15.Text3.SetText firstPart
         report15.Text10.Suppress = False
         If Len(pScale) > 0 Then
            report15.Text10.SetText " " & UCase(Mid(his, 1, 1)) & "" & Mid(his, 2, 2) & " present pay scale is Tk." & pScale & "  and " & he & "" + helper
         Else
            report15.Text10.SetText "As an employee of this hospital s/he" + helper
        End If
      Else
         If underDepartmentorNot = 0 Then
             report15.Text3.SetText "This is to  certify  that, " & EmployeeName & " has been working  as  " & DesignationOfEmp & " under   " & Department & " of the Dhaka National Medical Institute Hospital, 53/1, Johnson Road, Dhaka Since " & Format(DatofJoin, "dd/mm/yyyy") & ".  " & he & " is a " & JobType & " employee of this hospital and  as per  service record  " & his & " date of retirement is  " & Format(DateofRetirement, "dd/mm/yyyy") & "."
         Else
            report15.Text3.SetText "This is to  certify  that, " & EmployeeName & " has been working  as  " & DesignationOfEmp & " in Dhaka National Medical Institute Hospital, 53/1, Johnson Road, Dhaka Since " & Format(DatofJoin, "dd/mm/yyyy") & ".  " & he & " is a " & JobType & " employee of this hospital and  as per  service record  " & his & " date of retirement is  " & Format(DateofRetirement, "dd/mm/yyyy") & "."
         End If
         report15.Text10.Suppress = False
         report15.Text10.SetText "As an employee of this hospital s/he" + helper
      End If
       
       
        
        If twoPage = 1 Then
           report15.GroupHeaderSection4.Suppress = False
        Else
          report15.GroupHeaderSection4.Suppress = True
        End If
        If currentFormat = 0 Then
           report15.DetailSection1.Suppress = True
           report15.GroupTitle2.Suppress = False
           report15.Box1.Suppress = False
           report15.Line2.Suppress = False
           report15.SHOWSUBBASICSUM1.Font.Bold = False
           report15.GroupTitle2.Font.Bold = False
           report15.SHOWSUBBASICSUM1.TopLineStyle = crLSNoLine
           report15.Text5.Suppress = False
           report15.Text6.Suppress = False
           report15.Line3.Suppress = False
           report15.Line4.Suppress = False
           report15.SHOWSUBBASICSUM1.Suppress = True
           report15.SHOWBASICSUM2.Suppress = False
           report15.Text11.Suppress = False
           
        Else
           report15.DetailSection1.Suppress = False
           report15.GroupTitle2.Suppress = True
           report15.Box1.Suppress = True
           report15.Line2.Suppress = True
           ''report15.SHOWSUBBASICSUM1.
           report15.SHOWSUBBASICSUM1.Font.Bold = True
           report15.SHOWSUBBASICSUM1.TopLineStyle = crLSSingleLine
           report15.GroupTitle2.Font.Bold = True
           report15.Text5.Suppress = True
           report15.Text6.Suppress = True
           report15.Line3.Suppress = True
           report15.Line4.Suppress = True
           report15.SHOWSUBBASICSUM1.Suppress = False
           report15.SHOWBASICSUM2.Suppress = True
           report15.Text11.Suppress = True
           
       End If
    
        report15.Database.SetDataSource RS
        report15.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report15
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
Case 20
    
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, Emp_ID_Value_ForLoan)
        cmd.Parameters.Append Param1 'combo
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param2 'combo
        
        
        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param3 'combo
                  
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL  Rpt_Loan_Statment(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        report16.Database.SetDataSource RS

        
        
        'report16.Text11.SetText Right(BEGINYEARFORWHOM, 4) & " - " & Right(ENDDATEFORWHOM, 4) & " ( Ist July /  " & Right(BEGINYEARFORWHOM, 4) & " to 30th June /  " & Right(ENDDATEFORWHOM, 4) & "  )"
        
                
        report16.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report16
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        Emp_ID_Value_ForLoan = ""
        BeginDateForReport = ""
        End_Fiscal_Year = ""


Case 21

    
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        
       
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1 'combo
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2 'combo
                  
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_Scheduleof_Loan_EndOfYear(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        report17.Database.SetDataSource RS

        
        
        report17.Text2.SetText Mid(EnddateforReport, 6, 4)
        
                
        report17.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report17
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
        
        
        
Case 22
        '=============================PROMOTION RECORD REPORT
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText



        Set Param1 = cmd.CreateParameter("param1", adChar, adParamInput, 10, Emp_ID_Value)
        cmd.Parameters.Append Param1 'combo


        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, BeginDateOfIncremnt)
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, EndDateOfIncremnt)
        cmd.Parameters.Append Param3 'combo


        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL Rpt_Promotion_Record(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report18.Database.SetDataSource RS


        report18.Text11.SetText "( From " & Format(BeginDateOfIncremnt, "dd-mmm-yyyy") & " To " & Format(EndDateOfIncremnt, "dd-mmm-yyyy") & " )"
     

        report18.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report18
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        Emp_ID_Value = ""
        BeginDateOfIncremnt = ""
        EndDateOfIncremnt = ""
   
        
        
        
Case 23
        '=============================EMPLOYEE RECORD
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText



        Set Param1 = cmd.CreateParameter("param1", adChar, adParamInput, 100, DEPARMENTNAMEFORTPT)
        cmd.Parameters.Append Param1 'combo


        Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 15, CheckStatusofEmployee)
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 15, SEXFORREPORT)
        cmd.Parameters.Append Param3 'combo

            
        Set Param4 = cmd.CreateParameter("param4", adChar, adParamInput, 100, DESIGNATIONFORRPT)
        cmd.Parameters.Append Param4 'combo
            
        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL Rpt_All_Employee_ID_No(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        REPORT19.Database.SetDataSource RS


        'REPORT19.Text11.SetText "( From " & Format(BeginDateOfIncremnt, "dd-mmm-yyyy") & " To " & Format(EndDateOfIncremnt, "dd-mmm-yyyy") & " )"
        
        If ReportStatusofEmployee = 0 Then
            REPORT19.Text5.SetText "Employee Information " & "( Department wise - Department Name:" & DEPARMENTNAMEFORTPT & ")"
            
        ElseIf ReportStatusofEmployee = 1 Then
            REPORT19.Text5.SetText "Employee(s) Information " & "( All Employee )"
        ElseIf ReportStatusofEmployee = 2 Then
            REPORT19.Text5.SetText "Employee(s) Information " & "( Class wise )"
        ElseIf ReportStatusofEmployee = 3 Then
            REPORT19.Text5.SetText "Employee(s) Information " & "( Sex wise )"
        ElseIf ReportStatusofEmployee = 4 Then
            REPORT19.Text5.SetText "Employee(s) Information " & "( Designation wise- Designation Name:" & DESIGNATIONFORRPT & ")"
        End If
        
        
        REPORT19.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = REPORT19
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        DEPARMENTNAMEFORTPT = ""
        StatusofEmployee = 5
        SEXFORREPORT = 5
        DESIGNATIONFORRPT = ""
          
        
Case 24
        '=============================EMPLOYEE RECORD (RETIRED)
        
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText



        Set Param1 = cmd.CreateParameter("param1", adChar, adParamInput, 100, DEPARMENTNAMEFORTPT)
        cmd.Parameters.Append Param1 'combo


        Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 15, CheckStatusofEmployee)
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adInteger, adParamInput, 15, SEXFORREPORT)
        cmd.Parameters.Append Param3 'combo

            
        Set Param4 = cmd.CreateParameter("param4", adChar, adParamInput, 100, DESIGNATIONFORRPT)
        cmd.Parameters.Append Param4 'combo
            
        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL Rpt_All_Emp_Under_Retierment(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        REPORT20.Database.SetDataSource RS


        'REPORT19.Text11.SetText "( From " & Format(BeginDateOfIncremnt, "dd-mmm-yyyy") & " To " & Format(EndDateOfIncremnt, "dd-mmm-yyyy") & " )"
        
        If ReportStatusofEmployee = 0 Then
           REPORT20.Text8.SetText "Retired Employee(s) Information " & "( Department wise - Department Name:" & DEPARMENTNAMEFORTPT & ")"
            
        ElseIf ReportStatusofEmployee = 1 Then
            REPORT20.Text8.SetText "Retired Employee(s) Information " & "( All Employee )"
        ElseIf ReportStatusofEmployee = 2 Then
            REPORT20.Text8.SetText "Retired Employee(s) Information " & "( Class wise )"
        ElseIf ReportStatusofEmployee = 3 Then
            REPORT20.Text8.SetText "Retired Employee(s) Information " & "( Sex wise )"
        ElseIf ReportStatusofEmployee = 4 Then
            REPORT20.Text8.SetText "Retired Employee(s) Information " & "( Designation wise- Designation Name:" & DESIGNATIONFORRPT & ")"
        End If
        
        
        REPORT20.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = REPORT20
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        DEPARMENTNAMEFORTPT = ""
        StatusofEmployee = 5
        SEXFORREPORT = 5
        DESIGNATIONFORRPT = ""
        
Case 25

        Screen.MousePointer = vbHourglass
     
        con.Open strCN.Connection_String
        
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1 'combo
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2 'combo
                  

        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, GetSalaryPreparationYaer)
        cmd.Parameters.Append Param3 'combo
        
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 20, GetMonthOftheYear)
        cmd.Parameters.Append Param4 'combo
        
        
        cmd.Properties("PLSQLRSet") = True
        cmd.CommandText = "{CALL Rpt_Salary_Prep_Individual(?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
        
        report21.Database.SetDataSource RS
        
        report21.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report21
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        Emp_ID_Value = ""
        EnddateforReport = ""

Case 26
  
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

         Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1 'combo

        Set Param2 = cmd.CreateParameter("Param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 15, Trim(GetMonthOftheYear))
        cmd.Parameters.Append Param3

        Set Param4 = cmd.CreateParameter("Param4", adVarChar, adParamInput, 15, GetSalaryPreparationYaer)
        cmd.Parameters.Append Param4

        Set Param5 = cmd.CreateParameter("Param5", adVarChar, adParamInput, 15, Trim(StatusofEmployee))
        cmd.Parameters.Append Param5


        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL Rpt_Bonus_Preaparation(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False



        report22.Database.SetDataSource RS


        report22.Text47.SetText "(" & GetMonthOftheYear & "-" & GetSalaryPreparationYaer & " )"

        If StatusofEmployee = "1" Then
            report22.Text46.SetText "First Class"
        ElseIf StatusofEmployee = "2" Then
            report22.Text46.SetText "Second Class"
        ElseIf StatusofEmployee = "3" Then
            report22.Text46.SetText " Third Class"
        Else
            report22.Text46.SetText "Fourth Class"
        End If

        report22.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report22
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing

Case 27
         Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
         Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 15, GetMonthOftheYear)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("Param2", adVarChar, adParamInput, 15, GetSalaryPreparationYaer)
        cmd.Parameters.Append Param2
        
        Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 15, Trim(StatusofEmployee))
        cmd.Parameters.Append Param3
                
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Bonus_sendto_Bank(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False


                
        report23.Database.SetDataSource RS

        
        report23.Text1.SetText "(" & GetMonthOftheYear & "-" & GetSalaryPreparationYaer & " )"
        
        If StatusofEmployee = "1" Then
            report23.Text9.SetText ":First Class"
        ElseIf StatusofEmployee = "2" Then
            report23.Text9.SetText ":Second Class"
        ElseIf StatusofEmployee = "3" Then
            report23.Text9.SetText " :Third Class"
        Else
            report23.Text9.SetText ":Fourth Class"
        End If
        
        report23.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report23
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing


Case 28

        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, GetMonthOftheYear)
        cmd.Parameters.Append Param1 'combo
        
        Set Param2 = cmd.CreateParameter("Param2", adVarChar, adParamInput, 100, Trim(frmPaySlipInfoReport.cboYear))
        cmd.Parameters.Append Param2 'combo
        
        Set Param3 = cmd.CreateParameter("Param3", adVarChar, adParamInput, 100, Emp_ID_Value)
        cmd.Parameters.Append Param3 'combo
        
        
        cmd.Properties("PLSQLRSet") = True
        
        cmd.CommandText = "{CALL Rpt_BonusSlip_Pre(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False
                
        report24.Database.SetDataSource RS
        
        report24.Text103.SetText "Pay Slip(Bonus) For the Month  " & GetMonthOftheYear & " " & Trim(frmPaySlipInfoReport.cboYear) & " :- "
        'report24.Text88.SetText "Pay Slip(Bonus) For the Month  " & GetMonthOftheYear & " " & Trim(frmPaySlipInfoReport.cboYear) & " :- "
        
        report24.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report24
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing

'Case 29
'
'
'        Screen.MousePointer = vbHourglass
'        con.Open strCN.Connection_String
'        Set cmd.ActiveConnection = con
'        cmd.CommandType = adCmdText
'
'
'
'        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
'        cmd.Parameters.Append Param1 'combo
'
'
'        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
'        cmd.Parameters.Append Param2 'combo
'
'        cmd.Properties("PLSQLRSet") = True
'
'        cmd.CommandText = "{CALL RPT_SUMM_SCHE_LOAN_ENDOFYEAR(?,?)}"
'        Set RS = cmd.Execute
'        cmd.Properties("PLSQLRSet") = False
'
'        report26.Database.SetDataSource RS
'
'
'
'        report26.Text2.SetText Mid(EnddateforReport, 6, 4)
'
'
'        report26.DiscardSavedData
'        Screen.MousePointer = vbHourglass
'        CRViewer2.ReportSource = report26
'        CRViewer2.ViewReport
'        Screen.MousePointer = vbDefault
'        Set cmd = Nothing
'        Set RS = Nothing
'        BeginDateForReport = ""
'        End_Fiscal_Year = ""
        
Case 30


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText



        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1 'combo


        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2 'combo

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_INCOME_EXPENCE(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report28.Database.SetDataSource RS



        'report28.Text4.SetText " From  " & frmGratuityReport.Begin_Fiscal_date.Value & " To  " & frmGratuityReport.End_Fiscal_Year
        report28.Text4.SetText "From  " & BeginDateForReport & "  To  " & EnddateforReport


        report28.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report28
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
        
Case 31


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText



        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1 'combo


        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2 'combo

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_CASHBOOK(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report29.Database.SetDataSource RS

        'report29.Text13.SetText Mid(EnddateforReport, 6, 4)
        report29.Text13.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report29.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report29
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
        
Case 32


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText



       ' Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
       ' cmd.Parameters.Append Param1 'combo


       ' Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        'cmd.Parameters.Append Param2 'combo

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_OPEN_CLOSE_ENDOFYEAR()}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report30.Database.SetDataSource RS
        'report30.Text2.SetText Mid(EnddateforReport, 6, 4)
        report30.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report30
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
Case 33


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText


        Set Param5 = cmd.CreateParameter("Param5", adChar, adParamInput, 15, sBankCode)
        cmd.Parameters.Append Param5 'combo
        
        Set Param6 = cmd.CreateParameter("Param6", adChar, adParamInput, 15, sAccountType)
        cmd.Parameters.Append Param6 'combo
        
        Set Param7 = cmd.CreateParameter("Param7", adChar, adParamInput, 15, sAccountNo)
        cmd.Parameters.Append Param7 'combo

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_BANK_ACCO_STATEMENT(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report31.Database.SetDataSource RS

        report31.Text6.SetText "From : " & BeginDateForReport & "   To : " & EnddateforReport
        report31.Text7.SetText "Bank Name : " & sBankName
        report31.Text8.SetText "Account Type : " & sAccountTypeName
        report31.Text9.SetText "Account No : " & sAccountNo

        report31.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report31
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
Case 34


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText


        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL GRATUITY_BALANCE_SHEET(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report32.Database.SetDataSource RS

        report32.Text7.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report32.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report32
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
Case 35


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        Set SourceId = cmd.CreateParameter("SourceId", adVarChar, adParamInput, 15, sSourceId)
        cmd.Parameters.Append SourceId

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_REC_SOURCE_OF_GRATUITY(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report35.Database.SetDataSource RS

        report35.Text11.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport
        report35.Text12.SetText "Source of Gratuity : " & sSourceName
        report35.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report35
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
Case 36
        
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set PurposeId = cmd.CreateParameter("PurposeId", adVarChar, adParamInput, 15, sPurposeId)
        cmd.Parameters.Append PurposeId

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_PUR_OF_PAYMENT_WISE(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report38.Database.SetDataSource RS

        report38.Text13.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport
        report38.Text12.SetText "Purpose of Payment : " & sPurposeName
        report38.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report38
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
Case 37
        
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set BankCode = cmd.CreateParameter("BankCode", adVarChar, adParamInput, 15, sBankCode)
        cmd.Parameters.Append BankCode

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_REC_BANK_WISE(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report39.Database.SetDataSource RS

        report39.Text13.SetText "Bank Name : " & sBankName
        report39.Text14.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport
        report39.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report39
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
Case 38
        
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        
        Set BankCode = cmd.CreateParameter("BankCode", adVarChar, adParamInput, 15, sBankCode)
        cmd.Parameters.Append BankCode

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_GRA_Payment_BANK_WISE(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report40.Database.SetDataSource RS

        report40.Text13.SetText "Bank Name : " & sBankName
        report40.Text14.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport
        report40.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report40
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
Case 40


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText


        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_PF_INCOME_EXPENCE(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report33.Database.SetDataSource RS

        report33.Text5.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report33.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report33
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
Case 41


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText


        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_PF_CASHBOOK(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report34.Database.SetDataSource RS

        report34.Text5.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report34.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report34
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
 
Case 42


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText


        Set Param5 = cmd.CreateParameter("Param5", adChar, adParamInput, 15, sBankCode)
        cmd.Parameters.Append Param5 'combo
        
        Set Param6 = cmd.CreateParameter("Param6", adChar, adParamInput, 15, sAccountType)
        cmd.Parameters.Append Param6 'combo
        
        Set Param7 = cmd.CreateParameter("Param7", adChar, adParamInput, 15, sAccountNo)
        cmd.Parameters.Append Param7 'combo

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_PF_BANK_ACCO_STATEMENT(?,?,?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report37.Database.SetDataSource RS

        report37.Text7.SetText "From : " & BeginDateForReport & "   To : " & EnddateforReport
        report37.Text8.SetText "Bank Name : " & sBankName
        report37.Text9.SetText "Account Type : " & sAccountTypeName
        report37.Text10.SetText "Account No : " & sAccountNo

        report37.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report37
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
 
 
Case 43


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText


        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL PF_BALANCE_SHEET(?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report36.Database.SetDataSource RS

        'report36.Text5.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report36.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report36
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
Case 44


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        Set SourceId = cmd.CreateParameter("SourceId", adVarChar, adParamInput, 15, sSourceId)
        cmd.Parameters.Append SourceId

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_PFRECEIVE_SOURCE_OF_FUND(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report41.Database.SetDataSource RS
        report41.Text13.SetText "Source of Fund : " & sSourceName
        report41.Text14.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report41.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report41
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
          Case 45


        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        Set PurposeId = cmd.CreateParameter("PurposeId", adVarChar, adParamInput, 15, sPurposeId)
        cmd.Parameters.Append PurposeId

        Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 15, BeginDateForReport)
        cmd.Parameters.Append Param1
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, EnddateforReport)
        cmd.Parameters.Append Param2

        cmd.Properties("PLSQLRSet") = True

        cmd.CommandText = "{CALL RPT_PF_PURPOSE_OF_PAYMENT_WISE(?,?,?)}"
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report42.Database.SetDataSource RS
        report42.Text15.SetText "Payment Purpose : " & sPurposeName
        report42.Text16.SetText "From : " & BeginDateForReport & "  To : " & EnddateforReport

        report42.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report42
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
        BeginDateForReport = ""
        End_Fiscal_Year = ""
 
Case 46
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'combo
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, CDate(GetFromMonthtoWhom))
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, CDate(GetToMonthtoWhom))
        cmd.Parameters.Append Param3 'combo
        
         Set Param4 = cmd.CreateParameter("param4", adChar, adParamInput, 5, " ")
        cmd.Parameters.Append Param4 'department




       '
        cmd.Properties("PLSQLRSet") = True
     
       cmd.CommandText = "{CALL RPT_SALARY_SUMMARY(?,?,?,?)}"
       
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report43.Database.SetDataSource RS
        
       If BEGINYEARFORWHOM = ENDDATEFORWHOM Then
         report43.Text16.SetText (BEGINYEARFORWHOM)
       Else
         report43.Text16.SetText (BEGINYEARFORWHOM + "     To    " + ENDDATEFORWHOM)
       End If
        
        report43.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report43
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
'        BeginDateForReport = ""
'        End_Fiscal_Year = ""
      
      
      
      
Case 47
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, currentOption)
        cmd.Parameters.Append Param1 'combo
        
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, CDate(GetFromMonthtoWhom))
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, CDate(GetToMonthtoWhom))
        cmd.Parameters.Append Param3 'combo
       
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 40, currentDept)
        cmd.Parameters.Append Param4 'Department


        cmd.Properties("PLSQLRSet") = True
     
       cmd.CommandText = "{CALL RPT_SALARY_SUMMARY(?,?,?,?)}"
       
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report44.Database.SetDataSource RS
        
       If BEGINYEARFORWHOM = ENDDATEFORWHOM Then
         report44.Text26.SetText (BEGINYEARFORWHOM)
       Else
         report44.Text26.SetText (BEGINYEARFORWHOM + "     To    " + ENDDATEFORWHOM)
       End If
       If currentOption = 3 Then
          Dim helper1 As String
          helper1 = "Salary Disbursement Break-down for the department of : " + currentDept
          report44.upperDEPTNM1.Font.Bold = False
          report44.DISBURSEE1.Suppress = True
          report44.Line2.Suppress = True
          report44.Text1.Suppress = True
          report44.Text20.SetText ("Name of Disbursees")
          report44.Text20.Width = "2850"
          report44.Text27.Suppress = False
          report44.Text27.SetText (helper)
       Else
          report44.Text20.Width = "1965"
          report44.DISBURSEE1.Suppress = False
          report44.upperDEPTNM1.Font.Bold = True
          report44.Line2.Suppress = False
          report44.Text20.SetText ("Department")
          report44.Text1.Suppress = False
          report44.Text27.Suppress = True
       End If
        report44.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report44
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
'        BeginDateForReport = ""
'        End_Fiscal_Year = ""
      
 Case 48
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'mode
        
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, paramMonth)
        cmd.Parameters.Append Param2 'Month


        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 15, paramYear)
        cmd.Parameters.Append Param3 'Year
        
         



       '
        cmd.Properties("PLSQLRSet") = True
     
       cmd.CommandText = "{CALL rpt_payroll_statistics(?,?,?)}"
       
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report48.Database.SetDataSource RS
        
'       If BEGINYEARFORWHOM = ENDDATEFORWHOM Then
'         report43.Text16.SetText (BEGINYEARFORWHOM)
'       Else
'         report43.Text16.SetText (BEGINYEARFORWHOM + "     To    " + ENDDATEFORWHOM)
'       End If
        report48.Text8.SetText ("As per salary disbursement of " + paramMonth + ", " + paramYear)
        report48.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report48
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
'        BeginDateForReport = ""
'        End_Fiscal_Year = ""
Case 49
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, 1)
        cmd.Parameters.Append Param1 'mode
        
        
        Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, paramMonth)
        cmd.Parameters.Append Param2 'Month


        Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 15, paramYear)
        cmd.Parameters.Append Param3 'Year
        
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 20, paramDepartment)
        cmd.Parameters.Append Param4 'department
       
     


       '
        cmd.Properties("PLSQLRSet") = True
     
       cmd.CommandText = "{CALL rpt_payroll_staff(?,?,?,?)}"
       
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        report49.Database.SetDataSource RS
        

        report49.Text10.SetText ("As per salary disbursement of " + paramMonth + ", " + paramYear)
        report49.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = report49
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
'        BeginDateForReport = ""
'        End_Fiscal_Year = ""
Case 50
      
        Screen.MousePointer = vbHourglass
        con.Open strCN.Connection_String
        Set cmd.ActiveConnection = con
        cmd.CommandType = adCmdText

        
        Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 5, currentOption)
        cmd.Parameters.Append Param1 'combo
        
        
        
        Set Param2 = cmd.CreateParameter("param2", adDate, adParamInput, 15, frmITCalc.BeginDatePicker.Value)
        cmd.Parameters.Append Param2 'combo


        Set Param3 = cmd.CreateParameter("param3", adDate, adParamInput, 15, frmITCalc.EndDatePicker.Value)
        cmd.Parameters.Append Param3 'combo
       
        Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 40, frmITCalc.employeeIDComboBox.Text)
        cmd.Parameters.Append Param4 'Department


        cmd.Properties("PLSQLRSet") = True
     
       cmd.CommandText = "{CALL rpt_Income_tax_calc(?,?,?,?)}"
       
        Set RS = cmd.Execute
        cmd.Properties("PLSQLRSet") = False

        reportIT6.Database.SetDataSource RS
        reportIT6.RptTxtSex.SetText (sGender)
        
'       If BEGINYEARFORWHOM = ENDDATEFORWHOM Then
'         report44.Text26.SetText (BEGINYEARFORWHOM)
'       Else
'         report44.Text26.SetText (BEGINYEARFORWHOM + "     To    " + ENDDATEFORWHOM)
'       End If
'       If currentOption = 3 Then
'          Dim helper1 As String
'          helper1 = "Salary Disbursement Break-down for the department of : " + currentDept
'          report44.upperDEPTNM1.Font.Bold = False
'          report44.DISBURSEE1.Suppress = True
'          report44.Line2.Suppress = True
'          report44.Text1.Suppress = True
'          report44.Text20.SetText ("Name of Disbursees")
'          report44.Text20.Width = "2850"
'          report44.Text27.Suppress = False
'          report44.Text27.SetText (helper)
'       Else
'          report44.Text20.Width = "1965"
'          report44.DISBURSEE1.Suppress = False
'          report44.upperDEPTNM1.Font.Bold = True
'          report44.Line2.Suppress = False
'          report44.Text20.SetText ("Department")
'          report44.Text1.Suppress = False
'          report44.Text27.Suppress = True
'       End If
        reportIT6.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer2.ReportSource = reportIT6
        CRViewer2.ViewReport
        Screen.MousePointer = vbDefault
        Set cmd = Nothing
        Set RS = Nothing
'        BeginDateForReport = ""
'        End_Fiscal_Year = ""
 
End Select
Exit Sub
Errdes:
    MsgBox "Wrong Opeation !!! Try once again.......", vbExclamation, "IT Division, DNMIH"
    Screen.MousePointer = vbDefault
End Sub

