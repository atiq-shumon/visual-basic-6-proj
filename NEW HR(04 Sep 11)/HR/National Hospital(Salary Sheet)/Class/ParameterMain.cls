VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParameterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarEFFDATE As Date 'local copy
Private mvarEMPCONTRPF As Double 'local copy
Private mvarEMRCONTRPF As Double 'local copy
Private mvarYEARLYINTOFPF As Double 'local copy
Private mvarPFINCOMEINVESTDIST As Double 'local copy
Private mvarCUYRTOTALCONTRIBUTION As Double 'local copy
Private mvarMAXSEUPCOM As Double 'local copy
Private mvarMAXSEUPWOM As Double 'local copy
Private mvarMAXWHINMONTH As Double 'local copy
Private mvarOTPAYPHR As Double 'local copy
Private mvarHOLIDAYOT As Double 'local copy
Private mvarHOLIDAYOTHOUR As Double 'local copy
Private mvarSEAWRSALINOFFSE As Double 'local copy
Private mvarHOUSERENTUNWCOM As Double 'local copy
Private mvarDAUNDERPCOM As Double 'local copy
Private mvarDAUNDERWCOM As Double 'local copy
Private mvarConnstring As String 'local copy
Private mvarSPEASSUPCOM As Double 'local copy
Private mvarSPEASSUWCOM As Double 'local copy
Public Property Let SPEASSUWCOM(ByVal vData As Double)
    mvarSPEASSUWCOM = vData
End Property
Public Property Get SPEASSUWCOM() As Double
    SPEASSUWCOM = mvarSPEASSUWCOM
End Property
Public Property Let SPEASSUPCOM(ByVal vData As Double)
    mvarSPEASSUPCOM = vData
End Property
Public Property Get SPEASSUPCOM() As Double
    SPEASSUPCOM = mvarSPEASSUPCOM
End Property
Public Property Let Connstring(ByVal vData As String)
    mvarConnstring = vData
End Property
Public Property Get Connstring() As String
    Connstring = mvarConnstring
End Property
Public Property Let DAUNDERWCOM(ByVal vData As Double)
    mvarDAUNDERWCOM = vData
End Property
Public Property Get DAUNDERWCOM() As Double
    DAUNDERWCOM = mvarDAUNDERWCOM
End Property
Public Property Let DAUNDERPCOM(ByVal vData As Double)
    mvarDAUNDERPCOM = vData
End Property
Public Property Get DAUNDERPCOM() As Double
    DAUNDERPCOM = mvarDAUNDERPCOM
End Property
Public Property Let HOUSERENTUNWCOM(ByVal vData As Double)
    mvarHOUSERENTUNWCOM = vData
End Property
Public Property Get HOUSERENTUNWCOM() As Double
    HOUSERENTUNWCOM = mvarHOUSERENTUNWCOM
End Property
Public Property Let SEAWRSALINOFFSE(ByVal vData As Double)
    mvarSEAWRSALINOFFSE = vData
End Property
Public Property Get SEAWRSALINOFFSE() As Double
    SEAWRSALINOFFSE = mvarSEAWRSALINOFFSE
End Property
Public Property Let HOLIDAYOTHOUR(ByVal vData As Double)
    mvarHOLIDAYOTHOUR = vData
End Property
Public Property Get HOLIDAYOTHOUR() As Double
    HOLIDAYOTHOUR = mvarHOLIDAYOTHOUR
End Property
Public Property Let HOLIDAYOT(ByVal vData As Double)
    mvarHOLIDAYOT = vData
End Property
Public Property Get HOLIDAYOT() As Double
    HOLIDAYOT = mvarHOLIDAYOT
End Property
Public Property Let OTPAYPHR(ByVal vData As Double)
    mvarOTPAYPHR = vData
End Property
Public Property Get OTPAYPHR() As Double
    OTPAYPHR = mvarOTPAYPHR
End Property
Public Property Let MAXWHINMONTH(ByVal vData As Double)
    mvarMAXWHINMONTH = vData
End Property
Public Property Get MAXWHINMONTH() As Double
    MAXWHINMONTH = mvarMAXWHINMONTH
End Property
Public Property Let MAXSEUPWOM(ByVal vData As Double)
    mvarMAXSEUPWOM = vData
End Property
Public Property Get MAXSEUPWOM() As Double
    MAXSEUPWOM = mvarMAXSEUPWOM
End Property
Public Property Let MAXSEUPCOM(ByVal vData As Double)
    mvarMAXSEUPCOM = vData
End Property
Public Property Get MAXSEUPCOM() As Double
    MAXSEUPCOM = mvarMAXSEUPCOM
End Property
Public Property Let CUYRTOTALCONTRIBUTION(ByVal vData As Double)
    mvarCUYRTOTALCONTRIBUTION = vData
End Property
Public Property Get CUYRTOTALCONTRIBUTION() As Double
    CUYRTOTALCONTRIBUTION = mvarCUYRTOTALCONTRIBUTION
End Property
Public Property Let PFINCOMEINVESTDIST(ByVal vData As Double)
    mvarPFINCOMEINVESTDIST = vData
End Property
Public Property Get PFINCOMEINVESTDIST() As Double
    PFINCOMEINVESTDIST = mvarPFINCOMEINVESTDIST
End Property
Public Property Let YEARLYINTOFPF(ByVal vData As Double)
    mvarYEARLYINTOFPF = vData
End Property
Public Property Get YEARLYINTOFPF() As Double
    YEARLYINTOFPF = mvarYEARLYINTOFPF
End Property
Public Property Let EMRCONTRPF(ByVal vData As Double)
    mvarEMRCONTRPF = vData
End Property
Public Property Get EMRCONTRPF() As Double
    EMRCONTRPF = mvarEMRCONTRPF
End Property
Public Property Let EMPCONTRPF(ByVal vData As Double)
    mvarEMPCONTRPF = vData
End Property
Public Property Get EMPCONTRPF() As Double
    EMPCONTRPF = mvarEMPCONTRPF
End Property
Public Property Let EFFDATE(ByVal vData As Date)
    mvarEFFDATE = vData
End Property
Public Property Get EFFDATE() As Date
    EFFDATE = mvarEFFDATE
End Property
Public Function Save()
  ' If mvarEmp_Id = Empty Then Exit Function

    Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New ADODB.Parameter
    Dim Param2 As New ADODB.Parameter
    Dim Param3 As New ADODB.Parameter
    Dim Param4 As New ADODB.Parameter
    Dim Param5 As New ADODB.Parameter
    Dim Param6 As New ADODB.Parameter
    Dim Param7 As New ADODB.Parameter
    Dim Param8 As New ADODB.Parameter
    Dim Param9 As New ADODB.Parameter
    Dim Param10 As New ADODB.Parameter
    Dim Param11 As New ADODB.Parameter
    Dim Param12 As New ADODB.Parameter
    Dim Param13 As New ADODB.Parameter
    Dim Param14 As New ADODB.Parameter
    Dim Param15 As New ADODB.Parameter
    Dim Param16 As New ADODB.Parameter
    
        
    conn.Open mvarConnstring
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adDate, adParamInput, 10, mvarEFFDATE)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 5, mvarEMPCONTRPF)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 5, mvarEMRCONTRPF)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 5, mvarYEARLYINTOFPF)
    cmd.Parameters.Append Param4

    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 5, mvarPFINCOMEINVESTDIST)
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 5, mvarCUYRTOTALCONTRIBUTION)
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 5, mvarMAXSEUPCOM)
    cmd.Parameters.Append Param7

    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 5, mvarMAXSEUPWOM)
    cmd.Parameters.Append Param8
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 5, mvarMAXWHINMONTH)
    cmd.Parameters.Append Param9
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 5, mvarOTPAYPHR)
    cmd.Parameters.Append Param10

    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 5, mvarHOLIDAYOT)
    cmd.Parameters.Append Param11
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 5, mvarHOLIDAYOTHOUR)
    cmd.Parameters.Append Param12
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 5, mvarSEAWRSALINOFFSE)
    cmd.Parameters.Append Param13
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 5, mvarHOUSERENTUNWCOM)
    cmd.Parameters.Append Param14
    
    Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 5, mvarDAUNDERPCOM)
    cmd.Parameters.Append Param15
    
    Set Param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 5, mvarDAUNDERWCOM)
    cmd.Parameters.Append Param16
    
    Set Param17 = cmd.CreateParameter("param15", adDouble, adParamInput, 5, mvarSPEASSUPCOM)
    cmd.Parameters.Append Param17
    
    Set Param18 = cmd.CreateParameter("param16", adDouble, adParamInput, 5, mvarSPEASSUWCOM)
    cmd.Parameters.Append Param18
    
    '----------------------------------------------------------------------------------

' Enable PLSQLRSet property
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Parameter_setup_save(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ,?,?)}"
    
    Set RS = cmd.Execute
    
' Disable PLSQLRSet property
cmd.Properties("PLSQLRSet") = False
End Function

