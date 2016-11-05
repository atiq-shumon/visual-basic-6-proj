VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rtpViewer 
   Caption         =   "Viewer"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   ScaleHeight     =   5565
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   5385
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6735
      lastProp        =   500
      _cx             =   11880
      _cy             =   9499
      DisplayGroupTree=   -1  'True
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
Attribute VB_Name = "rtpViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rptShowAllUser As New CrystalReport1
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim RS As New ADODB.Recordset
Private connstring As New CLSCONNECTION
Private param1 As New ADODB.Parameter
Private security As New clsSecurity
Private Sub Form_Load()
 CRViewer91.Zoom 100
    If mode = 1 Then
       Dim param1 As New ADODB.Parameter
       conn.Open connstring.ConnectionString
       Set cmd.ActiveConnection = conn
       cmd.CommandType = adCmdText
       cmd.Properties("PLSQLRSet") = True
       cmd.CommandText = "{Call rptShowAllUser}"
       
       Set RS = cmd.Execute
       cmd.Properties("PLSQLRSet") = False
       rptShowAllUser.Database.SetDataSource RS
       rptShowAllUser.DiscardSavedData
       CRViewer91.ReportSource = rptShowAllUser
       CRViewer91.ViewReport
      
      
    End If
End Sub

