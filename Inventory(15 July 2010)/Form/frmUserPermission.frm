VERSION 5.00
Begin VB.Form frmUserPermission 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Permision"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      Height          =   3570
      Index           =   1
      Left            =   3975
      TabIndex        =   9
      Top             =   2400
      Width           =   3540
   End
   Begin VB.ListBox lst 
      Height          =   3180
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   2400
      Width           =   3540
   End
   Begin VB.CommandButton cmdView 
      Height          =   285
      Index           =   0
      Left            =   2775
      Picture         =   "frmUserPermission.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   1500
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   7
      Top             =   450
      Width           =   5115
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   1500
      TabIndex        =   6
      Top             =   825
      Width           =   5115
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   1500
      TabIndex        =   5
      Top             =   1200
      Width           =   5115
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   1500
      TabIndex        =   4
      Top             =   1575
      Width           =   5115
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   2
      Top             =   525
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "User ID."
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   0
      Top             =   225
      Width           =   585
   End
End
Attribute VB_Name = "frmUserPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
Dim strMenuSL() As String
Dim strMenuCaption() As String
Dim intTotalRecord As Integer
Private Sub Get_MenuInfo()
Dim intMenuSLLen As Integer
Dim intIncrementval As Integer
Set objRs = objcom.Get_RS("SELECT MenuSL, ManuCaption From MenuInfo ORDER BY MenuSL", objmyCon)
With objRs
    If Not .EOF Then
        .MoveFirst
        intTotalRecord = .RecordCount
        ReDim strMenuSL(intTotalRecord)
        ReDim strMenuCaption(intTotalRecord)
        
        Do Until .EOF
            strMenuSL(intIncrementval) = Trim(objRs(0))
            strMenuCaption(intIncrementval) = Trim(objRs(1))
            intMenuSLLen = Len(Trim(objRs(0)))
             Select Case intMenuSLLen
                Case Is > 2
                    lst(0).AddItem Space(intMenuSLLen) & Trim(objRs(1))
                Case Else
                    lst(0).AddItem Trim(objRs(1))
             End Select
             intIncrementval = intIncrementval + 1
            .MoveNext
        Loop
    End If
End With

Set objRs = Nothing
End Sub
Private Function Check_Dupilicate(strItem As String) As Boolean
Dim intTotalItem As Integer
Check_Dupilicate = False
For intTotalItem = 0 To lst(1).ListCount
    If strItem = Trim(lst(1).List(intTotalItem)) Then
        Check_Dupilicate = True
        Exit For
    End If
Next
End Function

Private Function Get_ParrentManu(strMenu As String) As String()
Dim intposition As Integer
Dim intposition1 As Integer
Dim intposition2 As Integer
Dim intposition3 As Integer
Dim intposition4 As Integer

Dim strCapation(5) As String


Select Case Len(strMenu)
   Case Is > 2
   

        For intposition = 0 To intTotalRecord
           If strMenuSL(intposition) = Mid(Trim(strMenu), 1, Len(Trim(strMenu)) - 2) Then
                strCapation(1) = strMenuCaption(intposition)
            If Len(strMenuSL(intposition)) > 2 Then
               For intposition1 = 0 To intTotalRecord
                    If strMenuSL(intposition1) = Mid(Trim(strMenu), 1, Len(Trim(strMenu)) - 4) Then
                        strCapation(2) = strMenuCaption(intposition1)
                        Exit For
                    End If
               Next
            End If
            Exit For
           End If
        Next
        Get_ParrentManu = strCapation
End Select
End Function

Private Sub cmdView_Click(Index As Integer)
Select Case Index
    Case 0
        Set objRs = objcom.Get_RS("SELECT EmpInfo.EmpCode, EmpInfo.EmpName, DesigInfo.DesigName, " _
                                & "ProjectInfo.ProjectName, DeptInfo.DeptName FROM EmpInfo " _
                                & "INNER JOIN ProjectInfo ON EmpInfo.FactoryCode = ProjectInfo.ProjectCode " _
                                & "INNER JOIN DeptInfo ON EmpInfo.DeptCode = DeptInfo.DeptCode " _
                                & "INNER JOIN DesigInfo ON EmpInfo.DesigCode = DesigInfo.DesigCode " _
                                & "ORDER BY EmpInfo.EmpCode", objmyCon)
        Dim frmfindform As New frmFind
        Set frmfindform.objFindRS = objRs
        Set objRs = Nothing
        Set frmfindform.OwnerForm = Me
        frmfindform.intInputsel = 0
        frmfindform.Show 1
        txtFields(0).SetFocus
End Select
End Sub

Private Sub Form_Load()
Get_MenuInfo
End Sub

Private Sub lst_DblClick(Index As Integer)
Dim i As Integer
Select Case Index
    Case 0
        'If Check_Dupilicate(Trim(lst(0).Text)) = False Then
            
            MsgBox Get_ParrentManu(strMenuSL(lst(0).ListIndex))(1)
         '   lst(1).AddItem Trim(lst(0).Text)
        'Else
            'MsgBox "This Permission Already Given.", vbInformation + vbOKOnly, strmsgtitle
        'End If
End Select
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Select Case Index
        Case 0
            txtFields(0) = Format(Trim(txtFields(0)), "000000")
            If Len(Trim(txtFields(1))) <> 0 Then
                Set objRs = objcom.Get_RS("SELECT EmpCode From EmpInfo WHERE EmpCode ='" & Trim(txtFields(0)) & "'", objmyCon)
                If Not objRs.EOF Then
                    lst(0).SetFocus
                Else
                    cmdView_Click (0)
                End If
            Else
                Set objRs = objcomfam.Get_RS("SELECT EmpCode From EmpInfo WHERE EmpCode ='" & Trim(txtFields(0)) & "'", objmyCon)
                If Not objRs.EOF Then
                    txtFields(2).SetFocus
                Else
                    cmdView_Click (0)
                End If
            
            End If
            lst(0).SetFocus
    End Select
End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Select Case Index
    Case 0
       Set objRs = objcom.Get_RS("SELECT EmpInfo.EmpName, DesigInfo.DesigName, " _
                        & "DeptInfo.DeptName, ProjectInfo.ProjectName, EmpInfo.EmpCode " _
                        & "FROM EmpInfo " _
                        & "INNER JOIN ProjectInfo ON EmpInfo.FactoryCode = ProjectInfo.ProjectCode " _
                        & "INNER JOIN DeptInfo ON EmpInfo.DeptCode = DeptInfo.DeptCode " _
                        & "INNER JOIN DesigInfo ON EmpInfo.DesigCode = DesigInfo.DesigCode " _
                        & "WHERE EmpInfo.EmpCode = '" & Trim(txtFields(0)) & "'", objmyCon)
                                
        If Not objRs.EOF Then
            lblCaption(0) = objRs(0)
            lblCaption(1) = objRs(1)
            lblCaption(2) = objRs(2)
            lblCaption(3) = objRs(3)
        Else
            lblCaption(0) = ""
            lblCaption(1) = ""
            lblCaption(2) = ""
            lblCaption(3) = ""
        End If
End Select
End Sub
