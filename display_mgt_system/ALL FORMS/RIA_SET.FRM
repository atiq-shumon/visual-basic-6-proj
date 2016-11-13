VERSION 4.00
Begin VB.Form RIA_SET 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "UPTAKE SETTINGS"
   ClientHeight    =   5265
   ClientLeft      =   1425
   ClientTop       =   2625
   ClientWidth     =   8610
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   5955
   Left            =   1365
   LinkTopic       =   "Form2"
   ScaleHeight     =   5265
   ScaleWidth      =   8610
   Top             =   1995
   Width           =   8730
   Begin VB.CommandButton Command3 
      Caption         =   "BACK TO CALIBRATION"
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   840
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Text            =   "1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Text            =   "1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Text            =   "1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ACCEPT ALL ABOVE SETTINGS?"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   4320
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   2640
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "High voltage:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type channel width:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type level:        or"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "High Voltage"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ISOTOPE:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type channel width:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type level:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select source"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Menu iso_list 
      Caption         =   "Source List"
      Begin VB.Menu add 
         Caption         =   "Add"
      End
      Begin VB.Menu remove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "RIA_SET"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub add_Click()
      Open "c:\ria\chaninfo.rib" For Append As #1
         Print #1, Text4, Text
         Print #1, Text1.Text
         Print #1, Text2.Text
         Print #1, Text3.Text
         Print #1, Text5.Text
         Print #1, Text6.Text
         Print #1, Text7.Text
      Close #1
         Combo1.AddItem Text4.Text
Open "c:\ria\chaninfo.rib" For Input As #1
  Do
      Input #1, Xisotope$(mm%)
      Input #1, Xlevel%(mm%)
      Input #1, Xchwidth%(mm%)
      Input #1, Xhivo%(mm%)
      Input #1, Xllevel%(mm%)
      Input #1, Xlchwidth%(mm%)
      Input #1, Xlhivo%(mm%)
   'recalculation: Volts to byte value  #####
        'Xhivo%(mm%) = Int(Xhivo%(mm%) / 4.7)
       ' XLhivo%(mm%) = Int(XLhivo%(mm%) / 4.7)
      mm% = mm% + 1
      mmax% = mm%
  Loop While Not EOF(1)
Close #1
         
End Sub

Private Sub Combo1_Click()
      mm% = Combo1.ListIndex
      Text1.Text = Str$(Xlevel%(mm%))
      Text2.Text = Str$(Xchwidth%(mm%))
      Text3.Text = Str$(Xhivo%(mm%))
      Text4.Text = Xisotope$(mm%)
      Text5.Text = Str$(Xllevel%(mm%))
      Text6.Text = Str$(Xlchwidth%(mm%))
      Text7.Text = Str$(Xlhivo%(mm%))
      
      myselected% = mm%
End Sub

Private Sub Command1_Click()
   
       

    llevel% = Val(Text5.Text)
        If llevel% > 0 And llevel% < 256 Then GoTo over2
        MsgBox "LEVEL VALUE SHOULD BE FROM 1 TO 255"
        
       Exit Sub
over2:
     lchwidth% = Val(Text6.Text)
        If lchwidth% > 0 And lchwidth% < 256 Then GoTo over4:
        MsgBox "CHANNEL WIDTH VALUE SHOULD BE FROM 1 TO 255"
        
       Exit Sub
     
over4:
        lhivo% = Val(Text7.Text)
        If lhivo% < 1201 Then GoTo over5:
        If lhivo% > 1201 Then
          MsgBox " High voltage can not exceed 1200V"
        End If
       Exit Sub
over5:

  
      RIA_SPEC.Cursor(0).Value = llevel% / 2.55
      RIA_SPEC.Cursor(1).Value = 100 - Int((llevel% + lchwidth%) / 2.55)
  
   
       Form1.MScomm1.Output = "B"
       Form1.MScomm1.Output = Chr$(llevel%)
       Form1.MScomm1.Output = "A"
       lcorchwidth% = Int(lchwidth% * 1.33)
       If lcorchwidth% > 255 Then lcorchwidth% = 255
       Form1.MScomm1.Output = Chr$(lcorchwidth%)
       
      ' Form1.MScomm1.Output = "b"
       'Form1.MScomm1.Output = Chr$(llevel%)
      ' Form1.MScomm1.Output = "a"
      ' lcorchwidth% = Int(lchwidth% * 1.33)
       'If corchwidth% > 255 Then lcorchwidth% = 255
       'Form1.MScomm1.Output = Chr$(lcorchwidth%)
       
  'still HV commands:
       ff = 255 / 1200 'mapping 255 into 1200V
       
       Form1.MScomm1.Output = "H"
       Form1.MScomm1.Output = Chr$(Int(lhivo% * ff))
       
    RIA_SET.Hide
End Sub


Private Sub command3_Click()
  RIA_SET.Hide
  RIA_SPEC.Show
End Sub

Private Sub Form_Load()
 On Error GoTo sos22
 
  Open "c:\ria\chaninfo.rib" For Input As #1
  Do
      Input #1, Xisotope$(mm%)
      Input #1, Xlevel%(mm%)
      Input #1, Xchwidth%(mm%)
      Input #1, Xhivo%(mm%)
      Input #1, Xllevel%(mm%)
      Input #1, Xlchwidth%(mm%)
      Input #1, Xlhivo%(mm%)
      Combo1.AddItem Xisotope$(mm%)
      mm% = mm% + 1
      mmax% = mm%
  Loop While Not EOF(1)
  Close #1
     Exit Sub
sos22:
 ' missing "chaninfo.rib". So we can create one:
 Open "c:\ria\chaninfo.rib" For Output As #1
      Print #1, "Our First"
      Print #1, 100
      Print #1, 50
      Print #1, 760
      Print #1, 100
      Print #1, 50
      Print #1, 760
  Close #1
      Resume
End Sub

Private Sub Label2_Click()

End Sub

Private Sub iso_edit_Click()

End Sub

Private Sub remove_Click()
     Open "c:\ria\chaninfo.rib" For Output As #1
    Do
      If gg% = myselected% Then gg% = gg% + 1
      Print #1, Xisotope$(gg%)
      Print #1, Xlevel%(gg%)
      Print #1, Xchwidth%(gg%)
      Print #1, Xhivo%(gg%)
      Print #1, Xllevel%(gg%)
      Print #1, Xlchwidth%(gg%)
      Print #1, Xlhivo%(gg%)
      gg% = gg% + 1
    Loop While gg% < mmax%
     Close #1
 'number of items reduced for 1:
     mmax% = mmax% - 1
 'clean combo1 complete, then reload
    
    Combo1.RemoveItem myselected%
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
     Text6.Text = ""
     Text7.Text = ""
   Call combo_handle
End Sub

