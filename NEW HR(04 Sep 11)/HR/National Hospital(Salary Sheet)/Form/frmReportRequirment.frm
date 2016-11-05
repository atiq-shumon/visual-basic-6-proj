VERSION 5.00
Begin VB.Form frmReportRequirment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter your Requirment........."
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   3570
      Begin VB.TextBox txtfields 
         Height          =   330
         Left            =   945
         TabIndex        =   4
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emp ID"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   338
         Width           =   525
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   425
      Left            =   735
      Picture         =   "frmReportRequirment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Height          =   425
      Left            =   135
      Picture         =   "frmReportRequirment.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   540
   End
End
Attribute VB_Name = "frmReportRequirment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'rptmode = 7
'Form20.Show vbModal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
