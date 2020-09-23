VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Animated Software [ Test Form ]  - Jim Jose"
   ClientHeight    =   6420
   ClientLeft      =   180
   ClientTop       =   525
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      Height          =   6420
      Left            =   0
      ScaleHeight     =   6360
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdAnimate 
         Caption         =   "Load / Animate Form..."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txtCount 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Text            =   "33"
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Text            =   "11"
         Top             =   4560
         Width           =   735
      End
      Begin VB.ListBox lstEffect 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":002E
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count"
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   5160
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Frame Time"
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Animation Effect"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2040
      End
   End
   Begin VB.PictureBox picAsh 
      AutoSize        =   -1  'True
      Height          =   5115
      Left            =   3240
      Picture         =   "frmMain.frx":0164
      ScaleHeight     =   5055
      ScaleWidth      =   7335
      TabIndex        =   2
      Top             =   840
      Width           =   7395
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Click Here to View the Reverse Animation"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   2760
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Don't dare to make a second click on her. She is copyrighted to me ONLY!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   360
      Width           =   7920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnimate_Click()
    frmTest.Show
End Sub

Private Sub Form_Load()
    lstEffect.Selected(0) = True
End Sub

Private Sub Form_Resize()
    picAsh.Move (picLeft.Width + ScaleWidth - picAsh.Width) / 2, (ScaleHeight - picAsh.Height) / 2
End Sub

Private Sub lstEffect_Click()
    AnimateForm picAsh, aload, lstEffect.ListIndex, Val(txtTime), Val(txtCount)
End Sub

Private Sub picAsh_Click()
    AnimateForm picAsh, aUnload, lstEffect.ListIndex, Val(txtTime), Val(txtCount)
End Sub
