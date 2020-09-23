VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Test"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FDDBAC&
      Caption         =   "&Unload"
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1755
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AnimateForm Me, aload, frmMain.lstEffect.ListIndex, Val(frmMain.txtTime), Val(frmMain.txtCount)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AnimateForm Me, aUnload, frmMain.lstEffect.ListIndex, Val(frmMain.txtTime), Val(frmMain.txtCount)
End Sub
