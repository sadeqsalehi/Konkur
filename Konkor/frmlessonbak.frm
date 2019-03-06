VERSION 5.00
Begin VB.Form frmlessons 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_physics 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9360
      MouseIcon       =   "frmlesson.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lbl_Exit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmlesson.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   -120
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmlesson.frx":02A4
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmlessons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lbl_Exit_Click()
Unload Me
End Sub

Private Sub lbl_physics_Click()
frmquest.Show
End Sub
