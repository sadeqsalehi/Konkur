VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_about 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   10365
      MouseIcon       =   "frmmain.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3870
      Width           =   1140
   End
   Begin VB.Label lbl_exam 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   8895
      MouseIcon       =   "frmmain.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2340
      Width           =   1185
   End
   Begin VB.Label lbl_print 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   9090
      MouseIcon       =   "frmmain.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   1725
   End
   Begin VB.Label lbl_Pract 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9360
      MouseIcon       =   "frmmain.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lbl_Exit 
      BackStyle       =   0  'Transparent
      Height          =   465
      Left            =   270
      MouseIcon       =   "frmmain.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   645
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmmain.frx":069A
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_about_Click()
frmAbout.Show 1
End Sub

Private Sub lbl_exam_Click()
frmlesson.Show
End Sub

Private Sub lbl_Exit_Click()
End
End Sub

Private Sub lbl_Pract_Click()
frmlessonsp.Show
End Sub

Private Sub lbl_print_Click()
frmprintK.Show
End Sub
