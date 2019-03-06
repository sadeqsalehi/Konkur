VERSION 5.00
Begin VB.Form frmlesson 
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
   Begin VB.Label lbl_ZT 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   8895
      MouseIcon       =   "frmlesson.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5730
      Width           =   1830
   End
   Begin VB.Label lbl_SD 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9690
      MouseIcon       =   "frmlesson.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4545
      Width           =   1710
   End
   Begin VB.Label lbl_ZB 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9930
      MouseIcon       =   "frmlesson.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3300
      Width           =   1860
   End
   Begin VB.Label lbl_BR 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9315
      MouseIcon       =   "frmlesson.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2115
      Width           =   2430
   End
   Begin VB.Label lbl_SA 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9480
      MouseIcon       =   "frmlesson.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lbl_Exit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmlesson.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmlesson.frx":07EC
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmlesson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call ResetArrayPK
End Sub

Private Sub lbl_BR_Click()
Iq = 0
gLesson = c_BR
frmquest.Show
End Sub

Private Sub lbl_Exit_Click()
Unload Me
End Sub

Private Sub lbl_SA_Click()
Iq = 0
gLesson = c_SA
frmquest.Show
End Sub

Public Sub ResetArrayPK()
 Dim i As Integer
  For i = 0 To 13
    Pkarname(i) = 0
  Next i
End Sub

Private Sub lbl_SD_Click()
Iq = 0
gLesson = c_SD
frmquest.Show

End Sub

Private Sub lbl_ZB_Click()
Iq = 0
gLesson = c_ZB
frmquest.Show

End Sub

Private Sub lbl_ZT_Click()
Iq = 0
gLesson = c_ZT
frmquest.Show

End Sub
