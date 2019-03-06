VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ”  ò‰òÊ—"
   ClientHeight    =   4155
   ClientLeft      =   4860
   ClientTop       =   3615
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2970
      Width           =   2385
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2364
      Width           =   2385
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1761
      Width           =   2385
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1158
      Width           =   2385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   555
      Width           =   2385
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
gLesson = c_SA
frmTquest.Show
End Sub

Private Sub Command2_Click()
gLesson = c_BR
frmTquest2.Show
End Sub

Private Sub Command3_Click()
gLesson = c_ZB
frmTquest3.Show
End Sub

Private Sub Command4_Click()
gLesson = c_SD
frmTquest4.Show
End Sub

Private Sub Command5_Click()
gLesson = c_ZT
frmTquest5.Show
End Sub

Private Sub Form_Load()
Command1.Caption = c_SA
Command2.Caption = c_BR
Command3.Caption = c_ZB
Command4.Caption = c_SD
Command5.Caption = c_ZT
End Sub
