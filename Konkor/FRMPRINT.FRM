VERSION 5.00
Begin VB.Form frmprint 
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
   Begin VB.TextBox Text2 
      Height          =   720
      Left            =   945
      TabIndex        =   3
      Text            =   "END"
      Top             =   7425
      Width           =   2880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   480
      Left            =   4065
      TabIndex        =   2
      Top             =   7920
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "print"
      Height          =   240
      Left            =   1455
      TabIndex        =   1
      Top             =   8700
      Width           =   1425
   End
   Begin VB.Label lbl_Exit 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   15
      MouseIcon       =   "frmprint.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      Width           =   1695
   End
End
Attribute VB_Name = "frmprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 DataReport1.Show
End Sub

Private Sub Command2_Click()
MSChart1.EditCopy
Printer.Print Text1.Text
Printer.PaintPicture Clipboard.GetData(), 200, 200
Printer.EndDoc

End Sub

Private Sub Form_Load()

End Sub

Private Sub lbl_Exit_Click()
Unload Me
End Sub


