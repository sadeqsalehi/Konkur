VERSION 5.00
Begin VB.Form frmDB 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   13
      Top             =   5880
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8640
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   255
      Left            =   8640
      TabIndex        =   7
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   6000
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   5
      Top             =   4880
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   4
      Top             =   3760
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   7935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   9120
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim cnn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim num As Integer
Dim Sbool As Boolean
Dim IndexOp As Integer
Dim ls As String
Private Sub Command1_Click()

ls = "ÓíÓÊã ÚÇãá"
Call ShowQuest(num, ls)
cnn.Close
End Sub

Private Sub Command2_Click()
num = num + 1
 Call ShowQuest(num, ls)
IndexOp = SelectedOption
'MsgBox Str(IndexOp)
 Call UpdateQuest(IndexOp, num - 1)
End Sub

Private Sub Form_Load()
num = 1
End Sub

Public Sub ShowQuest(ByVal Tnum As Integer, ByVal Tless As String)
'On Error GoTo er
cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=db.mdb"
Tless = "'" & Tless & "'"
With cmd
.ActiveConnection = cnn
.CommandText = "select * from Tquest where num=" & Tnum & " and lesson =" & Tless
.CommandType = adCmdText
End With
Set rs = cmd.Execute
Do While rs.EOF = False
 Text1(0).Text = Trim(rs.Fields(0).Value)
 Text1(1).Text = Trim(rs.Fields(1).Value)
 Text1(2).Text = Trim(rs.Fields(2).Value)
rs.MoveNext
Loop
'Set rs = Nothing
rs.Close
'cnn.Close
'er:
'If Err.Number <> 0 Then
 '   MsgBox Str(Err.Number) + "          " + Err.Description
'End If
End Sub

Public Sub UpdateQuest(ByVal Index As Integer, ByVal Tnum As Integer)
'MsgBox Str(index)
On Error GoTo er
'cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=db.mdb"
With cmd
.ActiveConnection = cnn
.CommandText = "update Tquest set selected=" & Index & " where num=" & Tnum
.CommandType = adCmdText
End With
Set rs = cmd.Execute
'rs.Close
cnn.Close
er:
If Err.Number <> 0 Then
    MsgBox Str(Err.Number) + "          " + Err.Description
End If

End Sub

Public Function SelectedOption() As Integer
Dim i As Integer
If Option2.Value = True Then
   SelectedOption = 0
 Else:
    Do While (Option1(i).Value = False)
    i = i + 1
  Loop

SelectedOption = i + 1
End If
MsgBox Str(SelectedOption)
End Function

Private Sub Option1_Click(Index As Integer)
 Command2.Value = True
End Sub
