VERSION 5.00
Begin VB.Form frmquest 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10845
      MouseIcon       =   "frmquest.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   3180
      Width           =   210
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   10845
      MouseIcon       =   "frmquest.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   6390
      Width           =   210
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   10845
      MouseIcon       =   "frmquest.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5355
      Width           =   210
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   10845
      MouseIcon       =   "frmquest.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4275
      Width           =   210
   End
   Begin VB.OptionButton Option2 
      Height          =   135
      Left            =   11685
      TabIndex        =   0
      Top             =   8460
      Width           =   15
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   360
      Left            =   750
      TabIndex        =   13
      Top             =   9045
      Width           =   1140
   End
   Begin VB.TextBox txtSh 
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C2C13&
      Height          =   5115
      Index           =   1
      Left            =   1350
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2925
      Width           =   3825
   End
   Begin VB.TextBox txtGz 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C2C13&
      Height          =   675
      Index           =   0
      Left            =   5370
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "frmquest.frx":0548
      Top             =   3480
      Width           =   5430
   End
   Begin VB.TextBox txtGz 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C2C13&
      Height          =   675
      Index           =   1
      Left            =   5370
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "frmquest.frx":054E
      Top             =   4515
      Width           =   5430
   End
   Begin VB.TextBox txtGz 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C2C13&
      Height          =   675
      Index           =   2
      Left            =   5370
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "frmquest.frx":0554
      Top             =   5595
      Width           =   5430
   End
   Begin VB.TextBox txtGz 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C2C13&
      Height          =   675
      Index           =   3
      Left            =   5370
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "frmquest.frx":055A
      Top             =   6675
      Width           =   5430
   End
   Begin VB.TextBox txtSh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F6E4D9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006C2C13&
      Height          =   705
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2175
      Width           =   9540
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   45
      Top             =   2100
   End
   Begin VB.Label lbl_Next 
      BackStyle       =   0  'Transparent
      Height          =   555
      Left            =   795
      MouseIcon       =   "frmquest.frx":0560
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   7770
      Width           =   420
   End
   Begin VB.Label lbl_Lesson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì”"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7245
      TabIndex        =   21
      Top             =   645
      Width           =   2220
   End
   Begin VB.Label lbl_Timer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2595
      TabIndex        =   20
      Top             =   660
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":‰«„ œ—”"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   0
      Left            =   9600
      TabIndex        =   19
      Top             =   645
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": “„«‰ »«ﬁÌ„«‰œÂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   2940
      TabIndex        =   18
      Top             =   645
      Width           =   1395
   End
   Begin VB.Label lbl_num 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11250
      TabIndex        =   12
      Top             =   2265
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10575
      TabIndex        =   5
      Top             =   6360
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( Ã"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10530
      TabIndex        =   4
      Top             =   5325
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( »"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   10500
      TabIndex        =   3
      Top             =   4230
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "( «·›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   10395
      TabIndex        =   2
      Top             =   3165
      Width           =   420
   End
   Begin VB.Label lbl_Exit 
      BackStyle       =   0  'Transparent
      Height          =   660
      Left            =   255
      MouseIcon       =   "frmquest.frx":06B2
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   540
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "frmquest.frx":0804
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmquest"
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
Dim TrQuest As Byte
Dim SlQuest As Byte
Dim iQuest As Integer
Dim i As Integer
Dim iTimer As Integer

Private Sub Command1_Click()
If Iq = 26 Then
 MsgBox " <<<< «‰ Â‹‹‹«Ì ¬“„‹‹‹‹‹‹Ê‰ >>>>  ", vbExclamation, "Â‘‹‹‹œ«—"
  Exit Sub
End If
num = Qnum(Iq)
SlQuest = SelectedOption
If TrQuest = SlQuest Then
  TrueCounter = TrueCounter + 1
  PutInArray (ls)
End If
 Call ShowQuest(num, ls)

IndexOp = SelectedOption
 Call UpdateQuest(IndexOp, num - 1, ls)
 Option2.SetFocus
Iq = Iq + 1
iTimer = 0
End Sub



Private Sub Form_Load()
Call Set_Align
iQuest = 1
lbl_num.Caption = ":" & 1
iTimer = 0
lbl_Lesson = gLesson
i = 0
Call RndGenerate(25)
num = Qnum(0)
ls = gLesson
Call ShowQuest(num, ls)
cnn.Close
Iq = Iq + 1
End Sub

Public Sub ShowQuest(ByVal Tnum As Integer, ByVal Tless As String)
'cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=db.mdb"
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db.mdb;Persist Security Info=False"
Tless = "'" & Tless & "'"
With cmd
.ActiveConnection = cnn
.CommandText = "select * from Tquest where num=" & Tnum & " and lesson =" & Tless
.CommandType = adCmdText
End With
Set rs = cmd.Execute
Do While rs.EOF = False
 txtSh(0).Text = Trim(rs.Fields(1).Value)
 If (rs.Fields(2).Value <> "") Then
  txtSh(1).Text = Trim(rs.Fields(2).Value)
Else: txtSh(1).Text = ""
 End If
 txtGz(0).Text = Trim(rs.Fields(3).Value)
 txtGz(1).Text = Trim(rs.Fields(4).Value)
 txtGz(2).Text = Trim(rs.Fields(5).Value)
 txtGz(3).Text = Trim(rs.Fields(6).Value)
 TrQuest = Trim(rs.Fields(8).Value)


rs.MoveNext
Loop
rs.Close

i = i + 1
End Sub

Public Sub UpdateQuest(ByVal Index As Integer, ByVal Tnum As Integer, ByVal Tless As String)
On Error GoTo er
Tless = "'" & Tless & "'"
With cmd
.ActiveConnection = cnn
.CommandText = "update Tquest set selected=" & Index & " where num=" & Tnum & " and lesson =" & Tless
.CommandType = adCmdText
End With
Set rs = cmd.Execute
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
End Function

Private Sub lbl_Next_Click()
Command1.Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
 Command1.Value = True
If iQuest < 25 Then iQuest = iQuest + 1
lbl_num.Caption = ":" & iQuest

End Sub

Private Sub lbl_Exit_Click()
Unload Me
End Sub


Public Sub PutInArray(strLess As String)
  Select Case strLess
    Case c_SA: Pkarname(0) = Pkarname(0) + 1
    Case c_BR: Pkarname(1) = Pkarname(1) + 1
    Case c_ZB: Pkarname(2) = Pkarname(2) + 1
    Case c_SD: Pkarname(3) = Pkarname(3) + 1
    Case c_ZT: Pkarname(4) = Pkarname(4) + 1
  End Select
End Sub


Private Sub Timer1_Timer()
  iTimer = iTimer + 1
  If iTimer = 40 Then
    Command1.Value = True
    iTimer = 0
  End If
lbl_Timer.Caption = 40 - iTimer
  End Sub



Public Sub Set_Align()
Dim iA As Byte
If gLesson = c_ZT Then
txtSh(0).RightToLeft = False
txtSh(0).Alignment = vbLeftJustify
For iA = 0 To 3
txtGz(iA).Alignment = vbLeftJustify
Next iA
Else:
txtSh(0).RightToLeft = True
txtSh(0).Alignment = vbRightJustify
For iA = 0 To 3
txtGz(iA).Alignment = vbRightJustify
Next iA

End If

End Sub

