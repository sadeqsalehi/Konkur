VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFields 
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim cmd As New ADODB.Command
Private Sub Form_Load()

 Dim X As Variant
cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=db.mdb"
With cmd
.ActiveConnection = cnn
.CommandText = "select count(*) from Tquest "
.CommandType = adCmdText
End With
Set rst = cmd.Execute
Do While rst.EOF = False
txtFields.Text = rst.Fields(0).Value
rst.MoveNext
Loop
 'Text1.Text = Str(ss)
  Ss = Val(txtFields.Text) + 1
  txtFields.Text = Str(Ss)
If rst.EOF And rst.BOF Then txtFields.Text = "1"
 Ss = Val(txtFields.Text)
rst.Close
cnn.Close

  ' Call setrecord
  mbDataChanged = False
 

End Sub
