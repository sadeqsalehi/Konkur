VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmprintK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   75
      MouseIcon       =   "frmPrintk.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmPrintk.frx":0152
      ScaleHeight     =   840
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   540
      Width           =   840
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   12300
      ExtentX         =   21696
      ExtentY         =   15875
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmprintK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOpFile As String
Dim prt As Printer

Private Sub Form_Load()
Dim ind As Byte

 WebBrowser1.FullScreen = True
 StrOpFile = TextFromFile(App.Path + "\data\ctemplate.htm")
 Call Savefile(App.Path + "\data\ctemplate.htm")
 StrOpFile = Replace(StrOpFile, "$S1", Str(Pkarname(0)), 1)
 StrOpFile = Replace(StrOpFile, "$S2", Str(Pkarname(1)), 1)
 StrOpFile = Replace(StrOpFile, "$S3", Str(Pkarname(2)), 1)
 StrOpFile = Replace(StrOpFile, "$S4", Str(Pkarname(3)), 1)
 StrOpFile = Replace(StrOpFile, "$S5", Str(Pkarname(4)), 1)
 
 StrOpFile = Replace(StrOpFile, "$T1", CalcPr(Pkarname(0)), 1)
 StrOpFile = Replace(StrOpFile, "$T2", CalcPr(Pkarname(1)), 1)
 StrOpFile = Replace(StrOpFile, "$T3", CalcPr(Pkarname(2)), 1)
 StrOpFile = Replace(StrOpFile, "$T4", CalcPr(Pkarname(3)), 1)
 StrOpFile = Replace(StrOpFile, "$T5", CalcPr(Pkarname(4)), 1)
  
  For ind = 0 To 4
   Pkarname(5) = Pkarname(5) + Pkarname(ind)
  Next ind
 
 StrOpFile = Replace(StrOpFile, "$SK", Str(Pkarname(5)), 1)
 Pkarname(6) = (100 * Pkarname(5)) / 125
 StrOpFile = Replace(StrOpFile, "$TK", Str(Pkarname(6)) + "%", 1)

 StrOpFile = Replace(StrOpFile, "$G1", Str(25 - Pkarname(0)), 1)
 StrOpFile = Replace(StrOpFile, "$G2", Str(25 - Pkarname(1)), 1)
 StrOpFile = Replace(StrOpFile, "$G3", Str(25 - Pkarname(2)), 1)
 StrOpFile = Replace(StrOpFile, "$G4", Str(25 - Pkarname(3)), 1)
 StrOpFile = Replace(StrOpFile, "$G5", Str(25 - Pkarname(4)), 1)
 StrOpFile = Replace(StrOpFile, "$GK", Str(125 - Pkarname(5)), 1)

Call Savefile(App.Path + "\data\template.htm")
 
 WebBrowser1.Navigate App.Path + "\data\index.htm"
End Sub

Public Sub CreateHtml()

End Sub

Public Function TextFromFile(fInStream As String) As String
 
  Dim i As Long, strText As String
  i = FreeFile
  strText = ""
  Open fInStream For Input Lock Write As #i
  strText = StrConv(InputB$(LOF(i), i), vbUnicode)
  Close #i
  Screen.MousePointer = vbDefault
  TextFromFile = strText
End Function

Private Sub Savefile(ByVal Path As String)
Dim FileName As String
FileName = Path
F = FreeFile
Open FileName For Output As #F
Print #F, StrOpFile
Close #F
End Sub

Public Function CalcPr(ByVal Val1 As Integer) As String
  CalcPr = Str((100 * Val1) / 25) + "%"
End Function

Private Sub Picture1_Click()
Unload Me
End Sub
