Attribute VB_Name = "m1"
Option Explicit
Dim cnn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Global Qnum(100) As Integer
Global Pkarname(17) As Integer
Global Const c_SA = "”Ì” „ ⁄«„·"
Global Const c_BR = "»—‰«„Â ”«“Ì ò«„ÅÌÊ —"
Global Const c_ZB = "–ŒÌ—Â Ê »«“Ì«»Ì"
Global Const c_SD = "”«Œ „«‰ œ«œÂ"
Global Const c_ZT = "“»«‰  Œ’’Ì"
Global gLesson As String
Global Iq As Integer
Dim Sbool As Boolean
Global TrueCounter As Integer

Public Sub RndGenerate(ByVal CeilNum As Integer)
 Dim i As Integer
 Dim RndNum As Integer
 i = 0
 Do While (i <> CeilNum)
  Randomize
  RndNum = Round(Rnd(1) * 40) + 1
  Sbool = Rptcheck(RndNum)
  If Sbool <> True Then
   Debug.Print Str(i) + "  " + Str(RndNum)
   Qnum(i) = RndNum
   i = i + 1
  End If
 Loop
 
End Sub

Public Function Rptcheck(ByVal Rn As Integer) As Boolean
 Dim k As Integer
 k = 0
Do While Qnum(k) <> 0
   If Qnum(k) = Rn Then
    Rptcheck = True
    GoTo End_func
   End If
  k = k + 1
 Loop
End_func:
End Function





