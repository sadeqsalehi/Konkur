VERSION 5.00
Begin VB.Form frmsplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4485
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4320
   End
   Begin VB.Image Image1 
      Height          =   4260
      Left            =   0
      Picture         =   "Form2.frx":14862
      Top             =   0
      Width           =   5430
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Dim i As Integer
Private Sub Form_Load()
    Dim lngRegion As Long
    Dim lngReturn As Long
    Dim lngFormWidth As Long
    Dim lngFormHeight As Long
    
    lngFormWidth = Image1.Width / Screen.TwipsPerPixelX
    lngFormHeight = Image1.Height / Screen.TwipsPerPixelY
    lngRegion = CreateEllipticRgn(12, 12, lngFormWidth, lngFormHeight)
    lngReturn = SetWindowRgn(Me.hWnd, lngRegion, True)
End Sub

Private Sub Image1_Click()
 Unload Me
frmmain.Show
End Sub

Private Sub Timer1_Timer()
  i = i + 1
  If i = 4 Then Call GoInit
  End Sub
Private Sub GoInit()
   Unload Me
   frmmain.Show
End Sub



