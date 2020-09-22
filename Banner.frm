VERSION 5.00
Begin VB.Form Banner 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Banner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Banner.frx":000C
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBanner 
      Interval        =   100
      Left            =   6960
      Top             =   120
   End
End
Attribute VB_Name = "Banner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngTickCount As Long

Private Sub Form_Load()
    lngTickCount = 0
    Banner.Show

    Dim curFont As Long
    Dim strVersion As String
    
    strVersion = "Version 1.1"
    curFont = CreateFont(12, 0, 0, 0, 400, True, False, False, 1, 7, 0, 2, 0, "MS Sans Serif")
    Call SelectObject(Me.hdc, curFont)
    Call TextOut(Me.hdc, 252, 38, strVersion, Len(strVersion))
    Call DeleteObject(curFont)
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1

    Load Main
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Banner.Hide
    'MsgBox "This program is in beta-stage of development." & vbNewLine & _
           "There is a known bug which causes problems" & vbNewLine & _
           "with the redraw when the program has been" & vbNewLine & _
           "running for a long time. However this will" & vbNewLine & _
           "differ per system." & vbNewLine & vbNewLine & _
           "Make sure you save all your files prior to" & vbNewLine & _
           "running this application.", vbOKOnly Or vbInformation, "Ant Colony Silumation Stability Warning"
End Sub

Private Sub tmrBanner_Timer()
    lngTickCount = lngTickCount + 1
    If lngTickCount = 30 Then Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub
