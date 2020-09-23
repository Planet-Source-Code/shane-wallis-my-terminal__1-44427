VERSION 5.00
Begin VB.Form frmTerminal 
   Caption         =   "Terminal"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   Icon            =   "frmTerminal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDisplay 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The Transparent code comes from
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=31050&lngWId=1

'Terminal by S.Wallis 2003
Dim StartPos As Integer

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private WithEvents objDOS As DOSOutputs
Attribute objDOS.VB_VarHelpID = -1
Private Sub Form_Load()
Dim NormalWindowStyle As Long

NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
Set objDOS = New DOSOutputs
SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED

SetLayeredWindowAttributes Me.hwnd, 0, 155, LWA_ALPHA

Me.BackColor = RGB(0, 0, 140)
txtDisplay.BackColor = RGB(0, 0, 140)
    txtDisplay = "Shane's Bodacious Terminal [Version 1.0]" & vbNewLine _
    & "(C) Copyright 2003 HBRC - Hooterville Board Riders Club"
    txtDisplay = txtDisplay & vbNewLine & vbNewLine
    txtDisplay.SelStart = Len(txtDisplay)
    StartPos = Len(txtDisplay) + 1
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)
    txtDisplay = txtDisplay & CommandOutputs
    txtDisplay.SelStart = Len(txtDisplay)
End Sub
Private Sub txtDisplay_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        txtDisplay = txtDisplay & vbNewLine
        Dim CmdLine As Variant
        Dim EndPos As Integer, i As Integer
        EndPos = Len(txtDisplay)
        CmdLine = Mid(txtDisplay, StartPos, EndPos - StartPos + 1)
        objDOS.CommandLine = "cmd.exe /c " & CmdLine
        objDOS.ExecuteCommand
        StartPos = Len(txtDisplay) + 3
    End If
    
Exit Sub
ErrHandler:
    MsgBox Err.Number & vbNewLine & Err.Description & vbNewLine & "Email elvis007now@hotmail.com", vbOKOnly Or vbCritical, "Terminal"
    Resume Next
End Sub
