VERSION 5.00
Object = "{1A028BB9-0262-48D8-9B67-D74CE4C2A7E8}#4.0#0"; "IngotTextBoxCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#4.0#0"; "IngotButtonCtl.dll"
Begin VB.Form frmHoraOficial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hora oficial"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2070
   BeginProperty Font 
      Name            =   "AFPalm"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   138
   StartUpPosition =   2  'CenterScreen
   Begin IngotButtonCtl.AFButton Command1 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "frmHoraOficial.frx":0000
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin IngotTextBoxCtl.AFTextBox text3 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "frmHoraOficial.frx":004C
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin IngotTextBoxCtl.AFTextBox text2 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmHoraOficial.frx":00A9
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin IngotTextBoxCtl.AFTextBox text1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmHoraOficial.frx":0106
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmHoraOficial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Private Sub Command1_Click()
    Dim lpSystemTime As SYSTEMTIME
    If text1.Text = "" And text2.Text = "" And text3.Text = "" Then
        Exit Sub
    End If
    lpSystemTime.wYear = 2003
    lpSystemTime.wMonth = 8
    lpSystemTime.wDayOfWeek = -1
    lpSystemTime.wDay = 28
    If text1.Text = 21 Then
        lpSystemTime.wHour = 0
    ElseIf text1.Text = 22 Then
        lpSystemTime.wHour = 1
    ElseIf text1.Text = 23 Then
        lpSystemTime.wHour = 2
    ElseIf text1.Text = 24 Then
        lpSystemTime.wHour = 3
    Else
        lpSystemTime.wHour = text1.Text + 3
    End If
    lpSystemTime.wMinute = text2.Text
    lpSystemTime.wSecond = text3.Text
    lpSystemTime.wMilliseconds = 0
    'set the new time
    SetSystemTime lpSystemTime
    Unload Me
End Sub

Private Sub Text1_Change()
    If Not IsNumeric(text1.Text) Then
        text1.Text = ""
    End If
End Sub

Private Sub Text2_Change()
    If Not IsNumeric(text2.Text) Then
        text2.Text = ""
    End If
End Sub

Private Sub Text3_Change()
    If Not IsNumeric(text3.Text) Then
        text3.Text = ""
    End If
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub


