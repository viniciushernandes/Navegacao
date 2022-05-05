VERSION 5.00
Begin VB.Form frmHoraOficial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hora oficial"
   ClientHeight    =   1245
   ClientLeft      =   3030
   ClientTop       =   3345
   ClientWidth     =   3855
   Icon            =   "form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ajustar Hora"
      Height          =   975
      Left            =   2640
      Picture         =   "form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      MaxLength       =   2
      TabIndex        =   0
      Top             =   360
      Width           =   735
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
    lpSystemTime.wYear = 2003
    lpSystemTime.wMonth = 8
    lpSystemTime.wDayOfWeek = -1
    lpSystemTime.wDay = 28
    If Text1.Text = 21 Then
        lpSystemTime.wHour = 0
    ElseIf Text1.Text = 22 Then
        lpSystemTime.wHour = 1
    ElseIf Text1.Text = 23 Then
        lpSystemTime.wHour = 2
    ElseIf Text1.Text = 24 Then
        lpSystemTime.wHour = 3
    Else
        lpSystemTime.wHour = Text1.Text + 3
    End If
    lpSystemTime.wMinute = Text2.Text
    lpSystemTime.wSecond = Text3.Text
    lpSystemTime.wMilliseconds = 0
    'set the new time
    SetSystemTime lpSystemTime
End Sub

Private Sub Text1_Change()
    If Not IsNumeric(Text1.Text) Then
        Text1.Text = ""
    End If
End Sub

Private Sub Text2_Change()
    If Not IsNumeric(Text2.Text) Then
        Text2.Text = ""
    End If
End Sub

Private Sub Text3_Change()
    If Not IsNumeric(Text3.Text) Then
        Text3.Text = ""
    End If
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

