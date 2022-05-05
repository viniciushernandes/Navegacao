VERSION 5.00
Begin VB.Form frmHoraLargada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hora da Largada"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      MaxLength       =   2
      TabIndex        =   0
      Top             =   390
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ajustar"
      Height          =   720
      Left            =   1920
      Picture         =   "frmHoraLargada.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   885
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      MaxLength       =   2
      TabIndex        =   1
      Top             =   390
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   2
      Top             =   390
      Width           =   495
   End
End
Attribute VB_Name = "frmHoraLargada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
        Exit Sub
    End If
    HoraLargada = Text1.Text & ":" & Text2.Text & ":" & Text3.Text
    Unload Me
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    If Trim(HoraLargada) <> "" Then
        Text1.Text = Mid(HoraLargada, 1, 2)
        Text2.Text = Mid(HoraLargada, 4, 2)
        Text3.Text = Mid(HoraLargada, 7, 2)
    End If
End Sub
