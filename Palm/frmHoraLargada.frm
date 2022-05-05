VERSION 5.00
Object = "{1A028BB9-0262-48D8-9B67-D74CE4C2A7E8}#4.0#0"; "IngotTextBoxCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#4.0#0"; "IngotButtonCtl.dll"
Begin VB.Form frmHoraLargada 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Largada"
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
      OleObjectBlob   =   "frmHoraLargada.frx":0000
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin IngotTextBoxCtl.AFTextBox text3 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "frmHoraLargada.frx":004C
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin IngotTextBoxCtl.AFTextBox text2 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmHoraLargada.frx":00A9
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin IngotTextBoxCtl.AFTextBox text1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmHoraLargada.frx":0106
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmHoraLargada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If text1.Text = "" And text2.Text = "" And text3.Text = "" Then
        Exit Sub
    End If
    HoraLargada = text1.Text & ":" & text2.Text & ":" & text3.Text
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
        text1.Text = Mid(HoraLargada, 1, 2)
        text2.Text = Mid(HoraLargada, 4, 2)
        text3.Text = Mid(HoraLargada, 7, 2)
    End If
End Sub

