VERSION 5.00
Object = "{BAE83016-CA11-4D8A-BBFA-0AE9863B82DE}#4.0#0"; "IngotLabelCtl.dll"
Object = "{1A298ECE-60AD-4C91-BB12-3092A2250D21}#4.0#0"; "IngotTimerCtl.dll"
Begin VB.Form frmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Navegação"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "AFPalm"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   StartUpPosition =   2  'CenterScreen
   Begin IngotLabelCtl.AFLabel Label14 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "frmPrincipal.frx":08CA
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin IngotLabelCtl.AFLabel Label8 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "frmPrincipal.frx":0918
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin IngotLabelCtl.AFLabel lblTempoDeslocamento 
      Height          =   375
      Left            =   1020
      OleObjectBlob   =   "frmPrincipal.frx":096A
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin IngotLabelCtl.AFLabel lblMostrador 
      Height          =   495
      Left            =   480
      OleObjectBlob   =   "frmPrincipal.frx":09B9
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin IngotLabelCtl.AFLabel lblCronômetro 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "frmPrincipal.frx":0A05
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin IngotTimerCtl.AFTimer TimerCronômetro 
      Height          =   480
      Left            =   1920
      OleObjectBlob   =   "frmPrincipal.frx":0A54
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotTimerCtl.AFTimer TimerLargada 
      Height          =   480
      Left            =   1440
      OleObjectBlob   =   "frmPrincipal.frx":0A75
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotTimerCtl.AFTimer TimerNavegação 
      Height          =   480
      Left            =   960
      OleObjectBlob   =   "frmPrincipal.frx":0A96
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotTimerCtl.AFTimer TimerDeslocamento 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmPrincipal.frx":0AB7
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin IngotLabelCtl.AFLabel Label13 
      Height          =   165
      Left            =   120
      OleObjectBlob   =   "frmPrincipal.frx":0AD8
      TabIndex        =   5
      Top             =   2190
      Width           =   2175
   End
   Begin IngotLabelCtl.AFLabel Label11 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmPrincipal.frx":0B29
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin IngotLabelCtl.AFLabel Label7 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmPrincipal.frx":0B77
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin IngotLabelCtl.AFLabel Label1 
      Height          =   135
      Left            =   120
      OleObjectBlob   =   "frmPrincipal.frx":0BC7
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin IngotLabelCtl.AFLabel Label12 
      Height          =   135
      Left            =   240
      OleObjectBlob   =   "frmPrincipal.frx":0C18
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Menu mnuOpções 
      Caption         =   "Opções"
      Begin VB.Menu mnuHoraOficial 
         Caption         =   "Hora oficial"
      End
      Begin VB.Menu mnuHoraLargada 
         Caption         =   "Hora largada"
      End
      Begin VB.Menu mnuProgramação 
         Caption         =   "Programação"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecutar 
         Caption         =   "Executar programção"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim H As Integer
Dim Min As Integer
Dim Seg As Integer
Dim lngTimerID As Long
Dim strDN As String
Dim VMs As Currency
Dim Total As Currency
Dim HCr As Integer
Dim MinCr As Integer
Dim SegCr As Integer
Dim DecCr As Integer

Private Sub Form_Load()
    Dim Registro As String
    On Error Resume Next
    
    TimerDeslocamento.Interval = 0
    
    ArqRelatório = "C:\Navegação.Rel"
    Open ArqRelatório For Input As #2
    While Not EOF(2)
        Line Input #2, Registro
        U = Mid(Registro, 1, 1)
        Linha(U) = "S"
        If Mid(Registro, 3, 12) = "Deslocamento" Or Mid(Registro, 3, 12) = "Neutralizado" Then
            Tipo(U) = Mid(Registro, 3, 12)
            t(U) = Mid(Registro, 16)
        Else
            Tipo(U) = Mid(Registro, 3, 9)
            PI(U) = Trim(Mid(Registro, 13, 9))
            PF(U) = Trim(Mid(Registro, 22, 9))
            VM(U) = Mid(Registro, 31)
        End If
    Wend
    Close #2
End Sub

Private Sub mnuExecutar_Click()
    If MsgBox("Confirma execução agora?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    U = 0
    Total = 0
    HCr = 0
    MinCr = 0
    SegCr = 0
    TimerCronômetro.Interval = 100
    Inicia
End Sub

Private Sub DN()
    lblTempoDeslocamento.Caption = t(U)
    H = Mid(t(U), 1, 2)
    Min = Mid(t(U), 4, 2)
    Seg = Mid(t(U), 7, 2)
    TimerDeslocamento.Interval = 1000
End Sub

Private Sub mnuHoraLargada_Click()
    frmHoraLargada.Show
End Sub

Private Sub mnuHoraOficial_Click()
    frmHoraOficial.Show
End Sub

Private Sub mnuProgramação_Click()
    frmOpções.Show
End Sub

Private Sub mnuSair_Click()
    If MsgBox("Deseja encerrar o programa?", vbInformation + vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub TimerCronômetro_Timer()
    If DecCr = 10 Then
        SegCr = SegCr + 1
        DecCr = 0
        If SegCr = 59 Then
            MinCr = MinCr + 1
        End If
        If MinCr = 600 Then
            HCr = HCr + 1
            MinCr = 0
        End If
    Else
        DecCr = DecCr + 1
    End If
    lblCronômetro.Caption = Format(HCr, "00") & ":" & Format(MinCr, "00") & ":" & Format(SegCr, "00") & ":" & Format(DecCr, "0")
End Sub

Private Sub TimerDeslocamento_Timer()
    If Seg > 0 Then
        Seg = Seg - 1
    ElseIf Min > 0 Then
        Seg = 59
        Min = Min - 1
    ElseIf H > 0 Then
        Seg = 59
        Min = 59
        H = H - 1
    End If
    lblTempoDeslocamento.Caption = Format(H, "00") & ":" & Format(Min, "00") & ":" & Format(Seg, "00")
    If H = 0 And Min = 5 And Seg = 0 Or H = 0 And Min = 1 And Seg = 0 Or H = 0 And Min = 0 And Seg <= 30 Then
        If Min = 5 Then
            sndPlaySound App.Path & "\CincoMinutos.WAV", SND_ASYNC
        ElseIf Min = 1 Then
            sndPlaySound App.Path & "\UmMinuto.WAV", SND_ASYNC
        ElseIf Seg <= 30 Then
            sndPlaySound App.Path & "\Segundos.WAV", SND_ASYNC
        End If
    End If
    If H = 0 And Min = 0 And Seg = 0 Then
        TimerDeslocamento.Interval = 0
        Label1.Visible = False
        lblTempoDeslocamento.Visible = False
        Inicia
    End If
End Sub

Private Sub TimerLargada_Timer()
    If Time = HoraLargada Then
        TimerLargada.Interval = 0
        HCr = 0
        MinCr = 0
        SegCr = 0
        TimerCronômetro.Interval = 1000
        U = 0
        Total = 0
        Inicia
    Else
        Label12.Caption = "Hora oficial: " & Time
        Label13.Caption = "Hora largada: " & HoraLargada
    End If
End Sub

Private Sub Inicia()
    U = U + 1
    Label12.Visible = False
    Label13.Visible = False
    If Tipo(U) = "Deslocamento" Then
        Label1.Caption = "Deslocamento"
        Label1.Visible = True
        lblTempoDeslocamento.Visible = True
        Label7.Visible = False
        Label11.Visible = False
        Label8.Visible = False
        Label14.Visible = False
        lblMostrador.Visible = False
        strDN = "Deslocamento"
        DN
    ElseIf Tipo(U) = "Neutralizado" Then
        Label1.Caption = "Neutralizado"
        Label1.Visible = True
        lblTempoDeslocamento.Visible = True
        Label7.Visible = False
        Label11.Visible = False
        Label8.Visible = False
        Label14.Visible = False
        lblMostrador.Visible = False
        strDN = "Neutralizado"
        DN
    ElseIf Tipo(U) = "Navegação" Then
        Label7.Visible = True
        Label11.Visible = True
        Label11.Caption = VM(U) & " Km\h"
        Label8.Visible = True
        Label14.Visible = True
        If VM(U + 1) = Empty Then
            Label14.Caption = Tipo(U + 1)
        Else
            Label14.Caption = VM(U + 1) & " Km\h"
        End If
        Total = PI(U)
        VMs = CCur(VM(U)) * 0.27777778
        lblMostrador.Visible = True
        TimerNavegação.Interval = 500
    Else
        TimerCronômetro.Interval = 0
    End If
End Sub

Private Sub TimerNavegação_Timer()
    On Error Resume Next
    Total = Total + (VMs / 2000)
    lblMostrador.Caption = Format(Total, "00.00")
    If (Total) >= (PF(U)) Then
        TimerNavegação.Interval = 0
        Inicia
    End If
End Sub

