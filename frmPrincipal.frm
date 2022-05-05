VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Controle de Navegação"
   ClientHeight    =   3105
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4560
   ForeColor       =   &H80000007&
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCronômetro 
      Left            =   1920
      Top             =   2880
   End
   Begin VB.Timer TimerLargada 
      Interval        =   1000
      Left            =   1440
      Top             =   2880
   End
   Begin VB.Timer TimerNavegação 
      Left            =   960
      Top             =   2880
   End
   Begin VB.Timer TimerMsg 
      Left            =   480
      Top             =   2880
   End
   Begin VB.Timer TimerDeslocamento 
      Left            =   0
      Top             =   2880
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Deslocamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2010
      End
      Begin VB.Label lblTempoDeslocamento 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cronômetro"
      Height          =   195
      Left            =   1200
      TabIndex        =   18
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label lblCronômetro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Hora largada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Hora oficial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Média ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblMostrador 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1080
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblVM 
      AutoSize        =   -1  'True
      Caption         =   "42 Km/h"
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Média          :"
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblPF 
      AutoSize        =   -1  'True
      Caption         =   "1.970,000"
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Ponto final   :"
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblPI 
      AutoSize        =   -1  'True
      Caption         =   "0,000"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ponto inicial :"
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblLinha 
      AutoSize        =   -1  'True
      Caption         =   "Navegação"
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Próxima linha:"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblMsg 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
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
         Caption         =   "Executar programação"
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

Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Registro As String
    
    TimerDeslocamento.Interval = 0
    TimerMsg.Interval = 0
    
    ArqRelatório = "C:\Navegação.Rel"
    Open ArqRelatório For Input As #2
    While Not EOF(2)
        Line Input #2, Registro
        X = Mid(Registro, 1, 1)
        Linha(X) = "S"
        If Mid(Registro, 3, 12) = "Deslocamento" Or Mid(Registro, 3, 12) = "Neutralizado" Then
            Tipo(X) = Mid(Registro, 3, 12)
            t(X) = Mid(Registro, 16)
        Else
            Tipo(X) = Mid(Registro, 3, 9)
            PI(X) = Trim(Mid(Registro, 13, 9))
            PF(X) = Trim(Mid(Registro, 22, 9))
            VM(X) = Mid(Registro, 31)
        End If
    Wend
    Close #2
    Exit Sub
Erro:
    Exit Sub
End Sub

Private Sub mnuExecutar_Click()
    If MsgBox("Confirma execução agora?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    X = 0
    Total = 0
    HCr = 0
    MinCr = 0
    SegCr = 0
    TimerCronômetro.Interval = 1000
    Inicia
End Sub

Private Sub DN()
    lblTempoDeslocamento.Caption = t(X)
    H = Mid(t(X), 1, 2)
    Min = Mid(t(X), 4, 2)
    Seg = Mid(t(X), 7, 2)
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
    If SegCr = 59 Then
        SegCr = 0
        MinCr = MinCr + 1
        If MinCr = 60 Then
            HCr = HCr + 1
            MinCr = 0
        End If
    Else
        SegCr = SegCr + 1
    End If
    lblCronômetro.Caption = Format(HCr, "00") & ":" & Format(MinCr, "00") & ":" & Format(SegCr, "00")
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
        Cont = 0
        TimerMsg.Interval = 1000
        If Min = 5 Then
            sndPlaySound App.Path & "\CincoMinutos.WAV", SND_ASYNC
            lblMsg.Caption = "Cinco minutos para terminar o " & strDN & vbCrLf
        ElseIf Min = 1 Then
            sndPlaySound App.Path & "\UmMinuto.WAV", SND_ASYNC
            lblMsg.Caption = "Um minuto para terminar o " & strDN & vbCrLf
        ElseIf Seg <= 30 Then
            sndPlaySound App.Path & "\Segundos.WAV", SND_ASYNC
            lblMsg.Caption = "Atenção! " & Seg & " segundos para terminar o " & strDN & vbCrLf
        End If
        X = X + 1
        Label2.Visible = True
        Label4.Visible = True
        Label6.Visible = True
        Label8.Visible = True
        lblLinha.Caption = Tipo(X)
        lblPI.Caption = PI(X) & " Km"
        lblPF.Caption = PF(X) & " Km"
        lblVM.Caption = VM(X) & " Km/h"
        lblLinha.Visible = True
        lblPI.Visible = True
        lblPF.Visible = True
        lblVM.Visible = True
        X = X - 1
    End If
    If H = 0 And Min = 0 And Seg = 0 Then
        TimerDeslocamento.Interval = 0
        Frame1.Visible = False
        Label2.Visible = False
        Label4.Visible = False
        Label6.Visible = False
        Label8.Visible = False
        lblLinha.Visible = False
        lblPI.Visible = False
        lblPF.Visible = False
        lblVM.Visible = False
        lblMsg.Caption = ""
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
        X = 0
        Total = 0
        Inicia
    Else
        Label12.Caption = "Hora oficial: " & Time
        Label13.Caption = "Hora largada: " & HoraLargada
    End If
End Sub

Private Sub TimerMsg_Timer()
    Cont = Cont + 1
    If Cont = 10 Then
        lblMsg.Caption = ""
        TimerMsg.Interval = 0
    End If
End Sub

Private Sub Inicia()
    X = X + 1
    Label12.Visible = False
    Label13.Visible = False
    If Tipo(X) = "Deslocamento" Then
        Label1.Caption = "Deslocamento"
        Frame1.Visible = True
        Label7.Visible = False
        Label11.Visible = False
        lblMostrador.Visible = False
        strDN = "Deslocamento"
        DN
    ElseIf Tipo(X) = "Neutralizado" Then
        Label1.Caption = "Neutralizado"
        Frame1.Visible = True
        Label7.Visible = False
        Label11.Visible = False
        lblMostrador.Visible = False
        strDN = "Neutralizado"
        DN
    ElseIf Tipo(X) = "Navegação" Then
        Label7.Visible = True
        Label11.Visible = True
        Label11.Caption = VM(X) & " Km\h"
        Total = PI(X)
        VMs = CCur(VM(X)) * 0.36
        lblMostrador.Visible = True
        TimerNavegação.Interval = 1
    Else
        TimerCronômetro.Interval = 0
    End If
End Sub

Private Sub TimerNavegação_Timer()
    On Error Resume Next
    Total = Total + (VMs * 0.00001)
    lblMostrador.Caption = Format(Total, "00.00")
    If (Total) >= (PF(X)) Then
        TimerNavegação.Interval = 0
        Inicia
    End If
End Sub
