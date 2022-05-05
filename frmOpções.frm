VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpções 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programação"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Ver programação"
      Height          =   615
      Left            =   2760
      TabIndex        =   21
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar programação"
      Height          =   615
      Left            =   2760
      TabIndex        =   20
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPróxima 
      Caption         =   "Nova linha"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Excluir programação"
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Confirmar"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtVM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtPF 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtPI 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtLinha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   375
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmOpções.frx":0000
         Left            =   960
         List            =   "frmOpções.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskTempo 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Km/h"
         Height          =   195
         Left            =   2040
         TabIndex        =   19
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Km"
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Km"
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tempo"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Média"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ponto final"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ponto inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Linha"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmOpções"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Anterior As Integer

Private Sub cboTipo_Click()
    If cboTipo.ListIndex = 0 Then
        Label3.Enabled = True
        Label5.Enabled = True
        Label6.Enabled = True
        txtPI.Enabled = True
        txtPF.Enabled = True
        txtVM.Enabled = True
        
        Label4.Enabled = False
        mskTempo.Enabled = False
    Else
        Label3.Enabled = False
        Label5.Enabled = False
        Label6.Enabled = False
        txtPI.Enabled = False
        txtPF.Enabled = False
        txtVM.Enabled = False
        
        Label4.Enabled = True
        mskTempo.Enabled = True
    End If
End Sub

Private Sub cmdConfirma_Click()
    On Error GoTo Erro
    
    If Trim(txtLinha.Text) = "" Then
        MsgBox "Informe o nº da linha!", vbInformation
        txtLinha.SetFocus
        Exit Sub
    End If
    
    If cboTipo.ListIndex = 0 Then
        If Trim(txtPI.Text) = "" Then
            MsgBox "Informe o ponto inicial!", vbInformation
            txtPI.SetFocus
            Exit Sub
        End If
        If Trim(txtPF.Text) = "" Then
            MsgBox "Informe o ponto final!", vbInformation
            txtPF.SetFocus
            Exit Sub
        End If
        If Trim(txtVM.Text) = "" Then
            MsgBox "Informe velocidade média!", vbInformation
            txtVM.SetFocus
            Exit Sub
        End If
        Tipo(txtLinha.Text) = cboTipo.Text
        PI(txtLinha.Text) = txtPI.Text
        PF(txtLinha.Text) = txtPF.Text
        VM(txtLinha.Text) = txtVM.Text
        t(txtLinha.Text) = ""
        Linha(txtLinha.Text) = "S"
    
    Else
        If mskTempo.Text = "__:__:__" Then
            MsgBox "Informe o tempo!", vbInformation
            mskTempo.SetFocus
            Exit Sub
        End If
        Tipo(txtLinha.Text) = cboTipo.Text
        PI(txtLinha.Text) = ""
        PF(txtLinha.Text) = ""
        VM(txtLinha.Text) = ""
        t(txtLinha.Text) = mskTempo.Text
        Linha(txtLinha.Text) = "S"
    End If
    InicializaCampos
    txtLinha.SetFocus
    Exit Sub
Erro:
    MsgBox "Erro!", vbInformation
End Sub

Private Sub InicializaCampos()
    txtLinha.Text = ""
    cboTipo.ListIndex = 0
    txtPI.Text = ""
    txtPF.Text = ""
    txtVM.Text = ""
    mskTempo.Mask = ""
    mskTempo.Text = ""
    mskTempo.Mask = "##:##:##"
End Sub

Private Sub cmdLimpar_Click()
    If MsgBox("Confirma?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    InicializaCampos
    For X = 1 To 99
        Linha(X) = ""
        Tipo(X) = ""
        PI(X) = ""
        PF(X) = ""
        VM(X) = ""
        t(X) = ""
    Next X
    Screen.MousePointer = vbArrow
End Sub

Private Sub cmdPróxima_Click()
    InicializaCampos
    For X = 1 To 99
        If Linha(X) = "" Then
            txtLinha.Text = X
            Exit For
        End If
    Next X
    txtLinha.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
    Screen.MousePointer = vbHourglass
    
    ArqRelatório = "C:\Navegação.Rel"
    Open ArqRelatório For Output As #1
     
    For X = 1 To 10000
        If Linha(X) = "S" Then
            If Tipo(X) = "Deslocamento" Or Tipo(X) = "Neutralizado" Then
                Print #1, Tab(1); Trim(X); " "; Tipo(X); Tab(16); t(X)
            Else
                Print #1, Tab(1); Trim(X); " "; Tipo(X); Tab(13); PI(X); Tab(22); PF(X); Tab(31); VM(X)
            End If
        Else
            Exit For
        End If
    Next X
    Close #1
    Screen.MousePointer = vbArrow
End Sub


Private Sub Command2_Click()
    Shell "NotePad C:\Navegação.Rel", vbNormalFocus
End Sub

Private Sub Form_Load()
    InicializaCampos
End Sub

Private Sub mskTempo_GotFocus()
    mskTempo.SelStart = 0
    mskTempo.SelLength = Len(mskTempo.Text)
End Sub

Private Sub txtLinha_LostFocus()
    On Error Resume Next
    If Trim(txtLinha.Text) = "" Then
        InicializaCampos
        Exit Sub
    End If
    If Linha(txtLinha.Text) = "" Then
        Anterior = txtLinha.Text - 1
        If Linha(Anterior) = "S" Then
            Do Until Tipo(Anterior) = "Navegação"
                Anterior = Anterior - 1
                If Anterior <= 0 Then
                    Exit Do
                End If
            Loop
            If Tipo(Anterior) = "Navegação" Then
                txtPI.Text = Format(PF(Anterior), "##,##0.000")
            End If
        End If
        Exit Sub
    End If
    cboTipo.Text = Tipo(txtLinha.Text)
    txtPI.Text = PI(txtLinha.Text)
    txtPF.Text = PF(txtLinha.Text)
    txtVM.Text = VM(txtLinha.Text)
    If t(txtLinha.Text) = "" Then
        mskTempo.Text = "__:__:__"
    Else
        mskTempo.Text = t(txtLinha.Text)
    End If
    cboTipo.SetFocus
End Sub

Private Sub txtPF_GotFocus()
    txtPF.SelStart = 0
    txtPF.SelLength = Len(txtPF.Text)
End Sub

Private Sub txtPF_LostFocus()
    On Error Resume Next
    txtPF.Text = Format(txtPF.Text, "##,##0.000")
End Sub

Private Sub txtPI_GotFocus()
    txtPI.SelStart = 0
    txtPI.SelLength = Len(txtPI.Text)
End Sub

Private Sub txtPI_LostFocus()
    On Error Resume Next
    txtPI.Text = Format(txtPI.Text, "##,##0.000")
End Sub

Private Sub txtVM_GotFocus()
    txtVM.SelStart = 0
    txtVM.SelLength = Len(txtVM.Text)
End Sub

Private Sub txtVM_LostFocus()
    On Error Resume Next
    txtVM.Text = Format(txtVM.Text, "##,##0.0")
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

