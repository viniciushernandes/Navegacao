VERSION 5.00
Object = "{1A028BB9-0262-48D8-9B67-D74CE4C2A7E8}#4.0#0"; "IngotTextBoxCtl.dll"
Object = "{A54BEB34-AAB3-4A8D-B736-42CB4DA7D664}#4.0#0"; "IngotComboBoxCtl.dll"
Object = "{BAE83016-CA11-4D8A-BBFA-0AE9863B82DE}#4.0#0"; "IngotLabelCtl.dll"
Object = "{9F4EED48-8EC7-4316-A47D-F6161874E478}#4.0#0"; "IngotButtonCtl.dll"
Begin VB.Form frmOpções 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programação"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2400
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
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   StartUpPosition =   2  'CenterScreen
   Begin IngotButtonCtl.AFButton cmdSair 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "frmOpções.frx":0000
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin IngotButtonCtl.AFButton Command2 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "frmOpções.frx":004B
      TabIndex        =   16
      Top             =   1200
      Width           =   615
   End
   Begin IngotButtonCtl.AFButton cmdLimpar 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "frmOpções.frx":0093
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin IngotButtonCtl.AFButton Command1 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "frmOpções.frx":00DF
      TabIndex        =   14
      Top             =   480
      Width           =   615
   End
   Begin IngotButtonCtl.AFButton cmdPróxima 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "frmOpções.frx":012A
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin IngotButtonCtl.AFButton cmdConfirma 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmOpções.frx":0174
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin IngotTextBoxCtl.AFTextBox mskTempo 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "frmOpções.frx":01C2
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel Label4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmOpções.frx":021F
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin IngotTextBoxCtl.AFTextBox txtVM 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "frmOpções.frx":0269
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel Label6 
      Height          =   135
      Left            =   120
      OleObjectBlob   =   "frmOpções.frx":02C6
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin IngotTextBoxCtl.AFTextBox txtPF 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "frmOpções.frx":0310
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel Label5 
      Height          =   135
      Left            =   120
      OleObjectBlob   =   "frmOpções.frx":036D
      TabIndex        =   10
      Top             =   1080
      Width           =   495
   End
   Begin IngotTextBoxCtl.AFTextBox txtPI 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "frmOpções.frx":03BA
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin IngotLabelCtl.AFLabel Label3 
      Height          =   135
      Left            =   120
      OleObjectBlob   =   "frmOpções.frx":0417
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin IngotLabelCtl.AFLabel AFLabel2 
      Height          =   165
      Left            =   960
      OleObjectBlob   =   "frmOpções.frx":0466
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin IngotComboBoxCtl.AFComboBox cboTipo 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmOpções.frx":04AF
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin IngotTextBoxCtl.AFTextBox txtLinha 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmOpções.frx":0503
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin IngotLabelCtl.AFLabel AFLabel1 
      Height          =   135
      Left            =   90
      OleObjectBlob   =   "frmOpções.frx":0560
      TabIndex        =   2
      Top             =   120
      Width           =   375
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
    Dim PFm, PIm, VMm As Currency
    Dim Tempo As AFTimer
    
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
            MsgBox "Informe a velocidade média!", vbInformation
            txtVM.SetFocus
            Exit Sub
        End If
        Tipo(txtLinha.Text) = cboTipo.Text
        PI(txtLinha.Text) = txtPI.Text
        PF(txtLinha.Text) = txtPF.Text
        VM(txtLinha.Text) = txtVM.Text
        VMm = CCur(txtVM.Text) * 0.27777778
        Set Tempo = (CCur(txtPF.Text) - CCur(txtPI.Text)) / VMm
        tn(txtLinha.Text) = Tempo
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
    mskTempo.Text = ""
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
    MsgBox "Sua programação foi salva!", vbInformation
    Screen.MousePointer = vbArrow
End Sub


Private Sub Command2_Click()
    Shell "NotePad C:\Navegação.Rel", vbNormalFocus
End Sub

Private Sub Form_Load()
    InicializaCampos
    cboTipo.AddItem "Navegação"
    cboTipo.AddItem "Deslocamento"
    cboTipo.AddItem "Neutralizado"
    cboTipo.ListIndex = 0
End Sub

Private Sub mskTempo_GotFocus()
    mskTempo.SelStart = 0
    mskTempo.SelLength = Len(mskTempo.Text)
End Sub

Private Sub txtLinha_GotFocus()
    txtLinha.SelStart = 0
    txtLinha.SelLength = Len(txtLinha.Text)
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
    If Tipo(txtLinha.Text) = "Navegação" Then
        cboTipo.ListIndex = 0
    ElseIf Tipo(txtLinha.Text) = "Deslocamento" Then
        cboTipo.ListIndex = 1
    ElseIf Tipo(txtLinha.Text) = "Neutralizado" Then
        cboTipo.ListIndex = 2
    End If
    cboTipo_Click
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


