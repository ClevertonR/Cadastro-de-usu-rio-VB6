VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0080C0FF&
   Caption         =   "Exemplo de Login"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "&Encerrar"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdManageUsers 
      Caption         =   "Gerenciar Usuários"
      Height          =   420
      Left            =   2880
      TabIndex        =   0
      Top             =   990
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exemplo de Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   4080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Made By Michael Ciurescu (CVMichael from vbforums.com)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdEncerrar_Click()
  If (MsgBox("Deseja encerrar a aplicação ? ", vbYesNo)) = vbYes Then
    End
  End If
End Sub

Private Sub cmdLogin_Click()
    frmLogin.Show vbModal, Me
End Sub

Private Sub cmdManageUsers_Click()
    frmManageUsers.Show vbModal, Me
End Sub

Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    ' carrega o banco de dados
    Set DBConn = LoadDatabase(App.Path & "\LogInExemplo.mdb")
    
    ' conta os registros para ver se existe usuários na tabela
    Set rsData = DBConn.Execute("SELECT Count(*) FROM tblUsers")
    
    ' se existem usuários na tabela então carrega o login
    If rsData.Fields(0).value > 0 Then
        If Not DoLogin Then
            Unload Me ' Login falhou, então descarrega o formulário
        End If
        frmMain.Caption = frmMain.Caption & " - " & NomeUsuario
    End If
End Sub

Private Function DoLogin() As Boolean
    Dim UserName As String, Password As String, Ret As Boolean
    Dim LoginComSucesso As Boolean, rsData As ADODB.Recordset
    Dim MD5 As New clsMD5
    
    Randomize
    
    ' Pega o usuário com último login no registro
    UserName = GetSetting(App.EXEName, "Settings", "LastLogIn", "")
    
    ' solicita ao usuário o login e a senha
    Ret = frmLogin.GetLogIn(UserName, Password, Me)
    
    Do While Ret
        Set rsData = DBConn.Execute("SELECT ID, UserName, Password FROM tblUsers WHERE UserName = '" & Replace(UserName, "'", "''") & "'")
        
        ' se o registro foi encontrado , então o usuário existe
        If rsData.RecordCount > 0 Then
            ' verifica se a senha esta correta
            If UCase(MD5.DigestStrToHexStr(Password)) = UCase(rsData("Password").value) Then
                
                ' a senha esta correta, logo salva o usuário que se logou
                LogInUserID = rsData("ID").value
                LogInUserName = rsData("UserName").value
                
                ' salva o nome do usuário no Registry
                SaveSetting App.EXEName, "Settings", "LastLogIn", rsData("UserName").value
                
                LoginComSucesso = True
                Exit Do
            End If
        End If
        
        If Not LoginComSucesso Then
            Ret = False
            
            If MsgBox("Login inválido, quer tentar novamente ?", vbQuestion + vbYesNo, "Login inválido") = vbYes Then
                ' para evitar o ataque de força bruta a partir da aplicação
                Sleep 200 + 300 * Rnd
                
                ' se o login falhou , solicita novamente até que o usuário cancele
                Ret = frmLogin.GetLogIn(UserName, Password, Me)
            End If
        End If
    Loop
    
    DoLogin = LoginComSucesso
End Function
