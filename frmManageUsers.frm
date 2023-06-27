VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageUsers 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerencia Usuários"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOldPassword 
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1305
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2595
      Width           =   2895
   End
   Begin VB.TextBox txtPassword2 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1305
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3555
      Width           =   2895
   End
   Begin VB.TextBox txtPassword1 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1305
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3075
      Width           =   2895
   End
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2115
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Sair"
      Height          =   375
      Left            =   4410
      TabIndex        =   6
      Top             =   3555
      Width           =   960
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   375
      Left            =   4410
      TabIndex        =   5
      Top             =   2070
      Width           =   960
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   1860
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   3281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome Usuário"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label lblOldPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Senha Anterior:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Confirma Senha:"
      Height          =   420
      Left            =   180
      TabIndex        =   9
      Top             =   3510
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Senha:"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   3150
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Usuário :"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   2190
      Width           =   630
   End
End
Attribute VB_Name = "frmManageUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Made By Michael Ciurescu (CVMichael from vbforums.com)

' Lista todos os usuários no ListView
Private Sub ListUsers()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DBConn.Execute("SELECT ID, UserName FROM tblUsers")
    
    lstUsers.ListItems.Clear
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        
        Do Until rsData.EOF Or rsData.BOF
            With lstUsers.ListItems.Add(, , rsData("ID").value & "")
                .SubItems(1) = rsData("UserName").value & ""
            End With
            rsData.MoveNext
        Loop
    End If
End Sub

Private Sub cmdAdd_Click()

On Error GoTo trata_erro

    Dim Request As String, NewID As Long, rsData As Recordset
    Dim MD5 As New clsMD5, NewPassword As String, OldPassword As String
    
    If Len(txtPassword1.Text) > 0 Or Len(txtPassword2.Text) > 0 Then
        If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "A senha de confirmação deve ser a mesma que a senha inforamda no campo Senha", vbExclamation
            Exit Sub
        End If
    End If
    
    ' pega o hash das senhas
    NewPassword = UCase(MD5.DigestStrToHexStr(Me.txtPassword1.Text))
    OldPassword = UCase(MD5.DigestStrToHexStr(Me.txtOldPassword.Text))
    
    If cmdAdd.Caption = "&Adicionar" Then
    
        If Len(txtUserName.Text) < 6 Then
            MsgBox "O nome do usuário deve ser que 5 caracteres", vbExclamation
            Exit Sub
        End If
        If Len(txtPassword1.Text) < 6 Then
            MsgBox "A senha deve ser maior que 5 caracteres", vbExclamation
            Exit Sub
        End If
    
        ' pega o ID para o novo registro
        NewID = SelectNewID(DBConn, "tblUsers")
        
        ' prepara a instrução INSERT
        Request = "INSERT INTO tblUsers VALUES(" & NewID & "," & _
            "'" & Replace(Me.txtUserName.Text, "'", "''") & "'," & _
            "'" & NewPassword & "')"
    Else
        ' se esta logado como um usuário diferente daquele que estamos tratando
        If LogInUserID <> Val(lstUsers.SelectedItem.Text) Then
            ' valida a senha se estiver logado com usuário distinto
            Set rsData = DBConn.Execute("SELECT Password FROM tblUsers WHERE ID = " & Me.lstUsers.SelectedItem.Text)
            If OldPassword <> rsData("Password").value Then
                MsgBox "Senha anterior inválida." & vbNewLine & "Você tem que entrar a senha válida para o usuário selecionado.", vbInformation
                Exit Sub
            End If
        End If
        
        ' Prepara a instrução Update
        Request = "UPDATE tblUsers SET UserName = '" & Replace(Me.txtUserName.Text, "'", "''") & "'"
        If Len(Me.txtPassword1.Text) > 0 Then
            Request = Request & ", [Password] = '" & NewPassword & "'"
        End If
        Request = Request & " WHERE ID = " & Me.lstUsers.SelectedItem.Text
        
        ' altera a caption de volta para "Adicionar"
        cmdAdd.Caption = "&Adicionar"
        txtUserName.Enabled = True
    End If
    
    ' executa o pedido
    DBConn.Execute Request
    
    ' Reseta os controles
    ListUsers
    txtUserName.Text = ""
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    txtOldPassword.Enabled = False
    lblOldPassword.Enabled = False
    lstUsers.Enabled = True
    MsgBox "Operação realizada com sucesso !!!", vbInformation, "OK"
    Exit Sub
    
trata_erro:
    MsgBox ("Ocoerreu um erro : " & Err.Description)
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmManageUsers.Caption = frmManageUsers.Caption & " - Usuário logado: " & NomeUsuario
    ListUsers
End Sub

Private Sub lstUsers_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lstUsers_DblClick()
    If Not (lstUsers.SelectedItem Is Nothing) Then
        Me.txtUserName.Text = lstUsers.SelectedItem.SubItems(1)
        
        txtUserName.Enabled = False
        lstUsers.Enabled = False
        cmdAdd.Caption = "&Atualiza"
        
        ' habilita os campos txtOldPassword quando l ogado como um usuário diferente
        If LogInUserID <> Val(lstUsers.SelectedItem.Text) Then
            txtOldPassword.Enabled = True
            lblOldPassword.Enabled = True
        End If
    End If
End Sub
