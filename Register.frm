VERSION 5.00
Begin VB.Form Register 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
   Icon            =   "Register.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Register.frx":3C3A
   ScaleHeight     =   9570
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirm Password: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   7620
      TabIndex        =   13
      Top             =   5850
      Width           =   7000
      Begin VB.TextBox txtConfirmPassword 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   420
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   570
         Width           =   6000
      End
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7800
      TabIndex        =   12
      Top             =   7860
      Width           =   4305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   2940
      TabIndex        =   11
      Top             =   7860
      Width           =   4305
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact No: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   7680
      TabIndex        =   9
      Top             =   2400
      Width           =   7000
      Begin VB.TextBox txtContactNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   420
         TabIndex        =   10
         Top             =   570
         Width           =   6000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First Name: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   7000
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   420
         TabIndex        =   8
         Top             =   570
         Width           =   6000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   7590
      TabIndex        =   5
      Top             =   4110
      Width           =   7000
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   420
         TabIndex        =   6
         Top             =   570
         Width           =   6000
      End
   End
   Begin VB.Frame UsernameFrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   7000
      Begin VB.TextBox txtUsername 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   420
         TabIndex        =   4
         Top             =   570
         Width           =   6000
      End
   End
   Begin VB.Frame PasswordFrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   270
      TabIndex        =   1
      Top             =   5850
      Width           =   7000
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   420
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   570
         Width           =   6000
      End
   End
   Begin VB.Label txtLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "McDelivery Software"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1905
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   9525
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rstUsername As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
    Login.Show
End Sub
Private Function textValidator() As Boolean
    textValidator = Trim(txtFirstName) <> "" And Trim(txtLastName) <> "" And Trim(txtContactNo) <> "" And Trim(txtUsername) <> "" And Trim(txtPassword) <> ""
End Function
Private Function isUsernameExist() As Boolean
    rstUsername.Close
    rstUsername.Open "Select * from loginTab", conn, adOpenDynamic, adLockPessimistic
    With rstUsername
        If .RecordCount <> 0 Then
            Do Until .EOF
                 If Trim(txtUsername) = !Username Then
                    isUsernameExist = True
                 End If
                .MoveNext
            Loop
        End If
    End With
End Function

Private Sub cmdRegister_Click()
    If txtConfirmPassword <> txtPassword Then
        MsgBox "Password doesnt match.", vbExclamation, "Warning!"
        txtConfirmPassword.SetFocus
        Exit Sub
    End If
    
    If isUsernameExist Then
        MsgBox "Username already exist.", vbExclamation, "Warning!"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If Not textValidator Then
        MsgBox "Incomplete data. Please fillup all fields.", vbExclamation, "Warning!"
        Exit Sub
    End If
    conn.Execute "insert into loginTab(Firstname,Lastname,ContactNo,Username,Pwd,Role) values('" & Trim(txtFirstName) & "','" & Trim(txtLastName) & "','" & Trim(txtContactNo) & "','" & Trim(txtUsername) & "','" & Trim(txtPassword) & "','CUSTOMER')"
            
    MsgBox "Registered Successfully Added.", vbInformation, "Success"
    cmdCancel_Click
End Sub

Private Sub Form_Load()
  On Error GoTo e
  
  conn.Provider = "Microsoft.jet.oledb.4.0"
  conn.ConnectionString = "Data Source=" & App.Path & "\db\Group2.mdb "
  conn.Open
  
  rst.Open "loginTab", conn
  rstUsername.Open "Select * from loginTab", conn, adOpenDynamic, adLockPessimistic
  
  Exit Sub
e:
  MsgBox Err.Description, vbCritical, "Warning!!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  rst.Close
  rstUsername.Close
  conn.Close
  
End Sub

