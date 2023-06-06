VERSION 5.00
Begin VB.Form Login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":3C3A
   ScaleHeight     =   9570
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2670
      TabIndex        =   5
      Top             =   6870
      Width           =   4305
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "LOGIN"
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
      Left            =   7530
      TabIndex        =   4
      Top             =   6870
      Width           =   4305
   End
   Begin VB.Frame PasswordFrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   4980
      Width           =   9195
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
         TabIndex        =   3
         Top             =   570
         Width           =   8505
      End
   End
   Begin VB.Frame UsernameFrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   3420
      Width           =   9195
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
         TabIndex        =   1
         Top             =   570
         Width           =   8505
      End
   End
   Begin VB.Label lbRegister 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dont have account?  register here"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   855
      Left            =   2670
      TabIndex        =   7
      Top             =   8190
      Width           =   9195
   End
   Begin VB.Label txtLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "McDelivery Sofware"
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
      Left            =   -360
      TabIndex        =   6
      Top             =   330
      Width           =   9525
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    
    Dim Username As String
    Username = txtUsername.Text
    
    Set rst = Nothing
    rst.Open "select * from loginTab where Username='" & Username & "'", conn
    
    If rst.EOF Then
       MsgBox "Username not found.", vbInformation, "Alert!"
    Else
       If rst!Status = "BANNED" Then
          MsgBox "This username has been banned.", vbInformation, "Alert!"
          Exit Sub
       End If
       
       Dim Role As String
       Role = rst!Role
       Username = rst!Username
       If rst!Pwd = txtPassword.Text Then
             Unload Me
             LoadingScreen.Role = Role
             LoadingScreen.Username = Username
             LoadingScreen.Show
             Else
             MsgBox "Incorrect Password!.", vbCritical, "Warning!"
       End If
          
    End If
       
End Sub

Private Sub Form_Load()
  On Error GoTo e
  conn.Provider = "Microsoft.jet.oledb.4.0"
  conn.ConnectionString = "Data Source=" & App.Path & "\db\Group2.mdb "

  conn.Open
  
  rst.Open "loginTab", conn
  
  Exit Sub
e:
  MsgBox Err.Description, vbCritical, "Warning!!"
End Sub

Private Sub lbRegister_Click()
    Unload Me
    Register.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  rst.Close
  conn.Close
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            cmdLogin_Click
            KeyAscii = 0
    End If
End Sub


Private Sub txtUsername_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
            txtPassword.SetFocus
            KeyAscii = 0
 End If
End Sub
