VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ManageUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Users"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15690
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ManageUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   15690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBan 
      Caption         =   "BAN"
      Height          =   975
      Left            =   14730
      TabIndex        =   15
      Top             =   1530
      Width           =   885
   End
   Begin VB.CommandButton cmdRemoved 
      Caption         =   "REMOVE"
      Height          =   975
      Left            =   14730
      TabIndex        =   14
      Top             =   420
      Width           =   885
   End
   Begin MSComctlLib.ListView lstUsers 
      Height          =   5865
      Left            =   6330
      TabIndex        =   0
      Top             =   390
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   10345
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add User"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5955
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6015
      Begin VB.Frame Frame6 
         Caption         =   "Role: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   13
         Top             =   4320
         Width           =   5805
         Begin VB.ComboBox txtRole 
            Height          =   345
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   5505
         End
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "REGISTER"
         Height          =   705
         Left            =   90
         TabIndex        =   12
         Top             =   5070
         Width           =   5745
      End
      Begin VB.Frame Frame5 
         Caption         =   "Password: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   10
         Top             =   3570
         Width           =   5805
         Begin VB.TextBox txtPassword 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   11
            Top             =   270
            Width           =   5355
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contact No: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   8
         Top             =   2790
         Width           =   5805
         Begin VB.TextBox txtContactNo 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   9
            Top             =   270
            Width           =   5355
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Last Name: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   6
         Top             =   2070
         Width           =   5805
         Begin VB.TextBox txtLName 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   7
            Top             =   270
            Width           =   5355
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "First Name: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   4
         Top             =   1290
         Width           =   5805
         Begin VB.TextBox txtFname 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   5
            Top             =   270
            Width           =   5355
         End
      End
      Begin VB.Frame UsernameFrm 
         Caption         =   "Username: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   90
         TabIndex        =   2
         Top             =   480
         Width           =   5805
         Begin VB.TextBox txtUsername 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   210
            TabIndex        =   3
            Top             =   270
            Width           =   5355
         End
      End
   End
End
Attribute VB_Name = "ManageUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset

Private Function textValidator() As Boolean
    textValidator = Trim(txtFname) <> "" And Trim(txtLName) <> "" And Trim(txtContactNo) <> "" And Trim(txtUsername) <> "" And Trim(txtPassword) <> ""
End Function

Private Function isUsernameExist() As Boolean
    rst.Close
    rst.Open "Select * from loginTab", conn, adOpenDynamic, adLockPessimistic
    With rst
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

Private Sub cmdBan_Click()
    Dim i As Integer
    Dim confirm As String
    With lstUsers
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                confirm = MsgBox("Do you want to ban Username: " & .ListItems(i).Text, vbYesNo + vbCritical, "Deletion Confirmation")
                If confirm = vbYes Then
                    rst.Close
                    rst.Open "Select * from loginTab where Username='" & .ListItems(i).Text & "'", conn, adOpenStatic, adLockPessimistic
                    rst.Fields("Status") = "BANNED"
                    
                    MsgBox "Record has been banned successfully", vbInformation, "Message"
                    rst.Update
                        Else
                    MsgBox "User Not Banned ..!!", vbInformation, "Message"
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdRegister_Click()
    If isUsernameExist Then
        MsgBox "Username already exist.", vbExclamation, "Warning!"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If Not textValidator Then
        MsgBox "Incomplete data. Please fillup all fields.", vbExclamation, "Warning!"
        Exit Sub
    End If
    conn.Execute "insert into loginTab(Firstname,Lastname,ContactNo,Username,Pwd,Role) values('" & Trim(txtFname) & "','" & Trim(txtLName) & "','" & Trim(txtContactNo) & "','" & Trim(txtUsername) & "','" & Trim(txtPassword) & "','" & Trim(txtRole) & "')"
            
    MsgBox "Registered Successfully Added.", vbInformation, "Success"
    displayUsers
    
    txtFname = ""
    txtLName = ""
    txtContactNo = ""
    txtUsername = ""
    txtPassword = ""
    txtRole.Text = txtRole.List(0)
End Sub

Private Sub cmdRemoved_Click()
    Dim i As Integer
    Dim confirm As String
    With lstUsers
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                confirm = MsgBox("Do you want to delete Username: " & .ListItems(i).Text, vbYesNo + vbCritical, "Deletion Confirmation")
                If confirm = vbYes Then
                    rst.Close
                    rst.Open "Select * from loginTab where Username='" & .ListItems(i).Text & "'", conn, adOpenStatic, adLockPessimistic
                    
                    rst.Delete adAffectCurrent
                    MsgBox "Record has been Deleted successfully", vbInformation, "Message"
                    rst.Update
                        Else
                    MsgBox "User Not Deleted ..!!", vbInformation, "Message"
                End If
            End If
        Next
    End With
    refreshdata
End Sub

Sub refreshdata()
        rst.Close
        rst.Open "Select * from loginTab", conn, adOpenStatic, adLockPessimistic
        If Not rst.EOF Then
            rst.MoveNext
            displayUsers
        Else
            MsgBox "No Record Found"
        End If
End Sub


Private Sub Form_Load()
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\Group2.mdb;Persist Security Info=False"
    rst.Open "Select * from loginTab", conn, adOpenDynamic, adLockPessimistic
    
    txtRole.AddItem "CUSTOMER", 0
    txtRole.AddItem "MANAGER", 1
    txtRole.AddItem "ADMIN", 1
    txtRole.Text = txtRole.List(0)
    
    With lstUsers.ColumnHeaders
        .Add , , "Username", lstUsers.Width / 5
        .Add , , "First Name", lstUsers.Width / 5
        .Add , , "Last Name", lstUsers.Width / 5
        .Add , , "Contact No", lstUsers.Width / 5
        .Add , , "Role", lstUsers.Width / 5
    End With
    displayUsers
    
End Sub

Public Sub displayUsers()
    rst.Close
    rst.Open "Select * from loginTab", conn, adOpenDynamic, adLockPessimistic
    
    With rst
        lstUsers.ListItems.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
                lstUsers.ListItems.Add , , !Username
                lstUsers.ListItems(lstUsers.ListItems.Count).SubItems(1) = !Firstname
                lstUsers.ListItems(lstUsers.ListItems.Count).SubItems(2) = !Lastname
                lstUsers.ListItems(lstUsers.ListItems.Count).SubItems(3) = !ContactNo
                lstUsers.ListItems(lstUsers.ListItems.Count).SubItems(4) = !Role
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    rst.Close
    conn.Close
End Sub
