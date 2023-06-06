VERSION 5.00
Begin VB.Form CustomerAddress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Address"
   ClientHeight    =   13485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerAddress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13485
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000080FF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   270
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveAddress 
      BackColor       =   &H000080FF&
      Caption         =   "SAVE ADDRESS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   12570
      Width           =   7545
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7515
      Left            =   90
      TabIndex        =   7
      Top             =   4890
      Width           =   9045
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Addtional Address Information"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2325
         Left            =   90
         TabIndex        =   16
         Top             =   5070
         Width           =   8805
         Begin VB.TextBox txtAddInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   300
            TabIndex        =   17
            Top             =   540
            Width           =   8205
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Street: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   90
         TabIndex        =   14
         Top             =   3870
         Width           =   8805
         Begin VB.TextBox txtStreet 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   330
            TabIndex        =   15
            Top             =   540
            Width           =   8205
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "District: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   120
         TabIndex        =   12
         Top             =   2730
         Width           =   8805
         Begin VB.TextBox txtDistrict 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   330
            TabIndex        =   13
            Top             =   450
            Width           =   8205
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Region: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   90
         TabIndex        =   10
         Top             =   480
         Width           =   8805
         Begin VB.TextBox txtRegion 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   330
            TabIndex        =   11
            Top             =   480
            Width           =   8205
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "City: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1065
         Left            =   90
         TabIndex        =   8
         Top             =   1500
         Width           =   8805
         Begin VB.TextBox txtCity 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   330
            TabIndex        =   9
            Top             =   480
            Width           =   8205
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   90
      TabIndex        =   2
      Top             =   2190
      Width           =   9045
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone Number: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   945
         Left            =   90
         TabIndex        =   5
         Top             =   1590
         Width           =   8805
         Begin VB.TextBox txtPhoneNO 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   330
            TabIndex        =   6
            Top             =   450
            Width           =   8205
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fullname: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   90
         TabIndex        =   3
         Top             =   480
         Width           =   8805
         Begin VB.TextBox txtFullname 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   330
            TabIndex        =   4
            Top             =   510
            Width           =   8205
         End
      End
   End
   Begin VB.Frame UsernameFrm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address Alias: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   1140
      Width           =   9045
      Begin VB.TextBox txtAddressAlias 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   330
         TabIndex        =   1
         Text            =   "Address 1"
         Top             =   420
         Width           =   8505
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Address"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   330
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -60
      Top             =   0
      Width           =   9945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checkout"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1770
      TabIndex        =   19
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "CustomerAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Public Username As String

Private Function textValidator() As Boolean
    textValidator = Trim(txtAddressAlias) <> "" And Trim(txtFullname) <> "" And Trim(txtPhoneNO) <> "" And Trim(txtRegion) <> "" _
    And Trim(txtDistrict) <> "" And Trim(txtCity) <> "" And Trim(txtAddInfo) <> "" And Trim(txtStreet) <> "" And Trim(txtAddInfo) <> ""
End Function

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdSaveAddress_Click()
    If Not textValidator Then
        MsgBox "Incomplete data. Please fillup all fields.", vbExclamation, "Warning!"
        Exit Sub
    End If
    If rst.EOF Then
       rst.AddNew
    End If
    
    rst.Fields("Fullname").Value = Trim(txtFullname)
    rst.Fields("PhoneNumber").Value = Trim(txtPhoneNO)
    rst.Fields("Region").Value = Trim(txtRegion)
    rst.Fields("City").Value = Trim(txtCity)
    rst.Fields("District").Value = Trim(txtDistrict)
    rst.Fields("Street").Value = Trim(txtStreet)
    rst.Fields("AdditionalInfo").Value = Trim(txtAddInfo)
    rst.Fields("AddressAlias").Value = Trim(txtAddressAlias)
    rst.Fields("Username").Value = Username
    
    MsgBox "Address saved successfully ..!!!", vbInformation
    rst.Update
End Sub


Private Sub Form_Load()
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\Group2.mdb;Persist Security Info=False"
    rst.Open "Select * from addressTab WHERE Username='" & Username & "'", conn, adOpenDynamic, adLockPessimistic
    If Not rst.EOF Then
       txtFullname.Text = rst!Fullname
       txtPhoneNO.Text = rst!PhoneNumber
       txtRegion.Text = rst!Region
       txtCity.Text = rst!City
       txtDistrict.Text = rst!District
       txtStreet.Text = rst!Street
       txtAddInfo.Text = rst!AdditionalInfo
       txtAddressAlias.Text = rst!AddressAlias
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  rst.Close
  conn.Close
End Sub

