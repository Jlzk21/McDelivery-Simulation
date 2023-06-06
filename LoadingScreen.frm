VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LoadingScreen 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "LoadingScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   705
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   1244
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   420
      Top             =   3270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   1680
      TabIndex        =   1
      Top             =   330
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8370
      TabIndex        =   0
      Top             =   480
      Width           =   1245
   End
End
Attribute VB_Name = "LoadingScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Role As String
Public Username As String

Private Sub Form_Activate()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If Timer1.Enabled = True Then
        ProgressBar1.Value = ProgressBar1.Value + 1
        Label1.Caption = Val(Label1.Caption) + 1
        Label1.Caption = Format(Label1.Caption) & "%"
End If

If ProgressBar1.Value = 100 Then
    Unload Me
    If Role = "CUSTOMER" Then
        CustomerView.Username = Username
        CustomerView.Show
        ElseIf Role = "ADMIN" Then
            Dashboard.Username = Username
            Dashboard.Show
            ElseIf Role = "MANAGER" Then
                Management.Username = Username
                Management.Show
    End If
    
End If

End Sub

