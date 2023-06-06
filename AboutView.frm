VERSION 5.00
Begin VB.Form AboutView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Us"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AboutView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2925
      Left            =   360
      Picture         =   "AboutView.frx":3C3A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Member:   Karyll Jane Carillo, Mia Bianca D. Regaas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   240
      TabIndex        =   3
      Top             =   6720
      Width           =   4275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Assistant Programmer:  Jomari Divina Gracia"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   240
      TabIndex        =   2
      Top             =   6090
      Width           =   4275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Head Programmer : JuliusBiascan (jlzkdev)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   240
      TabIndex        =   1
      Top             =   5730
      Width           =   4275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"AboutView.frx":FAD4
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   300
      TabIndex        =   0
      Top             =   4110
      Width           =   4275
   End
End
Attribute VB_Name = "AboutView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

