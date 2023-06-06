VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CustomerView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custumer View"
   ClientHeight    =   13560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17760
   Icon            =   "CustomerView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13560
   ScaleWidth      =   17760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H000080FF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4110
      Width           =   3795
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   12660
      Width           =   3795
   End
   Begin VB.CommandButton cmdAddress 
      BackColor       =   &H000080FF&
      Caption         =   "My Address"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3150
      Width           =   3795
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H000080FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2220
      Width           =   3795
   End
   Begin TabDlg.SSTab sstView 
      Height          =   9615
      Left            =   4170
      TabIndex        =   2
      Top             =   3780
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "CustomerView.frx":3C3A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "menuPhoto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtProductDescription"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDescTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtProductName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCategory"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPriceDesc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdNext"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrev"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOrder"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmQuantity"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "My Orders"
      TabPicture(1)   =   "CustomerView.frx":3C56
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstOrder"
      Tab(1).Control(1)=   "frmOrder"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "My Purchases"
      TabPicture(2)   =   "CustomerView.frx":3C72
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdRefresh"
      Tab(2).Control(1)=   "lstPurchases"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Me"
      TabPicture(3)   =   "CustomerView.frx":3C8E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "cmdSaveProfile"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "Frame4"
      Tab(3).Control(4)=   "Frame5"
      Tab(3).Control(5)=   "Label1"
      Tab(3).ControlCount=   6
      Begin VB.Frame Frame6 
         Caption         =   "Contact:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -67500
         TabIndex        =   47
         Top             =   3000
         Width           =   4575
         Begin VB.TextBox txtContact 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   48
            Top             =   480
            Width           =   3975
         End
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H000080FF&
         Caption         =   "Refresh"
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
         Left            =   -74760
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   8880
         Width           =   2775
      End
      Begin VB.Frame frmQuantity 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   6330
         TabIndex        =   42
         Top             =   6630
         Width           =   2715
         Begin VB.CommandButton cmdIncrement 
            BackColor       =   &H0000FF00&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2010
            TabIndex        =   46
            Top             =   420
            Width           =   525
         End
         Begin VB.CommandButton cmdDecrement 
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   150
            TabIndex        =   45
            Top             =   420
            Width           =   525
         End
         Begin VB.TextBox txtQuantity 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1020
            TabIndex        =   43
            Text            =   "1"
            Top             =   450
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   180
         TabIndex        =   35
         Top             =   660
         Width           =   12945
         Begin VB.CommandButton cmdMcCafe 
            Caption         =   "MC CAFE"
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
            Left            =   10080
            TabIndex        =   41
            Top             =   390
            Width           =   1245
         End
         Begin VB.CommandButton cmdFries 
            Caption         =   "FRIES"
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
            Left            =   11490
            TabIndex        =   40
            Top             =   390
            Width           =   1155
         End
         Begin VB.CommandButton cmdChickenAndPlaters 
            Caption         =   "CHICKEN AND PLATERS"
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
            Left            =   4320
            TabIndex        =   39
            Top             =   390
            Width           =   2835
         End
         Begin VB.CommandButton cmdDrinksAndDesserts 
            Caption         =   "DRINK AND DESERTS"
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
            Left            =   7320
            TabIndex        =   38
            Top             =   390
            Width           =   2565
         End
         Begin VB.CommandButton cmdBurgers 
            Caption         =   "BURGERS"
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
            Left            =   2340
            TabIndex        =   37
            Top             =   390
            Width           =   1875
         End
         Begin VB.CommandButton cmdBreakfast 
            Caption         =   "BREAKFAST"
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
            Left            =   300
            TabIndex        =   36
            Top             =   390
            Width           =   1875
         End
      End
      Begin VB.CommandButton cmdOrder 
         BackColor       =   &H0000FFFF&
         Caption         =   "ADD TO ORDER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   9270
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   6750
         Width           =   3045
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H000080FF&
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6960
         Width           =   2265
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H000080FF&
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3210
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6960
         Width           =   2235
      End
      Begin VB.Frame Frame2 
         Height          =   3315
         Left            =   -64920
         TabIndex        =   17
         Top             =   5670
         Width           =   2895
         Begin VB.CommandButton btnCancel 
            BackColor       =   &H0080FFFF&
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   150
            Picture         =   "CustomerView.frx":3CAA
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1740
            Width           =   2535
         End
         Begin VB.CommandButton btnPayment 
            BackColor       =   &H0080FFFF&
            Caption         =   "CHECK OUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   180
            Picture         =   "CustomerView.frx":486A
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   300
            Width           =   2535
         End
      End
      Begin VB.Frame frmOrder 
         Caption         =   "Order Detail"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -64920
         TabIndex        =   16
         Top             =   690
         Width           =   2955
         Begin VB.ComboBox cbPaymentMethod 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   3630
            Width           =   2505
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Php""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13321
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   210
            MaxLength       =   14
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   2190
            Width           =   2535
         End
         Begin VB.TextBox txtDeliveryFee 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   210
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   900
            Width           =   2535
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   3210
            Width           =   1920
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   210
            TabIndex        =   23
            Top             =   1830
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery Fee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   210
            TabIndex        =   21
            Top             =   540
            Width           =   1410
         End
      End
      Begin MSComctlLib.ListView lstOrder 
         Height          =   8175
         Left            =   -74790
         TabIndex        =   15
         Top             =   840
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   14420
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdSaveProfile 
         Caption         =   "SAVE CHANGES"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -73770
         TabIndex        =   14
         Top             =   5250
         Width           =   4665
      End
      Begin VB.Frame Frame1 
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -73770
         TabIndex        =   12
         Top             =   3000
         Width           =   4575
         Begin VB.TextBox txtUsername 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   3975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -73770
         TabIndex        =   9
         Top             =   1710
         Width           =   4575
         Begin VB.TextBox txtFName 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   3975
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1095
         Left            =   -67500
         TabIndex        =   7
         Top             =   1710
         Width           =   4575
         Begin VB.TextBox txtLName 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   3975
         End
      End
      Begin MSComctlLib.ListView lstPurchases 
         Height          =   8085
         Left            =   -74820
         TabIndex        =   26
         Top             =   690
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   14261
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label txtPriceDesc 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10710
         TabIndex        =   34
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label txtCategory 
         AutoSize        =   -1  'True
         Caption         =   "Breakfast"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5730
         TabIndex        =   33
         Top             =   1770
         Width           =   1230
      End
      Begin VB.Label txtProductName 
         AutoSize        =   -1  'True
         Caption         =   "Cheesy Eggdesal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5730
         TabIndex        =   32
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label txtDescTitle 
         AutoSize        =   -1  'True
         Caption         =   "Product Description"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5790
         TabIndex        =   31
         Top             =   3180
         Width           =   2550
      End
      Begin VB.Label txtProductDescription 
         Caption         =   "Melty cheese wrapped in a fluffy, folded egg, sandwiched between a soft, toasted pandesal bun."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2550
         Left            =   5820
         TabIndex        =   30
         Top             =   3720
         Width           =   6885
      End
      Begin VB.Image menuPhoto 
         Height          =   4680
         Left            =   180
         Picture         =   "CustomerView.frx":5755
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Setting"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69900
         TabIndex        =   11
         Top             =   780
         Width           =   2625
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   30
      Left            =   4020
      TabIndex        =   1
      Top             =   300
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image banner 
      Height          =   3285
      Left            =   4290
      Picture         =   "CustomerView.frx":1687B
      Stretch         =   -1  'True
      Top             =   300
      Width           =   13065
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   240
      Picture         =   "CustomerView.frx":3EF83
      Stretch         =   -1  'True
      Top             =   690
      Width           =   840
   End
   Begin VB.Label txtFullname 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   990
      Width           =   2025
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3795
      Y1              =   1920
      Y2              =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   13545
      Left            =   30
      Top             =   30
      Width           =   3795
   End
End
Attribute VB_Name = "CustomerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Username As String
Dim Category As String
Private TotAmount As Currency
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rsPurchaces As New ADODB.Recordset
Dim rsProfile As New ADODB.Recordset



Private Sub btnPayment_Click()
    Set rs = Nothing
    rs.Open "Select * from addressTab WHERE Username='" & Username & "'", con, adOpenDynamic, adLockPessimistic
    
    If Not lstOrder.ListItems.Count = 0 Then
        If rs.EOF Then
           MsgBox "Please add address first before you proceed to payment.", vbInformation, "Alert!"
           CustomerAddress.Username = Username
           CustomerAddress.Show
        Else
            PaymentView.Username = Username
            PaymentView.Show
        End If
    Else
        MsgBox "Please make an order first before you proceed to payment", vbInformation, "Alert"
    End If
    
    cmdBreakfast_Click
End Sub

Private Sub cmdAbout_Click()
    AboutView.Show
End Sub

Private Sub cmdAddress_Click()
    CustomerAddress.Username = Username
    CustomerAddress.Show
End Sub

Private Sub cmdBreakfast_Click()
    Category = "BREAKFAST"
    Set rs = Nothing
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    displayMenu
End Sub

Private Sub cmdBurgers_Click()
    Category = "BURGERS"
    Set rs = Nothing
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    displayMenu
End Sub

Private Sub cmdChickenAndPlaters_Click()
    Category = "CHICKENANDPLATERS"
    Set rs = Nothing
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    displayMenu
End Sub

Private Sub cmdDrinksAndDesserts_Click()
    Category = "DRINKSANDDESSERTS"
    Set rs = Nothing
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    displayMenu
End Sub

Private Sub cmdExit_Click()
    Dim answer As String
    answer = MsgBox("Do you want to quit?", vbExclamation + vbYesNo, "Confirm")
    If answer = vbYes Then
        End
    Else
        MsgBox "Action canceled", vbInformation, "Confirm"
    End If
End Sub

Private Sub cmdFries_Click()
    Category = "FRIES"
    Set rs = Nothing
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    displayMenu
End Sub

Private Sub cmdDecrement_Click()
    If Not Val(txtQuantity.Text) <= 1 Then
        txtQuantity.Text = Str(Val(txtQuantity.Text) - 1)
    End If
End Sub

Private Sub cmdIncrement_Click()
    txtQuantity.Text = Str(Val(txtQuantity.Text) + 1)
End Sub

Private Sub cmdLogout_Click()
    Dim answer As String
    answer = MsgBox("Are you sure you want to signout?", vbExclamation + vbYesNo, "Confirm")
    If answer = vbYes Then
        Unload Me
        Login.Show
    Else
        MsgBox "Action canceled", vbInformation, "Confirm"
    End If
    
End Sub

Private Sub cmdMcCafe_Click()
    Category = "MCCAFE"
    Set rs = Nothing
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    displayMenu
End Sub

Private Sub cmdNext_Click()
    txtQuantity.Text = 1
    rs.MoveNext
    If Not rs.EOF Then
        displayMenu
    Else
        rs.MoveFirst
        displayMenu
    End If
End Sub

Private Sub cmdOrder_Click()
    If Not IsNumeric(txtQuantity.Text) Or Val(txtQuantity.Text) < 1 Then
        MsgBox "Please enter valid quantity", vbInformation, "Alert"
        txtQuantity.Text = 1
        Exit Sub
    End If
    
    lstOrder.ListItems.Add , , rs!itemCode
    lstOrder.ListItems(lstOrder.ListItems.Count).SubItems(1) = rs!ProductName
    lstOrder.ListItems(lstOrder.ListItems.Count).SubItems(2) = txtQuantity.Text
    lstOrder.ListItems(lstOrder.ListItems.Count).SubItems(3) = Val(rs!Price) * Val(txtQuantity.Text)
    If Val(Val(txtQuantity.Text)) > 10 Then
        txtDeliveryFee.Text = 150
        Else
        txtDeliveryFee.Text = 70
    End If
    
    TotAmount = Format((Val(txtAmount) + (Val(rs!Price) * Val(txtQuantity.Text))), "###,##0.00")
    txtAmount = TotAmount
End Sub

Private Sub btnCancel_Click()
    If lstOrder.ListItems.Count > 0 Then
        TotAmount = Format((Val(txtAmount) - Val(lstOrder.SelectedItem.ListSubItems(3))), "###,##0.00")
        If TotAmount = 0 Then
            txtDeliveryFee.Text = 0
        End If
        
        txtAmount = TotAmount
        lstOrder.ListItems.Remove lstOrder.SelectedItem.Index
    End If
    
End Sub

Private Sub cmdPrev_Click()
    txtQuantity.Text = 1
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveLast
        displayMenu
    Else
        displayMenu
    End If
End Sub
Private Sub cmdSaveProfile_Click()
    With rsProfile
        .Fields("Firstname") = Trim(txtFName)
        .Fields("Lastname") = Trim(txtLName)
        .Fields("Username") = Trim(txtUsername)
        .Fields("ContactNo") = Trim(txtContact)
        .Update
        MsgBox "Profile update successfully.", vbInformation, "Success"
    End With
End Sub

Private Sub Form_Load()
    
    cbPaymentMethod.AddItem "COD", 0
    cbPaymentMethod.AddItem "GCASH", 1
    cbPaymentMethod.Text = cbPaymentMethod.List(0)
    
    Category = "BREAKFAST"
    With lstOrder.ColumnHeaders
        .Add , , "Item Code", lstOrder.Width / 4
        .Add , , "Order Name", lstOrder.Width / 4
        .Add , , "Quantity", lstOrder.Width / 4
        .Add , , "Price", lstOrder.Width / 4
    End With
    With lstPurchases.ColumnHeaders
        .Add , , "Date", lstPurchases.Width / 6
        .Add , , "Order Status", lstPurchases.Width / 6
        .Add , , "Order Details", lstPurchases.Width / 2
        .Add , , "Order Amount", lstPurchases.Width / 6
        .Add , , "Order Item", lstPurchases.Width / 6
        .Add , , "Payment Method", lstPurchases.Width / 6
    End With
    
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\Group2.mdb;Persist Security Info=False"
    rs.Open "Select * from menuTab WHERE ProductCategory = '" & Category & "'", con, adOpenDynamic, adLockPessimistic
    rsPurchaces.Open "Select * from purchaseTab WHERE CustomerUsername = '" & Username & "'", con, adOpenDynamic, adLockPessimistic
    rsProfile.Open "Select Firstname + ' ' + Lastname As Fullname, loginTab.Firstname, loginTab.Lastname, loginTab.Username, loginTab.ContactNo from loginTab WHERE Username = '" & Username & "'", con, adOpenDynamic, adLockPessimistic
    
    displayMenu
    displayPurchaces
    displayUserInfo
End Sub

Sub displayUserInfo()
    If rsProfile.EOF Then
        Exit Sub
    End If
    
    txtFName.Text = rsProfile!Firstname
    txtLName.Text = rsProfile!Lastname
    txtUsername.Text = rsProfile!Username
    txtFullname.Caption = rsProfile!Fullname
    txtContact.Text = rsProfile!ContactNo
End Sub

Public Sub displayPurchaces()
    With rsPurchaces
        lstPurchases.ListItems.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
                lstPurchases.ListItems.Add , , !OrderDate
                lstPurchases.ListItems(lstPurchases.ListItems.Count).SubItems(1) = !OrderStatus
                lstPurchases.ListItems(lstPurchases.ListItems.Count).SubItems(2) = !orderDetails
                lstPurchases.ListItems(lstPurchases.ListItems.Count).SubItems(3) = !OrderAmount
                lstPurchases.ListItems(lstPurchases.ListItems.Count).SubItems(4) = !OrderItem
                lstPurchases.ListItems(lstPurchases.ListItems.Count).SubItems(5) = !OrderPayment
                .MoveNext
            Loop
        End If
    End With
End Sub
Sub displayMenu()
    If rs.EOF Then
       MsgBox "There is no " & Category & " availableb yet.", vbInformation, "Alert!"
       cmdBreakfast_Click
    Else
       txtProductName = rs!ProductName
       txtCategory.Caption = rs!ProductCategory
       txtProductDescription.Caption = rs!ProductDescription
       txtPriceDesc.Caption = "Price: PHP" & rs!Price
       
       If Dir(rs!Photo) <> "" Then
            menuPhoto.Picture = LoadPicture(rs!Photo)
        Else
            menuPhoto.Picture = LoadPicture(App.Path & rs!Photo)
       End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    refreshPurchases
End Sub

Public Sub refreshPurchases()
    Set rsPurchaces = Nothing
    rsPurchaces.Open "Select * from purchaseTab WHERE CustomerUsername = '" & Username & "'", con, adOpenDynamic, adLockPessimistic
    displayPurchaces
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  rsProfile.Close
  rsPurchaces.Close
  rs.Close
  con.Close
End Sub
