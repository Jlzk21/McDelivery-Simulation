VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Management 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Management Center"
   ClientHeight    =   13560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17760
   Icon            =   "Management.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13560
   ScaleWidth      =   17760
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   11625
      Left            =   4020
      TabIndex        =   3
      Top             =   1800
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   20505
      _Version        =   393216
      TabHeight       =   706
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Manage Foods"
      TabPicture(0)   =   "Management.frx":3C3A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "UsernameFrm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAdd"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "filePicker"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdEdit"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdDel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCancel"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSave"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdClear"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Manage Orders"
      TabPicture(1)   =   "Management.frx":3C56
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Account"
      TabPicture(2)   =   "Management.frx":3C72
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtClientID"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   11205
         Left            =   -75030
         TabIndex        =   24
         Top             =   420
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   19764
         _Version        =   393216
         TabOrientation  =   1
         TabsPerRow      =   5
         TabHeight       =   706
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Order Request"
         TabPicture(0)   =   "Management.frx":3C8E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstPending"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdAction"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdReject"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame9"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Check1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Completed"
         TabPicture(1)   =   "Management.frx":3CAA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstCompleted"
         Tab(1).Control(1)=   "Frame10"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Canceled"
         TabPicture(2)   =   "Management.frx":3CC6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "treeCanceled"
         Tab(2).ControlCount=   1
         Begin MSComctlLib.TreeView treeCanceled 
            Height          =   10245
            Left            =   -74700
            TabIndex        =   45
            Top             =   300
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   18071
            _Version        =   393217
            Style           =   7
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
         End
         Begin VB.Frame Frame10 
            Caption         =   "Order Description:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   10455
            Left            =   -65730
            TabIndex        =   40
            Top             =   120
            Width           =   4335
            Begin VB.Frame Frame12 
               Caption         =   "Order Details"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7215
               Left            =   120
               TabIndex        =   43
               Top             =   270
               Width           =   4095
               Begin MSComctlLib.ImageList ImageList2 
                  Left            =   3240
                  Top             =   510
                  _ExtentX        =   1005
                  _ExtentY        =   1005
                  BackColor       =   -2147483643
                  MaskColor       =   12632256
                  _Version        =   393216
               End
               Begin MSComctlLib.TreeView orderTree 
                  Height          =   6615
                  Index           =   1
                  Left            =   210
                  TabIndex        =   44
                  Top             =   390
                  Width           =   3705
                  _ExtentX        =   6535
                  _ExtentY        =   11668
                  _Version        =   393217
                  HideSelection   =   0   'False
                  LabelEdit       =   1
                  Style           =   7
                  ImageList       =   "ImageList1"
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin VB.Frame Frame11 
               Caption         =   "Customer Information:"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2775
               Left            =   120
               TabIndex        =   41
               Top             =   7560
               Width           =   4095
               Begin VB.Label txtCustomerInfo 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2205
                  Index           =   1
                  Left            =   210
                  TabIndex        =   42
                  Top             =   420
                  Width           =   3765
               End
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Select All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   38
            Top             =   1080
            Width           =   2265
         End
         Begin VB.Frame Frame9 
            Caption         =   "Order Status"
            Height          =   705
            Left            =   180
            TabIndex        =   34
            Top             =   270
            Width           =   8925
            Begin VB.CommandButton cmdReceive 
               Caption         =   "To  Recieve"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   5190
               TabIndex        =   37
               Top             =   210
               Width           =   2445
            End
            Begin VB.CommandButton cmdDeliver 
               Caption         =   "To Deliver"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   2640
               TabIndex        =   36
               Top             =   210
               Width           =   2445
            End
            Begin VB.CommandButton cmdPending 
               Caption         =   "Pending"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   90
               TabIndex        =   35
               Top             =   210
               Width           =   2445
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Order Description:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9285
            Left            =   9270
            TabIndex        =   28
            Top             =   210
            Width           =   4335
            Begin VB.Frame Frame8 
               Caption         =   "Customer Information:"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2775
               Left            =   120
               TabIndex        =   30
               Top             =   6210
               Width           =   4095
               Begin VB.Label txtCustomerInfo 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2205
                  Index           =   0
                  Left            =   210
                  TabIndex        =   31
                  Top             =   420
                  Width           =   3765
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Order Details"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5895
               Left            =   120
               TabIndex        =   29
               Top             =   270
               Width           =   4095
               Begin MSComctlLib.ImageList ImageList1 
                  Left            =   3240
                  Top             =   510
                  _ExtentX        =   1005
                  _ExtentY        =   1005
                  BackColor       =   -2147483643
                  MaskColor       =   12632256
                  _Version        =   393216
               End
               Begin MSComctlLib.TreeView orderTree 
                  Height          =   5265
                  Index           =   0
                  Left            =   210
                  TabIndex        =   32
                  Top             =   390
                  Width           =   3705
                  _ExtentX        =   6535
                  _ExtentY        =   9287
                  _Version        =   393217
                  HideSelection   =   0   'False
                  LabelEdit       =   1
                  Style           =   7
                  ImageList       =   "ImageList1"
                  BorderStyle     =   1
                  Appearance      =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
         End
         Begin VB.CommandButton cmdReject 
            Caption         =   "CANCEL ORDER"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   9780
            Width           =   2655
         End
         Begin VB.CommandButton cmdAction 
            Caption         =   "ACCEPT ORDER"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   9780
            Width           =   2655
         End
         Begin MSComctlLib.ListView lstPending 
            Height          =   8385
            Left            =   150
            TabIndex        =   25
            Top             =   1350
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   14790
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lstCompleted 
            Height          =   10485
            Left            =   -74850
            TabIndex        =   39
            Top             =   150
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   18494
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Fields"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3600
         TabIndex        =   23
         Top             =   3930
         Width           =   2800
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10680
         TabIndex        =   22
         Top             =   10950
         Width           =   2445
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8070
         TabIndex        =   21
         Top             =   10950
         Width           =   2505
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "DELETE MENU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5490
         TabIndex        =   20
         Top             =   10950
         Width           =   2475
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "EDIT MENU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2910
         TabIndex        =   19
         Top             =   10950
         Width           =   2475
      End
      Begin VB.Frame Frame4 
         Caption         =   " Menu"
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
         Height          =   6225
         Left            =   330
         TabIndex        =   17
         Top             =   4620
         Width           =   12795
         Begin MSComctlLib.ListView lstMenu 
            Height          =   5805
            Left            =   150
            TabIndex        =   18
            Top             =   270
            Width           =   12465
            _ExtentX        =   21987
            _ExtentY        =   10239
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin MSComDlg.CommonDialog filePicker 
         Left            =   11130
         Top             =   2220
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "Menu Photo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3645
         Left            =   9510
         TabIndex        =   15
         Top             =   870
         Width           =   3615
         Begin VB.CommandButton cmdUploadPhoto 
            Caption         =   "Upload File"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   150
            TabIndex        =   16
            Top             =   3000
            Width           =   3345
         End
         Begin VB.Image imgMenuPhoto 
            Height          =   2625
            Left            =   90
            Stretch         =   -1  'True
            Top             =   240
            Width           =   3405
         End
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search Menu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6540
         TabIndex        =   14
         Top             =   3930
         Width           =   2800
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD MENU"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   360
         TabIndex        =   13
         Top             =   10950
         Width           =   2445
      End
      Begin VB.Frame UsernameFrm 
         Caption         =   "Product Name: "
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
         Left            =   360
         TabIndex        =   11
         Top             =   900
         Width           =   5805
         Begin VB.TextBox txtProductName 
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
            TabIndex        =   12
            Top             =   270
            Width           =   5355
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Product Description: "
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
         Height          =   2145
         Left            =   330
         TabIndex        =   9
         Top             =   1620
         Width           =   9045
         Begin VB.TextBox txtProductDesc 
            BackColor       =   &H8000000F&
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
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   270
            Width           =   8745
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Price: "
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
         Left            =   6390
         TabIndex        =   7
         Top             =   900
         Width           =   3000
         Begin VB.TextBox txtPrice 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2715
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Categories: "
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
         Height          =   705
         Left            =   330
         TabIndex        =   5
         Top             =   3840
         Width           =   3225
         Begin VB.ComboBox cbCategories 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label txtClientID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -74160
         TabIndex        =   33
         Top             =   1530
         Width           =   2445
      End
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
      Left            =   -30
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   12600
      Width           =   3855
   End
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
      Left            =   -60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1980
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Management Center"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1125
      Left            =   6930
      TabIndex        =   4
      Top             =   270
      Width           =   7500
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   3825
      Y1              =   1890
      Y2              =   1905
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
      Left            =   1290
      TabIndex        =   0
      Top             =   960
      Width           =   2025
   End
   Begin VB.Image imgUser 
      Height          =   900
      Left            =   270
      Picture         =   "Management.frx":3CE2
      Stretch         =   -1  'True
      Top             =   660
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   13545
      Left            =   0
      Top             =   -30
      Width           =   3795
   End
End
Attribute VB_Name = "Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Username As String
Dim conn As New ADODB.Connection
Dim rstMenu As New ADODB.Recordset
Dim rstUser As New ADODB.Recordset
Dim rstRequest As New ADODB.Recordset
Dim rstAddress As New ADODB.Recordset
Dim confirm As Integer
Dim photoPath As String
Dim AddingBA As Boolean
Dim itemCode As String
Dim statusCode As Integer
Dim lstClickCode As Integer

Private Sub Check1_Click()
    Dim i As Integer
    
    With lstPending
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = Check1.Value
            .ListItems(i).Selected = Check1.Value
        Next i
    End With
End Sub

Private Sub cmdAction_Click()
    Dim i As Integer
    With lstPending
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                rstRequest.Close
                rstRequest.Open "select * from purchaseTab WHERE ID=" & .ListItems(i).SubItems(1), conn, adOpenDynamic, adLockPessimistic
                    If statusCode = 1 Then
                             rstRequest.Fields("OrderStatus") = "To Deliver"
                        ElseIf statusCode = 2 Then
                             rstRequest.Fields("OrderStatus") = "To Receive"
                        ElseIf statusCode = 3 Then
                             rstRequest.Fields("OrderStatus") = "Completed"
                    End If
                rstRequest.Update
            End If
        Next i
    End With
    displayRequest
    displayCanceled
    displayCompleted
End Sub

Private Sub cmdReject_Click()
        Dim i As Integer
        With lstPending
                Dim response As String
                confirm = MsgBox("Do you want to reject order ", vbYesNo + vbCritical, "Cancel Confirmation")
                response = InputBox("Due to:", "Cancel Message")
        
                If confirm = vbYes Then
                    For i = 1 To .ListItems.Count
                        If .ListItems(i).Checked = True Then
                            rstRequest.Close
                            rstRequest.Open "select * from purchaseTab WHERE ID=" & .ListItems(i).SubItems(1), conn, adOpenDynamic, adLockPessimistic
                            rstRequest.Fields("OrderStatus") = "Canceled"
                            rstRequest.Fields("Response") = response
                            rstRequest.Update
                        End If
                    Next i
                End If
        End With
        displayRequest
        displayCanceled
End Sub

Private Sub cmdDel_Click()
    confirm = MsgBox("Do you want to delete itemCode: " & itemCode, vbYesNo + vbCritical, "Deletion Confirmation")
    If confirm = vbYes Then
        rstMenu.Delete adAffectCurrent
        MsgBox "Record has been Deleted successfully", vbInformation, "Message"
        rstMenu.Update
        refreshdata
            Else
        MsgBox "Menu Not Deleted ..!!", vbInformation, "Message"
    End If
    
    End Sub
    Sub refreshdata()
        rstMenu.Close
        rstMenu.Open "Select * from menuTab", conn, adOpenStatic, adLockPessimistic
        If Not rstMenu.EOF Then
            rstMenu.MoveNext
            displayMenu
            cmdClear_Click
        Else
            MsgBox "No Record Found"
        End If
End Sub

Private Sub cmdExit_Click()
    End
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
Private Sub cmdDeliver_Click()
    statusCode = 2
    displayRequest
    cmdReject.Visible = False
    cmdAction.Caption = "MARK AS TO RECIEVE"
End Sub

Private Sub cmdPending_Click()
    statusCode = 1
    displayRequest
    cmdReject.Visible = True
    cmdAction.Caption = "ACCEPT"
End Sub
Private Sub cmdReceive_Click()
    statusCode = 3
    displayRequest
    cmdReject.Visible = False
    cmdAction.Caption = "MARK AS DONE"
End Sub


Private Sub cmdSearch_Click()
    Dim itemCode As String
    itemCode = InputBox("Please enter Item Code To be Searched:", "Search Menu")
    
    Set rstMenu = Nothing
    rstMenu.Open "select * from menuTab where ITEMCODE=" & itemCode, conn
    
    If rstMenu.EOF Then
       MsgBox "Item ID not found.", vbInformation, "Search"
    Else
        txtProductName.Text = rstMenu!ProductName
        txtProductDesc.Text = rstMenu!ProductDescription
        cbCategories.Text = rstMenu!ProductCategory
        txtPrice.Text = rstMenu!Price
        
        photoPath = rstMenu!Photo
        imgMenuPhoto.Visible = True
        If Dir(photoPath) <> "" Then
            imgMenuPhoto.Picture = LoadPicture(photoPath)
        Else
            imgMenuPhoto.Picture = LoadPicture(App.Path & photoPath)
       End If
    End If
       
End Sub
Private Sub cmdAdd_Click()
    AddingBA = True
    cmdClear_Click
    txtProductName.SetFocus
    buttonValidator False
End Sub

Private Sub cmdEdit_Click()
    AddingBA = False
    txtProductName.SetFocus
    buttonValidator False
    cmdSearch.Enabled = True
End Sub

Private Sub cmdCancel_Click()
  AddingBA = False
  buttonValidator True
End Sub

Private Sub buttonValidator(e As Boolean)
  cmdSearch.Enabled = e
  cmdAdd.Enabled = e
  cmdEdit.Enabled = e
  cmdDel.Enabled = e
  lstMenu.Enabled = e
  cmdSave.Enabled = Not e
  cmdCancel.Enabled = Not e
End Sub

Private Sub cmdClear_Click()
    txtProductName.Text = ""
    txtProductDesc.Text = ""
    txtPrice.Text = ""
    photoPath = ""
    imgMenuPhoto.Visible = False
End Sub

Private Sub cmdUploadPhoto_Click()
    filePicker.Filter = "(*.jpg)"
    filePicker.DefaultExt = "jpg"
    filePicker.DialogTitle = "Choose Picture"
    filePicker.ShowOpen
    photoPath = filePicker.FileName
    imgMenuPhoto.Visible = True
    imgMenuPhoto.Picture = LoadPicture(filePicker.FileName)
End Sub

Private Function textValidator() As Boolean
    textValidator = Trim(txtProductName) <> "" And Trim(txtProductDesc) <> "" And Trim(txtPrice) <> "" And Trim(photoPath) <> ""
End Function

Private Sub cmdSave_Click()
    If Not textValidator Then
        MsgBox "Incomplete data. Please fillup all fields.", vbExclamation, "Warning!"
        Exit Sub
    End If
    
    If AddingBA Then
            rstMenu.Close
            rstMenu.Open "select * from menuTab", conn
            rstMenu.AddNew
            rstMenu.Fields("ProductName").Value = Trim(txtProductName)
            rstMenu.Fields("ProductDescription").Value = Trim(txtProductDesc)
            rstMenu.Fields("ProductCategory").Value = Trim(cbCategories)
            rstMenu.Fields("Price").Value = Trim(txtPrice)
            rstMenu.Fields("Photo").Value = photoPath
            rstMenu.Update
            MsgBox "Menu saved successfully ..!!!", vbInformation
        Else
            rstMenu.Close
            rstMenu.Open "select * from menuTab where ITEMCODE=" & itemCode, conn
    
            rstMenu.Fields("ProductName").Value = Trim(txtProductName)
            rstMenu.Fields("ProductDescription").Value = Trim(txtProductDesc)
            rstMenu.Fields("ProductCategory").Value = Trim(cbCategories)
            rstMenu.Fields("Price").Value = Trim(txtPrice)
            rstMenu.Fields("Photo").Value = photoPath
            rstMenu.Update
            MsgBox "Menu saved successfully Updated.", vbInformation, "Success"
    End If
    
    buttonValidator True
    displayMenu
End Sub

Sub displayMenu()
    rstMenu.Close
    rstMenu.Open "Select * from menuTab", conn, adOpenDynamic, adLockPessimistic
    With rstMenu
        lstMenu.ListItems.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
                lstMenu.ListItems.Add , , !itemCode
                lstMenu.ListItems(lstMenu.ListItems.Count).SubItems(1) = !ProductName
                lstMenu.ListItems(lstMenu.ListItems.Count).SubItems(2) = !ProductDescription
                lstMenu.ListItems(lstMenu.ListItems.Count).SubItems(3) = !ProductCategory
                lstMenu.ListItems(lstMenu.ListItems.Count).SubItems(4) = !Price
                .MoveNext
            Loop
        End If
    End With
End Sub
Public Sub displayRequest()
    rstRequest.Close
   
    If statusCode = 1 Then
            rstRequest.Open "select * from purchaseTab WHERE OrderStatus='Pending'", conn, adOpenDynamic, adLockPessimistic
        ElseIf statusCode = 2 Then
            rstRequest.Open "select * from purchaseTab WHERE OrderStatus='To Deliver'", conn, adOpenDynamic, adLockPessimistic
        ElseIf statusCode = 3 Then
            rstRequest.Open "select * from purchaseTab WHERE OrderStatus='To Receive'", conn, adOpenDynamic, adLockPessimistic
        Else
            rstRequest.Open "select * from purchaseTab WHERE OrderStatus='Pending'", conn, adOpenDynamic, adLockPessimistic
    End If
    
    With rstRequest
        lstPending.ListItems.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
                lstPending.ListItems.Add , , !OrderDate
                lstPending.ListItems(lstPending.ListItems.Count).SubItems(1) = !ID
                lstPending.ListItems(lstPending.ListItems.Count).SubItems(2) = !OrderStatus
                lstPending.ListItems(lstPending.ListItems.Count).SubItems(3) = !OrderAmount
                lstPending.ListItems(lstPending.ListItems.Count).SubItems(4) = !OrderPayment
                .MoveNext
            Loop
        End If
    End With
End Sub
Public Sub displayCompleted()
    rstRequest.Close
    rstRequest.Open "select * from purchaseTab WHERE OrderStatus='Completed'", conn, adOpenDynamic, adLockPessimistic
    
    With rstRequest
        lstCompleted.ListItems.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
                lstCompleted.ListItems.Add , , !OrderDate
                lstCompleted.ListItems(lstCompleted.ListItems.Count).SubItems(1) = !ID
                lstCompleted.ListItems(lstCompleted.ListItems.Count).SubItems(2) = !OrderStatus
                lstCompleted.ListItems(lstCompleted.ListItems.Count).SubItems(3) = !OrderAmount
                lstCompleted.ListItems(lstCompleted.ListItems.Count).SubItems(4) = !OrderPayment
                .MoveNext
            Loop
        End If
    End With
End Sub
Public Sub displayCanceled()
    Dim i As Integer

    rstRequest.Close
    rstRequest.Open "select * from purchaseTab WHERE OrderStatus='Canceled'", conn, adOpenDynamic, adLockPessimistic
    
    With rstRequest
        treeCanceled.Nodes.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
            
                treeCanceled.Nodes.Add , , "node1key" & Str(!ID), "[Canceled] Order Id:" & !ID
                treeCanceled.Nodes.Add "node1key" & Str(!ID), tvwChild, "child1key" & i & Str(!ID), "Due to: " & !response
                treeCanceled.Nodes.Add "node1key" & Str(!ID), tvwChild, "child2key" & i & Str(!ID), "Order Status: " & !OrderStatus
                treeCanceled.Nodes.Add "node1key" & Str(!ID), tvwChild, "child3key" & i & Str(!ID), "Order Details: " & !orderDetails
                treeCanceled.Nodes.Add "node1key" & Str(!ID), tvwChild, "child4key" & i & Str(!ID), "Order Amount: " & !OrderAmount
                
                .MoveNext
                i = i + 1
            Loop
        End If
        
    End With
End Sub
Private Sub Form_Load()
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\Group2.mdb;Persist Security Info=False"
    rstMenu.Open "Select * from menuTab", conn, adOpenDynamic, adLockPessimistic
    rstUser.Open "Select Firstname + ' ' + Lastname As Fullname,Username from loginTab WHERE Username='" & Username & "'", conn, adOpenDynamic, adLockPessimistic
    rstRequest.Open "Select * from purchaseTab where OrderStatus='Pending'", conn, adOpenDynamic, adLockPessimistic
    rstAddress.Open "Select * from addressTab", conn, adOpenDynamic, adLockPessimistic
    statusCode = 1
    
    txtFullname.Caption = rstUser!Fullname
    txtClientID.Caption = txtClientID.Caption & StringToHex(rstUser!Username)
    cbCategories.AddItem "BREAKFAST", 0
    cbCategories.AddItem "BURGERS", 1
    cbCategories.AddItem "CHICKENANDPLATERS", 2
    cbCategories.AddItem "DRINKSANDDESSERTS", 3
    cbCategories.AddItem "MCCAFE", 4
    cbCategories.AddItem "FRIES", 5
    cbCategories.Text = cbCategories.List(0)
    
    cmdAction.Caption = "ACCEPT"
    
    With lstMenu.ColumnHeaders
        .Add , , "Item Code", lstMenu.Width / 6
        .Add , , "Name", lstMenu.Width / 6
        .Add , , "Description", lstMenu.Width / 2
        .Add , , "Category", lstMenu.Width / 6
        .Add , , "Price", lstMenu.Width / 6
    End With
    
    With lstPending.ColumnHeaders
        .Add , , "Date", lstPending.Width / 5
        .Add , , "Order ID", lstPending.Width / 5
        .Add , , "Order Status", lstPending.Width / 5
        .Add , , "Order Amount", lstPending.Width / 5
        .Add , , "Payment Method", lstPending.Width / 5
    End With
    With lstCompleted.ColumnHeaders
        .Add , , "Date", lstPending.Width / 5
        .Add , , "Order ID", lstPending.Width / 5
        .Add , , "Order Status", lstPending.Width / 5
        .Add , , "Order Amount", lstPending.Width / 5
        .Add , , "Payment Method", lstPending.Width / 5
    End With
    displayMenu
    displayRequest
    displayCompleted
    displayCanceled
    buttonValidator True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    rstAddress.Close
    rstRequest.Close
    rstUser.Close
    rstMenu.Close
    conn.Close
End Sub

Private Sub lstMenu_ItemClick(ByVal item As MSComctlLib.ListItem)
    itemCode = item.Text
    
    rstMenu.Close
    rstMenu.Open "select * from menuTab where ITEMCODE=" & itemCode, conn
    
    If rstMenu.EOF Then
       MsgBox "Item ID not found.", vbInformation, "Search"
    Else
        txtProductName.Text = rstMenu!ProductName
        txtProductDesc.Text = rstMenu!ProductDescription
        cbCategories.Text = rstMenu!ProductCategory
        txtPrice.Text = rstMenu!Price
        
        photoPath = rstMenu!Photo
        imgMenuPhoto.Visible = True
        
        If Dir(photoPath) <> "" Then
            imgMenuPhoto.Picture = LoadPicture(photoPath)
        Else
            imgMenuPhoto.Picture = LoadPicture(App.Path & photoPath)
       End If
    End If
End Sub
Private Sub lstCompleted_ItemClick(ByVal item As MSComctlLib.ListItem)
    lstClickCode = 1
    listClickAction item
End Sub

Private Sub lstPending_ItemClick(ByVal item As MSComctlLib.ListItem)
    lstClickCode = 0
    listClickAction item
End Sub

Sub listClickAction(ByVal item As MSComctlLib.ListItem)
    Dim orderDetails() As String
    Dim i As Long
    
    rstRequest.Close
    rstRequest.Open "select * from purchaseTab where ID=" & item.SubItems(1), conn, adOpenDynamic, adLockPessimistic
    If rstRequest.EOF Then
       MsgBox "Pending req not found.", vbInformation, "Search"
    Else
       rstUser.Close
       rstUser.Open "select Firstname + ' ' + Lastname As Fullname from loginTab where Username='" & rstRequest!CustomerUsername & "'", conn
       rstAddress.Close
       rstAddress.Open "select Region + ', ' + City + ', ' + District + ', ' + Street As CompleteAddress,AdditionalInfo,PhoneNumber from addressTab where Username='" & rstRequest!CustomerUsername & "'", conn
       txtCustomerInfo(lstClickCode).Caption = "Customer Name: " & rstUser!Fullname & Chr$(13) & Chr$(10) _
       & "Contact Number: " & rstAddress!PhoneNumber & Chr$(13) & Chr$(10) _
       & "Customer Address: " & rstAddress!CompleteAddress & Chr$(13) & Chr$(10) _
       & "Addtional Information: " & rstAddress!AdditionalInfo
       orderDetails = Split(rstRequest!orderDetails, ", ")
       
       orderTree(lstClickCode).Nodes.Clear
       orderTree(lstClickCode).Nodes.Add , , "node1key" & Str(rstRequest!ID), "Order Item"
       
       For i = 0 To UBound(orderDetails)
          orderTree(lstClickCode).Nodes.Add "node1key" & Str(rstRequest!ID), tvwChild, "child1key" & i & Str(rstRequest!ID), Trim(orderDetails(i))
       Next
       
       orderTree(lstClickCode).Nodes(1).Expanded = True
       
       orderTree(lstClickCode).Nodes.Add , , "node2key" & Str(rstRequest!ID), "Order Amount"
       orderTree(lstClickCode).Nodes.Add "node2key" & Str(rstRequest!ID), tvwChild, "child2key" & Str(rstRequest!ID), rstRequest!OrderAmount
       orderTree(lstClickCode).Nodes(orderTree(lstClickCode).Nodes.Count - 1).Expanded = True
    End If
End Sub
Public Function StringToHex(ByVal StrToHex As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim i         As Long
        For i = 1 To Len(StrToHex)
            strTemp = Hex$(Asc(Mid$(StrToHex, i, 1)))
            If Len(strTemp) = 1 Then strTemp = "0" & strTemp
            strReturn = strReturn & strTemp
        Next i
        StringToHex = strReturn
End Function
