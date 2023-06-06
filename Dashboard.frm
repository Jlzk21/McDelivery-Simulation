VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Dashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   13560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17760
   Icon            =   "Dashboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13560
   ScaleWidth      =   17760
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   2895
      Left            =   10920
      OleObjectBlob   =   "Dashboard.frx":3C3A
      TabIndex        =   17
      Top             =   3960
      Width           =   6255
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2895
      Left            =   4200
      OleObjectBlob   =   "Dashboard.frx":5FA1
      TabIndex        =   16
      Top             =   3960
      Width           =   6375
   End
   Begin VB.CommandButton cmdManageUser 
      BackColor       =   &H000080FF&
      Caption         =   "Manage Users"
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
      TabIndex        =   13
      Top             =   2580
      Width           =   3855
   End
   Begin MSComctlLib.ListView lstReceivedOrders 
      Height          =   4785
      Left            =   4230
      TabIndex        =   15
      Top             =   7770
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   8440
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
      Top             =   3540
      Width           =   3855
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
      Left            =   -60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   12630
      Width           =   3855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received Orders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4260
      TabIndex        =   14
      Top             =   7290
      Width           =   2385
   End
   Begin VB.Image Image4 
      Height          =   840
      Left            =   11070
      Picture         =   "Dashboard.frx":82F7
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   900
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   14430
      Picture         =   "Dashboard.frx":11C73
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   7740
      Picture         =   "Dashboard.frx":1B036
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Here is our overall statistics."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   4320
      TabIndex        =   12
      Top             =   1380
      Width           =   4470
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hello Admin ,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4320
      TabIndex        =   11
      Top             =   870
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   4500
      Picture         =   "Dashboard.frx":23830
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label txtNetEarnings 
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   15510
      TabIndex        =   10
      Top             =   2910
      Width           =   1170
   End
   Begin VB.Label txtCustomerTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   12090
      TabIndex        =   9
      Top             =   2850
      Width           =   1410
   End
   Begin VB.Label txtOr 
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   8790
      TabIndex        =   8
      Top             =   2820
      Width           =   1410
   End
   Begin VB.Label txtOd 
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   5460
      TabIndex        =   7
      Top             =   2850
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Earnings"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   15480
      TabIndex        =   6
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Customer"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   12060
      TabIndex        =   5
      Top             =   2490
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Received"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8790
      TabIndex        =   4
      Top             =   2460
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Delivered"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5460
      TabIndex        =   3
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   14190
      Top             =   2070
      Width           =   3000
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   10890
      Top             =   2100
      Width           =   3000
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   7560
      Top             =   2070
      Width           =   3000
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   4260
      Top             =   2100
      Width           =   3000
   End
   Begin VB.Image imgUser 
      Height          =   900
      Left            =   330
      Picture         =   "Dashboard.frx":2BF4F
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
      Left            =   1350
      TabIndex        =   2
      Top             =   990
      Width           =   2025
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   13545
      Left            =   0
      Top             =   0
      Width           =   3795
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   13545
      Left            =   3780
      Top             =   0
      Width           =   13995
   End
End
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Username As String
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rstMenu As New ADODB.Recordset

Private Sub cmdManageUser_Click()
    ManageUsers.Show
End Sub

Private Sub Form_Load()
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\Group2.mdb;Persist Security Info=False"
    rst.Open "Select Firstname + ' ' + Lastname As Fullname,Username from loginTab WHERE Username='" & Username & "'", conn, adOpenDynamic, adLockPessimistic
    txtFullname.Caption = rst!Fullname
    
    With lstReceivedOrders.ColumnHeaders
        .Add , , "Customer", lstReceivedOrders.Width / 5
        .Add , , "Order ID", lstReceivedOrders.Width / 5
        .Add , , "Quantity", lstReceivedOrders.Width / 5
        .Add , , "Amount", lstReceivedOrders.Width / 5
        .Add , , "Status", lstReceivedOrders.Width / 5
    End With
    
    displayDelivered
    displayTopSelling
    displayActivity
    displayReceivedOrders
End Sub

Sub displayReceivedOrders()
    rst.Close
    rst.Open "select * from purchaseTab", conn, adOpenDynamic, adLockPessimistic
    
    With rst
        lstReceivedOrders.ListItems.Clear
        If .RecordCount <> 0 Then
            Do Until .EOF
                rstMenu.Open "Select Firstname + ' ' + Lastname As Fullname,Username from loginTab WHERE Username='" & !CustomerUsername & "'", conn, adOpenDynamic, adLockPessimistic
                If Not rstMenu.EOF Then
                    lstReceivedOrders.ListItems.Add , , rstMenu!Fullname
                    lstReceivedOrders.ListItems(lstReceivedOrders.ListItems.Count).SubItems(1) = !ID
                    lstReceivedOrders.ListItems(lstReceivedOrders.ListItems.Count).SubItems(2) = !OrderItem
                    lstReceivedOrders.ListItems(lstReceivedOrders.ListItems.Count).SubItems(3) = !OrderAmount
                    lstReceivedOrders.ListItems(lstReceivedOrders.ListItems.Count).SubItems(4) = !OrderStatus
                End If
                                
                rstMenu.Close
                
                .MoveNext
            Loop
        End If
    End With
End Sub

Sub displayTopSelling()
    Dim bfCount As Integer
    Dim bCount As Integer
    Dim cpCount As Integer
    Dim ddCount As Integer
    Dim mcCount As Integer
    Dim fCount As Integer
    
    rst.Close
    rst.Open "select * from purchaseTab where OrderStatus='Completed'", conn, adOpenDynamic, adLockPessimistic
    
    With rst
        bfCount = 0
        bCount = 0
        cpCount = 0
        ddCount = 0
        mcCount = 0
        fCount = 0
        If .RecordCount <> 0 Then
            Do Until .EOF
                Dim i As Integer
                Dim orderCodes() As String
                
                orderCodes = Split(!orderCodes, ",")
                
                For i = 0 To UBound(orderCodes) - 1
                    rstMenu.Open "select * from menuTab where ITEMCODE=" & orderCodes(i), conn, adOpenDynamic, adLockPessimistic
                    If rstMenu!ProductCategory = "BREAKFAST" Then
                        bfCount = bfCount + 1
                    ElseIf rstMenu!ProductCategory = "BURGERS" Then
                        bCount = bCount + 1
                    ElseIf rstMenu!ProductCategory = "CHICKENANDPLATERS" Then
                        cpCount = cpCount + 1
                    ElseIf rstMenu!ProductCategory = "DRINKSANDDESSERTS" Then
                        ddCount = ddCount + 1
                    ElseIf rstMenu!ProductCategory = "MCCAFE" Then
                        mcCount = mcCount + 1
                    ElseIf rstMenu!ProductCategory = "FRIES" Then
                        fCount = fCount + 1
                    End If
                    
                    rstMenu.Close
                Next
                
                .MoveNext
            Loop
        End If
    End With
    
    With MSChart2
        .ShowLegend = True
        .ColumnCount = 6
        .RowCount = 1
        .RowLabel = "Top Selling Items"
    End With
    
   With MSChart2
        .Column = 1
        .Row = 1
        .Data = bfCount
        .ColumnLabel = "BREAKFASTS"
   End With
   With MSChart2
        .Column = 2
        .Row = 1
        .Data = bCount
        .ColumnLabel = "BURGERS"
   End With
   With MSChart2
        .Column = 3
        .Row = 1
        .Data = cpCount
        .ColumnLabel = "CHICKEN AND PLATERS"
   End With
   With MSChart2
        .Column = 4
        .Row = 1
        .Data = ddCount
        .ColumnLabel = "DRINK AND DESSERTS"
   End With
   With MSChart2
        .Column = 5
        .Row = 1
        .Data = mcCount
        .ColumnLabel = "MC CAFE"
   End With
   With MSChart2
        .Column = 6
        .Row = 1
        .Data = fCount
        .ColumnLabel = "FRIES"
   End With
End Sub

Sub displayActivity()

    Dim bfCount As Integer
    Dim bCount As Integer
    Dim cpCount As Integer
    Dim ddCount As Integer
    Dim mcCount As Integer
    
    rst.Close
    rst.Open "select * from purchaseTab", conn, adOpenDynamic, adLockPessimistic
    
    With rst
        bfCount = 0
        bCount = 0
        cpCount = 0
        ddCount = 0
        mcCount = 0
        
        If .RecordCount <> 0 Then
            Do Until .EOF
                If !OrderStatus = "Pending" Then
                        bfCount = bfCount + 1
                    ElseIf !OrderStatus = "To Deliver" Then
                        bCount = bCount + 1
                    ElseIf !OrderStatus = "To Received" Then
                        cpCount = cpCount + 1
                    ElseIf !OrderStatus = "Completed" Then
                        ddCount = ddCount + 1
                    ElseIf !OrderStatus = "Canceled" Then
                        mcCount = mcCount + 1
                 End If
                .MoveNext
            Loop
        End If
    End With
    
    With MSChart1
        .ShowLegend = True
        .ColumnCount = 5
        .RowCount = 1
        .RowLabel = "Current Activity"
    End With
    
   With MSChart1
        .Column = 1
        .Row = 1
        .Data = bfCount
        .ColumnLabel = "Pendings"
   End With
   With MSChart1
        .Column = 2
        .Row = 1
        .Data = bCount
        .ColumnLabel = "To Deliver"
   End With
   With MSChart1
        .Column = 3
        .Row = 1
        .Data = cpCount
        .ColumnLabel = "To Revieced"
   End With
   With MSChart1
        .Column = 4
        .Row = 1
        .Data = ddCount
        .ColumnLabel = "Completed"
   End With
   With MSChart1
        .Column = 5
        .Row = 1
        .Data = mcCount
        .ColumnLabel = "Canceled"
   End With
End Sub

Sub displayDelivered()
 
    Dim i As Integer
    Dim earn As Long
    
    rst.Close
    rst.Open "select * from purchaseTab WHERE OrderStatus='To Received'", conn, adOpenDynamic, adLockPessimistic
    With rst
        i = 0
        If .RecordCount <> 0 Then
            Do Until .EOF
                .MoveNext
                i = i + 1
            Loop
        End If
    End With
    txtOd.Caption = Str(i)
    
    rst.Close
    rst.Open "select * from purchaseTab WHERE OrderStatus='Completed'", conn, adOpenDynamic, adLockPessimistic
    With rst
        i = 0
        If .RecordCount <> 0 Then
            Do Until .EOF
                .MoveNext
                i = i + 1
            Loop
        End If
    End With
    txtOr.Caption = Str(i)
    
    rst.Close
    rst.Open "select * from loginTab WHERE Role='CUSTOMER'", conn, adOpenDynamic, adLockPessimistic
    With rst
        i = 0
        If .RecordCount <> 0 Then
            Do Until .EOF
                .MoveNext
                i = i + 1
            Loop
        End If
    End With
    txtCustomerTotal.Caption = Str(i)
    
    rst.Close
    rst.Open "select * from purchaseTab WHERE OrderStatus='Completed'", conn, adOpenDynamic, adLockPessimistic
    With rst
        earn = 0
        If .RecordCount <> 0 Then
            Do Until .EOF
                Dim OrderAmount() As String
                OrderAmount = Split(!OrderAmount, "PHP")
                
                earn = earn + Val(OrderAmount(1))
               
                .MoveNext
            Loop
        End If
    End With
    txtNetEarnings.Caption = Str(earn)
    
    
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

Private Sub Form_Unload(Cancel As Integer)
    rst.Close
    conn.Close
End Sub
