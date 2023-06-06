VERSION 5.00
Begin VB.Form PaymentView 
   BorderStyle     =   0  'None
   Caption         =   "Checkout"
   ClientHeight    =   12345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   Icon            =   "PaymentView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12345
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Payment Details"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   210
      TabIndex        =   10
      Top             =   8970
      Width           =   5835
      Begin VB.Label txtPaymentDetails 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   210
         TabIndex        =   11
         Top             =   360
         Width           =   5265
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order Details"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   210
      TabIndex        =   8
      Top             =   2760
      Width           =   5835
      Begin VB.Label txtOrderDetails 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5565
         Left            =   210
         TabIndex        =   9
         Top             =   360
         Width           =   5265
      End
   End
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
      Left            =   180
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   1275
   End
   Begin VB.CommandButton cmdPlaceOrder 
      Caption         =   "Place Order"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   11520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Delivery Address"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   0
      Top             =   1230
      Width           =   5865
      Begin VB.CommandButton cmdChangeAddress 
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4710
         TabIndex        =   5
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label txtDeliveryAddress 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Label txtTotalPayment 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2490
      TabIndex        =   7
      Top             =   11850
      Width           =   1065
   End
   Begin VB.Label txtTotalPaymentTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2490
      TabIndex        =   6
      Top             =   11550
      Width           =   1110
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
      Left            =   1740
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -30
      Top             =   0
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -30
      Top             =   11280
      Width           =   6375
   End
End
Attribute VB_Name = "PaymentView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Public Username As String
Dim AddressID As String
Dim orderCount As Integer
Private Sub cmdBack_Click()
    Unload Me
End Sub
Private Sub cmdChangeAddress_Click()
    CustomerAddress.Username = Username
    CustomerAddress.Show
End Sub

Private Sub cmdPlaceOrder_Click()
    Dim itemX As ListItem
    Dim orderDetails As String
    Dim orderCodes As String
    
    Set rst = Nothing
    rst.Open "Select * from purchaseTab WHERE CustomerUsername='" & Username & "'", conn, adOpenDynamic, adLockPessimistic
    rst.AddNew
    
    rst.Fields("OrderDate").Value = Format(Now, "mm/dd/yy hh:mm")
    rst.Fields("OrderStatus").Value = "Pending"
    
    For Each itemX In CustomerView.lstOrder.ListItems
         orderDetails = orderDetails & " " & itemX.SubItems(1) & "(" & itemX.SubItems(2) & "x),"
         orderCodes = orderCodes & itemX.Text & ","
    Next
    
    rst.Fields("OrderDetails").Value = orderDetails
    
    rst.Fields("OrderAmount").Value = txtTotalPayment.Caption
    rst.Fields("OrderItem").Value = Str(orderCount)
    rst.Fields("OrderAddress").Value = AddressID
    rst.Fields("OrderPayment").Value = CustomerView.cbPaymentMethod.Text
    rst.Fields("CustomerUsername").Value = Trim(Username)
    rst.Fields("OrderCodes").Value = orderCodes
    MsgBox "Order purchase successfully ..!!!", vbInformation
    rst.Update
    
    CustomerView.Refresh
    Unload Me
   
End Sub

Private Sub Form_Load()
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\Group2.mdb;Persist Security Info=False"
    displayAddress
    displayOrderDetails
    displayPaymentDetails
End Sub

Sub displayAddress()
    Set rst = Nothing
    rst.Open "Select * from addressTab WHERE Username='" & Username & "'", conn, adOpenDynamic, adLockPessimistic
    txtDeliveryAddress.Caption = ""
    If Not rst.EOF Then
        AddressID = rst!ID
        txtDeliveryAddress.Caption = rst!Fullname & " | " & rst!PhoneNumber & Chr$(13) & Chr$(10) & _
        rst!Street & Chr$(13) & Chr$(10) & _
        rst!District & ", " & rst!City & ", " & rst!Region
    End If
End Sub
Sub displayOrderDetails()
    Dim itemX As ListItem
    txtOrderDetails.Caption = ""
    For Each itemX In CustomerView.lstOrder.ListItems
        txtOrderDetails.Caption = txtOrderDetails.Caption & Chr$(13) & Chr$(10) & itemX.SubItems(1) & "(" & itemX.SubItems(2) & "x)" & " - PHP " & itemX.SubItems(3)
        orderCount = orderCount + 1
    Next
    txtOrderDetails.Caption = txtOrderDetails.Caption & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "Order Total(" & Str(orderCount) & " Item): PHP " & CustomerView.txtAmount
End Sub
Sub displayPaymentDetails()
    txtPaymentDetails.Caption = "Order Subtotal: PHP " & CustomerView.txtAmount & Chr$(13) & Chr$(10) & _
    "Delivery Fee Subtotal: PHP " & CustomerView.txtDeliveryFee.Text & Chr$(13) & Chr$(10) & _
    "Total Payment: PHP " & Str(Val(CustomerView.txtDeliveryFee.Text) + Val(CustomerView.txtAmount)) & Chr$(13) & Chr$(10)
    txtTotalPayment.Caption = "PHP " & Str(Val(CustomerView.txtDeliveryFee.Text) + Val(CustomerView.txtAmount))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  rst.Close
  conn.Close
End Sub

