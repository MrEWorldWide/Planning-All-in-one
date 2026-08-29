VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form UPCOMING 
   Caption         =   "Upcoming Sales Orders"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton UPCcancelbut 
      Caption         =   "Exit"
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton UPCprintbut 
      Caption         =   "Print"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton UPCrefbut 
      Caption         =   "Refresh"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   1
      Top             =   4920
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid UPCOMgrid 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "UPCOMING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This displays upcoming sales orders. what makes an upcoming sales order goes as follows:
'1) The sales order was entered within

Private Sub Form_Load()
    
    DoEvents
    
    SetupConnection

    SQLSelect = "SELECT Pub.OrderDtl.OrderNum, Pub.OrderDtl.OrderDate, Pub.OrderDtl.PartNum "
    SQLSelect = SQLSelect & "Pub.OrderDtl.OrderQty, Pub.OrderDtl.WarrantyCode, Pub.OrderDtl.NeedByDate, "
    SQLSelect = SQLSelect & "Pub.OrderDtl.QuoteLine, Pub.OrderDtl.QuoteNum, Pub.OrderDtl.RequestDate, "
    SQLSelect = SQLSelect & "Pub.OrderDtl.RevisionNum, Pub.OrderRel.OrderRelNum, Pub.OrderRel.ReqDate"
    SQLSelect = SQLSelect & "Pub.OrderRel.OurStockQty, "
    
    RSRead.Open (SQLSelect), CNVantage
    RSRead.MoveFirst
    If RSRead.EOF Then
        Unload KANBAN
        MsgBox "No Records Found.", vbOKOnly
        Unload Me
        Exit Sub
    End If
   
    With UPCOMgrid
        .WordWrap = True
        .Rows = 2
        .Cols = 7
        .FixedCols = 1
        .ColWidth(0) = 300
        .ColWidth(1) = 1300
        .ColWidth(2) = 400
        .ColWidth(5) = 550
        .ColWidth(6) = 1000
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(1) = flexAlignLeftCenter

        .TextMatrix(0, 1) = "OrderNum"
        .TextMatrix(0, 2) = "OrderDate"
        .TextMatrix(0, 3) = "PartNum"
        .TextMatrix(0, 4) = "OrderQty"
        .TextMatrix(0, 5) = "RelDate"
        .TextMatrix(0, 6) = "RelQty"
        .TextMatrix(0, 7) = "Inventory"
        
        c = 1
        Do Until rs.EOF

            
            .TextMatrix(c, 1) = RSRead.Fields(0)
            .TextMatrix(c, 2) = RSRead.Fields(1)
            .TextMatrix(c, 3) = RSRead.Fields(2)
            .TextMatrix(c, 4) = RSRead.Fields(3)
            .TextMatrix(c, 5) = RSRead.Fields(4)
            .TextMatrix(c, 6) = RSRead.Fields(5)
            rs.MoveNext
            .Rows = .Rows + 1
            c = c + 1
            
        Loop
        For c = 1 To .Rows - 1
            If c = .Rows Then Exit For
            .TextMatrix(c, 0) = c
        Next
    End With
    rs.Close
    vantage.Close
    UPCOMING.MousePointer = 0
End Sub

Private Sub UPCcancelbut_Click(Index As Integer)
 Unload UPCOMING
End Sub

Private Sub UPCprintbut_Click(Index As Integer)
    If UPCOMgrid.TextMatrix(1, 1) = "" Then
        MsgBox "No report to print."
        Exit Sub
    End If
End Sub

Private Sub UPCrefbut_Click(Index As Integer)
     UPCOMgrid.Rows = 0
    Call Form_Load
End Sub
