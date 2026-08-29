VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MATDEC 
   Caption         =   "Partial Materials"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton MATprintbut 
      Caption         =   "Print"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid matdecgrid 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton MATcancelbut 
      Caption         =   "Exit"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton MATrefbut 
      Caption         =   "Refresh"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "MATDEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    DoEvents
    
    SetupConnection
    SQLSelect = "SELECT  Pub.JobMtl.PartNum, Pub.JobMtl.RequiredQty, Pub.JobMtl.AssemblySeq, Pub.JobMtl.MtlSeq, Pub.JobMtl.JobNum, Pub.JobHead.JobClosed "
    SQLSelect = SQLSelect & "FROM Pub.JobMtl INNER JOIN Pub.JobHead ON Pub.Part.PartNum=Pub.JobMtl.PartNum) "
    SQLSelect = SQLSelect & "ON (Pub.Part.PartNum = Pub.JobHead.PartNum) AND (Pub.JobHead.JobNum=Pub.JobMtl.JobNum) "
    SQLSelect = SQLSelect & "WHERE AND Pub.JobHead.JobClosed=0"
    Set RSRead = New ADODB.Recordset
    
    MATDEC.MousePointer = 11
    RSRead.Open (SQLSelect), CNVantage
    
    'if no records notify user then exit sub
    If RSRead.EOF Then
        MsgBox "No Purchase Direct items available."
        PURDIR.MousePointer = 0
        RSRead.Close
        CNVantage.Close
        Unload Me
        Exit Sub
    End If

    With matdecgrid
        .WordWrap = True
        .Rows = 2
        .Cols = 7
        .FixedCols = 1
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 400
        .ColWidth(3) = 1500
        .ColWidth(4) = 550
        .ColWidth(5) = 550
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(0, 1) = "Job Num"
        .TextMatrix(0, 2) = "Asm"
        .TextMatrix(0, 3) = "Mat Seq"
        .TextMatrix(0, 4) = "Part Num"
        .TextMatrix(0, 5) = "Req Qty"
        c = 1
        Do Until RSRead.EOF
            .TextMatrix(c, 0) = c
            If Int(RSRead.Fields(1)) - RSRead.Fields(1) <> 0 Then
            .TextMatrix(c, 1) = RSRead.Fields(4)
            .TextMatrix(c, 2) = RSRead.Fields(2)
            .TextMatrix(c, 3) = RSRead.Fields(3)
            .TextMatrix(c, 4) = RSRead.Fields(0)
            .TextMatrix(c, 6) = RSRead.Fields(1)
            RSRead.MoveNext
            Else
            RSRead.MoveNext
            End If
            .Rows = .Rows + 1
            c = c + 1
        Loop
        .Rows = .Rows - 1
    End With
    
    RSRead.Close
    CNVantage.Close
    MATDEC.MousePointer = 0
    
End Sub

Private Sub MATcancelbut_Click(Index As Integer)
    matdecgrid.Rows = 0
    Unload Me
End Sub

Private Sub MATprintbut_Click(Index As Integer)
    
    Dim lc, i As Long
    Dim header As String
    Dim col1x, col2x, col3x, col4x, col5x, ppage As Integer
    
    ppage = 1
    header = "  Page:" & ppage & "  " & Now() & "      Origin: " & Environ("USERNAME")
    header = header & "-----------------------------------------------------------------------"
    
    If bpogrid.Cols = 0 Then
        MsgBox "No report to print"
        Exit Sub
    End If
    
    Printer.FontName = "Verdana"
    Printer.FontSize = 11
    Printer.Print
    Printer.Print header
    Printer.Print
    Printer.Print
    
    lc = 0
    col1x = 300
    col2x = 900
    col3x = 3000
    col4x = 3300
    col5x = 3600
    col6x = 4900
    
    Printer.CurrentY = 1550
    Printer.CurrentX = col2x
    Printer.Print "JobNum";
    
    Printer.CurrentX = col3x
    Printer.Print "ASM"
    
    Printer.CurrentX = col4x
    Printer.Print "MatSeq"
    
    Printer.CurrentX = col5x
    Printer.Print "PartNum"
    
    Printer.CurrentX = col6x
    Printer.Print "ReqQty"
    
    For i = 1 To matdecgrid.Rows - 1
        Printer.CurrentX = col1x
        Printer.Print matdecgrid.TextMatrix(i, 0); ")";
        Printer.CurrentX = col2x
        Printer.Print matdecgrid.TextMatrix(i, 1);
        Printer.CurrentX = col3x
        Printer.Print matdecgrid.TextMatrix(i, 2);
        Printer.CurrentX = col4x
        Printer.Print matdecgrid.TextMatrix(i, 3);
        Printer.CurrentX = col5x
        Printer.Print matdecgrid.TextMatrix(i, 4);
        Printer.CurrentX = col6x
        Printer.Print matdecgrid.TextMatrix(i, 5)

        lc = lc + 1
        'if the end of the page is reached then start over
        If lc >= 54 Then
            Printer.NewPage
            ppage = ppage + 1
            Printer.Print header
            lc = 0
        End If
    Next
    Printer.EndDoc
    
End Sub

Private Sub MATrefbut_Click(Index As Integer)
    matdecgrid.Rows = 0
    Call Form_Load
End Sub
