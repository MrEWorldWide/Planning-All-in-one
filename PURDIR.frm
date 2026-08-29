VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PURDIR 
   Caption         =   "Purchase Direct Items"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog printdia 
      Left            =   1320
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton PURprintbut 
      Caption         =   "Print"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   3
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton PURcancelbut 
      Caption         =   "Exit"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton PURrefbut 
      Caption         =   "Refresh"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   6720
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid PURDIRgrid 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   11456
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "PURDIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
       
    DoEvents
    
    SetupConnection

    SQLSelect = SQLSelect & "SELECT Pub.JobMtl.AssemblySeq, Pub.JobMtl.JobNum, Pub.JobMtl.PartNum, Pub.JobMtl.BuyIt, "
    SQLSelect = SQLSelect & "Pub.JobMtl.Direct, Pub.JobMtl.WarehouseCode, Pub.JobHead.JobClosed, Pub.JobMtl.Description "
    SQLSelect = SQLSelect & "FROM Pub.JobMtl INNER JOIN Pub.JobHead ON Pub.JobHead.JobNum=Pub.JobMtl.JobNum WHERE "
    SQLSelect = SQLSelect & "Pub.JobHead.JobClosed=0 AND (Pub.JobMtl.BuyIt=1 OR Pub.JobMtl.Direct=1 OR "
    SQLSelect = SQLSelect & "Pub.JobMtl.WarehouseCode='AFA') AND Pub.JobHead.JobNum>='110000' AND Pub.JobMtl.JobNum NOT LIKE 'MARK%' AND "
    SQLSelect = SQLSelect & "Pub.JobMtl.PartNum NOT LIKE 'PACK%' AND Pub.JobMtl.PartNum NOT LIKE 'SHP%' "
    SQLSelect = SQLSelect & "AND Pub.JobMtl.JobNum NOT LIKE 'I%' AND Pub.JobMtl.JobNum NOT LIKE 'R%' AND Pub.JobHead.ReqDueDate<='12/31/2014'"
    
    PURDIR.MousePointer = 11
    RSRead.Open (SQLSelect), CNVantage
    
    'if no records notify user then exit sub
    If RSRead.EOF = True Then
        MsgBox "No Purchase Direct items available."
        PURDIR.MousePointer = 0
        RSRead.Close
        CNVantage.Close
        Unload Me
        Exit Sub
    End If
   
    With PURDIRgrid
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
        .ColWidth(6) = 550
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .TextMatrix(0, 1) = "Job Num"
        .TextMatrix(0, 2) = "Asm"
        .TextMatrix(0, 3) = "Part Num"
        .TextMatrix(0, 4) = "Pur Direct"
        .TextMatrix(0, 5) = "Direct"
        .TextMatrix(0, 6) = "Whse"
        
        'start at the first matrix cell and enter all records
        c = 1
        Do Until RSRead.EOF
            .TextMatrix(c, 1) = RSRead.Fields(1)
            .TextMatrix(c, 2) = RSRead.Fields(0)
            .TextMatrix(c, 3) = RSRead.Fields(2)
            'convert true and false to yes and no for the user
            If RSRead.Fields(3) = True Then
                .TextMatrix(c, 4) = "Yes"
            ElseIf RSRead.Fields(3) = False Then
                 .TextMatrix(c, 4) = "No"
            End If
            If RSRead.Fields(4) = True Then
                .TextMatrix(c, 5) = "Yes"
            ElseIf RSRead.Fields(4) = False Then
                 .TextMatrix(c, 5) = "No"
            End If
            .TextMatrix(c, 6) = RSRead.Fields(5)
            RSRead.MoveNext
            .Rows = .Rows + 1
            c = c + 1
            
        Loop
        'ID the rows
        For c = 1 To .Rows - 1
            If c = .Rows Then Exit For
            .TextMatrix(c, 0) = c
        Next
        .Rows = .Rows - 1
    End With
    'close records and connection
    RSRead.Close
    CNVantage.Close
    PURDIR.MousePointer = 0
End Sub

Private Sub PURcancelbut_Click(Index As Integer)
    PURDIRgrid.Rows = 0
    Unload PURDIR
End Sub

Private Sub PURprintbut_Click(Index As Integer)
    With printdia
        .ShowPrinter
    End With
    'If vbCancel Then
    '    Exit Sub
    'End If
    'lc=line counter and i=interval counter
    Dim lc, i As Long
    Dim header, title As String
    'cloumn locations and page counter
    Dim col1x, col2x, col3x, col4x, col5x, col6x, col7x, ppage As Integer
    
    If PURDIRgrid.Cols = 0 Then
        MsgBox "No report to print"
        Exit Sub
    End If
    
    title = "Purchase Direct Items"
    
    ppage = 1
    header = "  Page:" & ppage & "  " & Now() & "      Origin: " & Environ("USERNAME") & vbTab & vbTab & title & vbCrLf
    header = header & "-----------------------------------------------------------------------"
    
    Printer.FontName = "Verdana"
    Printer.FontSize = 11
    Printer.Print
    Printer.Print header
    Printer.Print
    Printer.Print
    
    lc = 0
    col1x = 300
    col2x = 900
    col3x = 2300
    col4x = 2900
    col5x = 5300
    col6x = 6200
    col7x = 7000
    Printer.CurrentY = 1550
    
    Printer.CurrentX = col2x
    Printer.Print "JobNum";
    
    Printer.CurrentX = col3x
    Printer.Print "ASM";
    
    Printer.CurrentX = col4x
    Printer.Print "PartNum";
    
    Printer.CurrentX = col5x
    Printer.Print "PurDir";
    
    Printer.CurrentX = col6x
    Printer.Print "Direct";
    
    Printer.CurrentX = col7x
    Printer.Print "Whse"
    
    For i = 1 To PURDIRgrid.Rows - 1
    
        Printer.CurrentX = col1x
        Printer.Print PURDIRgrid.TextMatrix(i, 0); ")";
        Printer.CurrentX = col2x
        Printer.Print PURDIRgrid.TextMatrix(i, 1);
        Printer.CurrentX = col3x
        Printer.Print PURDIRgrid.TextMatrix(i, 2);
        Printer.CurrentX = col4x
        Printer.Print PURDIRgrid.TextMatrix(i, 3);
        Printer.CurrentX = col5x
        Printer.Print PURDIRgrid.TextMatrix(i, 4);
        Printer.CurrentX = col6x
        Printer.Print PURDIRgrid.TextMatrix(i, 5);
        Printer.CurrentX = col7x
        Printer.Print PURDIRgrid.TextMatrix(i, 6)
        
        lc = lc + 1
        If lc >= 54 Then
            Printer.NewPage
            ppage = ppage + 1
            Printer.Print header
            lc = 0
        End If
    Next
    Printer.EndDoc
    
End Sub

Private Sub PURrefbut_Click(Index As Integer)
    SetupConnection

    SQLSelect = SQLSelect & "SELECT Pub.JobMtl.AssemblySeq, Pub.JobMtl.JobNum, Pub.JobMtl.PartNum, Pub.JobMtl.BuyIt, "
    SQLSelect = SQLSelect & "Pub.JobMtl.Direct, Pub.JobMtl.WarehouseCode, Pub.JobHead.JobClosed, Pub.JobMtl.Description "
    SQLSelect = SQLSelect & "FROM Pub.JobMtl INNER JOIN Pub.JobHead ON Pub.JobHead.JobNum=Pub.JobMtl.JobNum WHERE "
    SQLSelect = SQLSelect & "Pub.JobHead.JobClosed=0 AND (Pub.JobMtl.BuyIt=1 OR Pub.JobMtl.Direct=1 OR "
    SQLSelect = SQLSelect & "Pub.JobMtl.WarehouseCode='AFA') AND Pub.JobHead.JobNum>='110000' AND Pub.JobMtl.JobNum NOT LIKE 'MARK%' AND "
    SQLSelect = SQLSelect & "Pub.JobMtl.PartNum NOT LIKE 'PACK%' AND Pub.JobMtl.PartNum NOT LIKE 'SHP%' "
    SQLSelect = SQLSelect & "AND Pub.JobMtl.JobNum NOT LIKE 'I%' AND Pub.JobMtl.JobNum NOT LIKE 'R%' AND Pub.JobHead.ReqDueDate<='12/31/2014'"
    
    PURDIR.MousePointer = 11
    RSRead.Open (SQLSelect), CNVantage
    
    'if no records notify user then exit sub
    If RSRead.EOF = True Then
        MsgBox "No Purchase Direct items available."
        PURDIR.MousePointer = 0
        RSRead.Close
        CNVantage.Close
        Unload Me
        Exit Sub
    End If
   
    With PURDIRgrid
    
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
        .ColWidth(6) = 550
        
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        
        .TextMatrix(0, 1) = "Job Num"
        .TextMatrix(0, 2) = "Asm"
        .TextMatrix(0, 3) = "Part Num"
        .TextMatrix(0, 4) = "Pur Direct"
        .TextMatrix(0, 5) = "Direct"
        .TextMatrix(0, 6) = "Whse"
        
        
        c = 1
        Do Until RSRead.EOF
            .TextMatrix(c, 1) = RSRead.Fields(1)
            .TextMatrix(c, 2) = RSRead.Fields(0)
            .TextMatrix(c, 3) = RSRead.Fields(2)
            'convert true and false to yes and no for the user
            If RSRead.Fields(3) = True Then
                .TextMatrix(c, 4) = "Yes"
            ElseIf RSRead.Fields(3) = False Then
                 .TextMatrix(c, 4) = "No"
            End If
            If RSRead.Fields(4) = True Then
                .TextMatrix(c, 5) = "Yes"
            ElseIf RSRead.Fields(4) = False Then
                 .TextMatrix(c, 5) = "No"
            End If
            .TextMatrix(c, 6) = RSRead.Fields(5)
            RSRead.MoveNext
            .Rows = .Rows + 1
            c = c + 1
            
        Loop
        
        'number the rows
        For c = 1 To .Rows - 1
            If c = .Rows Then Exit For
            .TextMatrix(c, 0) = c
        Next
        .Rows = .Rows - 1
    End With
    'close records and connection
    RSRead.Close
    CNVantage.Close
    PURDIR.MousePointer = 0
End Sub
