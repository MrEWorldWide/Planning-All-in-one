VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form JOBHOLDS 
   Caption         =   "Jobs on Hold"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton holdprintbut 
      Caption         =   "Print"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton holdcanbut 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid holdgrid 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "JOBHOLDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub holdcanbut_Click()
    Unload Me
End Sub

Private Sub holdprintbut_Click()
    
    Dim lc, i As Long
    Dim header As String
    Dim col1x, col2x, col3x, ppage As Integer
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
    Printer.CurrentX = 300
    Printer.Print "Number of jobs to enter: " & totaljobs.Caption
    Printer.Print
    Printer.Print
    lc = 0
    col1x = 300
    col2x = 900
    col3x = 3000
    Printer.CurrentY = 1550
    Printer.CurrentX = col2x
    Printer.Print "PartNum";
    Printer.CurrentX = col3x
    Printer.Print "Job Req"
    
    For i = 1 To bpogrid.Rows - 1
        Printer.CurrentX = col1x
        Printer.Print bpogrid.TextMatrix(i, 0); ")";
        Printer.CurrentX = col2x
        Printer.Print bpogrid.TextMatrix(i, 1);
        Printer.CurrentX = col3x
        Printer.Print bpogrid.TextMatrix(i, 2)
        lc = lc + 1
        If lc >= 54 Then
            col1x = 4500
            col2x = 5100
            col3x = 7200
            If lc = 54 Then
                Printer.CurrentY = 1550
                Printer.CurrentX = col2x
                Printer.Print "PartNum";
                Printer.CurrentX = col3x
                Printer.Print "Job Req"
            ElseIf lc = 108 Then
                col1x = 300
                col2x = 900
                col3x = 3000
                Printer.CurrentY = 1550
                Printer.CurrentX = col2x
                Printer.Print "PartNum";
                Printer.CurrentX = col3x
                Printer.Print "Job Req"
                Printer.NewPage
                ppage = ppage + 1
                Printer.Print header
                lc = 0
            End If
        End If
    Next
    Printer.EndDoc
    
End Sub

Private Sub Form_Load()

    DoEvents
    
    'counter variables
    Dim c, c1 As Long
    
    JOBHOLDS.MousePointer = 11
    
    SetupConnection
    SQLSelect = "SELECT Pub.JobHead.JobNum, Pub.JobHead.ReqDueDate, Pub.JobHead.PartNum, Pub.JobHead.ProdCode, "
    SQLSelect = SQLSelect & "Pub.JobHead.JobClosed FROM Pub.JobHead WHERE "
    SQLSelect = SQLSelect & "Pub.JobHead.ReqDueDate>='1/01/2020' AND Pub.JobHead.ReqDueDate<='12/31/2020' AND Pub.JobHead.JobClosed=0"
    RSRead.Open (SQLSelect), CNVantage
    
    'if no records notifu and exit sub
    If RSRead.EOF Then
        MsgBox " No Jobs are on hold.", vbExclamation, "No Records"
        RSRead.Close
        CNVantage.Close
        JOBHOLDS.MousePointer = 11
        Unload Me
        Exit Sub
    End If
   
    With holdgrid
    
        .WordWrap = True
        .Rows = 2
        .Cols = 5
        .FixedCols = 1
        .ColWidth(0) = 450
        .ColWidth(1) = 1050
        .ColWidth(2) = 1200
        .ColWidth(3) = 1000
        .ColWidth(4) = 700
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        
        .TextMatrix(0, 1) = "Job Num"
        .TextMatrix(0, 2) = "Part Num"
        .TextMatrix(0, 3) = "Req Date"
        .TextMatrix(0, 4) = "CUS"
        
        c = 1
        Do Until RSRead.EOF
            
            .TextMatrix(c, 0) = c
            .TextMatrix(c, 1) = RSRead.Fields(0)
            .TextMatrix(c, 2) = RSRead.Fields(2)
            .TextMatrix(c, 3) = RSRead.Fields(1)
            .TextMatrix(c, 4) = RSRead.Fields(3)
            .Rows = .Rows + 1
            c = c + 1
            RSRead.MoveNext
        Loop
        'remove extra row needed for fixed rows requirement
        .Rows = .Rows - 1
        .Col = 2
        .ColSel = 2
        .Sort = flexSortNumericAscending
    End With
    
    'end connection
    RSRead.Close
    CNVantage.Close
    JOBHOLDS.MousePointer = 0


End Sub
