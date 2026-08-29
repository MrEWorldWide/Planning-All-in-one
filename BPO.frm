VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BPO 
   Caption         =   "BPO Jobs to Enter"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton listbut 
      Caption         =   "List"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton BPOrefbut 
      Caption         =   "Refresh"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton BPOcancelbut 
      Caption         =   "Exit"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton BPOprintbut 
      Caption         =   "Print"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   5520
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid bpogrid 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label totaljobs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label BPOlab 
      Caption         =   "Total required Jobs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "BPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bpocancel_Click()
    Unload BPO
End Sub
                                                                                                                                                                                                                                                               

Private Sub BPOPrintbut_Click()
    Dim lc, i As Long
    Dim header As String
    Dim col1x, col2x, col3x, col4x, ppage As Integer
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
    col3x = 2900
    col4x = 4000
    Printer.CurrentY = 1550
    Printer.CurrentX = col2x
    Printer.Print "PartNum";
    Printer.CurrentX = col3x
    Printer.Print "Job Req";
    Printer.CurrentX = col4x
    Printer.Print "Kan Ban?"
    
    For i = 1 To bpogrid.Rows - 1
        Printer.CurrentX = col1x
        Printer.Print bpogrid.TextMatrix(i, 0); ")";
        Printer.CurrentX = col2x
        Printer.Print bpogrid.TextMatrix(i, 1);
        Printer.CurrentX = col3x
        Printer.Print bpogrid.TextMatrix(i, 2);
        Printer.CurrentX = col4x
        Printer.Print bpogrid.TextMatrix(i, 3)
        lc = lc + 1
        If lc >= 48 Then
            col1x = 5500
            col2x = 6100
            col3x = 8200
            col4x = 9300
            If lc = 48 Then
                Printer.CurrentY = 1550
                Printer.CurrentX = col2x
                Printer.Print "PartNum";
                Printer.CurrentX = col3x
                Printer.Print "Job Req";
                Printer.CurrentX = col4x
                Printer.Print "Kan Ban?"
            ElseIf lc = 96 Then
                col1x = 300
                col2x = 900
                col3x = 2900
                colx4 = 4000
                Printer.CurrentY = 1550
                Printer.CurrentX = col2x
                Printer.Print "PartNum";
                Printer.CurrentX = col3x
                Printer.Print "Job Req";
                Printer.CurrentX = col4x
                Printer.Print "Kan Ban?"
                Printer.NewPage
                ppage = ppage + 1
                Printer.Print header
                lc = 0
            End If
        End If
    Next
    Printer.EndDoc
    
End Sub

Private Sub BPOcancelbut_Click(Index As Integer)
    Unload BPO
End Sub

Private Sub BPOrefbut_Click(Index As Integer)
     bpogrid.Rows = 0
     Call Form_Load
End Sub
Private Sub Form_Load()
    
    DoEvents
    'counter vars
    Dim c, c1, jo, lc As Long
    'this var holds the row data that swaps during sort.
    Dim swapvar1, swapvar2 As String

    On Error GoTo errhand
    SetupConnection
    
    'select all parts in the database that have a warranty code of 'BPO'
    ' then take all the open jobs for schweitzers BPO parts and ignore RMA IPR and Proto jobs
    SQLSelect = "SELECT Pub.Part.PartNum, Pub.Part.ProdCode, Pub.Part.WarrantyCode, Pub.JobHead.JobNum, "
    SQLSelect = SQLSelect + "Pub.JobHead.JobClosed, Pub.JobHead.ReqDueDate FROM Pub.Part LEFT JOIN Pub.JobHead ON Pub.JobHead.PartNum=Pub.Part.PartNum "
    SQLSelect = SQLSelect + "WHERE Pub.Part.WarrantyCode='BPO' AND Pub.Part.ProdCode='SEL' "
    SQLSelect = SQLSelect + "AND (Pub.JobHead.JobNum NOT LIKE 'I%' AND Pub.JobHead.JobNum NOT LIKE 'P%' AND "
    SQLSelect = SQLSelect + "Pub.JobHead.JobNum NOT LIKE 'R%') ORDER BY Pub.Part.PartNum, Pub.Jobhead.JobClosed"
    
    BPO.MousePointer = 11
    RSRead.Open (SQLSelect), CNVantage
    
    'exit sub if no records exist
    If RSRead.EOF Then
        MsgBox " No BPO jobs are required", vbExclamation, "No Records"
        Unload Me
        RSRead.Close
        CNVantage.Close
        BPO.MousePointer = 0
        Exit Sub
    End If
    
    With bpogrid
    
        .Visible = False
        .WordWrap = True
        .Rows = 2
        .Cols = 4
        .FixedCols = 1
        .ColWidth(0) = 450
        .ColWidth(1) = 1450
        .ColWidth(2) = 600
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        
        .TextMatrix(0, 1) = "Part Number"
        .TextMatrix(0, 2) = " Jobs"
        .TextMatrix(0, 3) = "KanBan?"
        
        c = 1
        jo = 0
        lc = 0

        Do Until RSRead.EOF
            'if the name is on the list, dont create a new row
            'just add to the previous records' job requirement to the last record's
            If RSRead.Fields(0) = .TextMatrix(c - 1, 1) Then
            'count the number of open jobs, lc is for the total count
                If RSRead.Fields(4) = False Then
                    jo = jo + 1
                    lc = lc + 1
                End If
                'job requirements  equals the previous records' count plus the current
                .TextMatrix(c - 1, 2) = Val(.TextMatrix(c - 1, 2)) + jo
                
                jo = 0
            Else
            'if the part number doesn't exist previously then create a new row
                .Rows = .Rows + 1
                .RowHeight(c + 1) = 250
                .TextMatrix(c, 1) = RSRead.Fields(0)
                
                If RSRead.Fields(4) = False Then
                    jo = jo + 1
                    lc = lc + 1
                End If
                .TextMatrix(c, 2) = jo
                c = c + 1
                jo = 0
            End If
            RSRead.MoveNext
        Loop
        'make the job requirement: 3 - jo
        For c = 1 To .Rows - 1
            .TextMatrix(c, 2) = 3 - Val(.TextMatrix(c, 2))
        Next
        'remove extra row needed for fixed rows requirement
        .Rows = .Rows - 1
        'number desc sort on column #2
        .Col = 2
        .ColSel = 2
        .Sort = flexSortNumericDescending
        
        For c = 1 To .Rows - 1
        
            If c = .Rows - 1 Then Exit For
            
            'if the two compared part numbers' jobs required are equal and sort them
            If .TextMatrix(c, 1) > .TextMatrix(c + 1, 1) And Val(.TextMatrix(c, 2)) = Val(.TextMatrix(c + 1, 2)) Then
            'swap the partnumbers to ascending order
                swapvar1 = .TextMatrix(c, 1)
                swapvar2 = .TextMatrix(c, 3)
                
                .TextMatrix(c, 1) = .TextMatrix(c + 1, 1)
                .TextMatrix(c + 1, 1) = swapvar1
                
                'be sure to bring kan ban reference with the cell swap
                .TextMatrix(c, 3) = .TextMatrix(c + 1, 3)
                .TextMatrix(c + 1, 3) = swapvar2
                
                c = 1
                If c = 10 Then Exit For
            End If
        .TextMatrix(c, 0) = c
        Next
        'remove negatives from job requirement.negatives are made from 3 jobs minus
        'current number of jobs. 5 current jobs would create -2
        While .TextMatrix((.Rows - 1), 2) <= 0
                .Rows = .Rows - 1
        Wend
    
    .Visible = True
    'total job counter label
    For c = 1 To .Rows - 1
        c1 = c1 + Val(.TextMatrix(c, 2))
    Next
    
    totaljobs.Caption = c1

    'check the local db for kanban parts
    LocalConnection
    For c = 0 To .Rows - 2
        Set RSRead = New ADODB.Recordset
        SQLSelect = "SELECT PartNum FROM Parts WHERE PartNum='" & .TextMatrix(c, 1) & "'"
        RSRead.Open (SQLSelect), CNACCESS
        Do Until RSRead.EOF
            If .TextMatrix(c, 1) = RSRead!PartNum Then
                    .TextMatrix(c, 3) = "KB"
            End If
            RSRead.MoveNext
        Loop


    Next
    End With
    
    'end connection
    RSRead.Close
    CNVantage.Close
    BPO.MousePointer = 0
    
errhand:
    
    
End Sub

Private Sub listbut_Click()
    KBLIST.Top = BPO.Top
    KBLIST.Show
    KBLIST.Left = BPO.Left
End Sub
