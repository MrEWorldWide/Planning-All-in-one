VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form KANBAN 
   Caption         =   "KanBan Reschedules"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton KANcancelbut 
      Caption         =   "Exit"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton KANrefbut 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid KANGRID 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "KANBAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

    DoEvents
    Dim c As Long
    
    SetupConnection
    
    KANBAN.MousePointer = 11
    'our process uses the year 2017 as a kan ban identifier, select all
    'jobs that have the first operation completed of any assembly
    SQLSelect = "SELECT Pub.JobHead.JobNum, Pub.JobHead.ReqDueDate, Pub.JobOper.OpComplete, Pub.JobOper.WCCode, "
    SQLSelect = SQLSelect & "Pub.JobHead.PartNum, Pub.JobOper.OprSeq, Pub.JobHead.JobClosed, Pub.JobOper.AssemblySeq FROM Pub.JobOper "
    SQLSelect = SQLSelect & "INNER JOIN Pub.JobHead ON Pub.JobHead.JobNum=Pub.JobOper.JobNum "
    SQLSelect = SQLSelect & "WHERE Pub.JobOper.OpComplete=1 AND (Pub.JobHead.ReqDueDate >= '01/01/2017' "
    SQLSelect = SQLSelect & "AND Pub.JobHead.ReqDueDate <= '12/31/2017') AND Pub.JobHead.Jobclosed=0 "
    SQLSelect = SQLSelect & "AND Pub.JobOper.WCCode<>'RAW' AND Pub.JobOper.WCCode<>'FPACK' AND Pub.JobOper.OprSeq=10"
    
    RSRead.Open (SQLSelect), CNVantage
    
    'if no jobs are started, then exit sub
    If RSRead.EOF Then
        MsgBox "No jobs need to be rescheduled.", vbExclamation, "No Records"
        RSRead.Close
        CNVantage.Close
        KANBAN.MousePointer = 0
        Exit Sub
    End If
    
    With KANGRID
        .WordWrap = True
        .Rows = 2
        .Cols = 6
        .FixedCols = 1
        .ColWidth(0) = 350
        .ColWidth(1) = 800
        .ColWidth(2) = 1300
        .ColWidth(3) = 450
        .ColWidth(4) = 450
        .ColWidth(5) = 1000
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter

        .TextMatrix(0, 1) = "Job Num"
        .TextMatrix(0, 2) = "Part Num"
        .TextMatrix(0, 3) = "Asm"
        .TextMatrix(c, 4) = "Op"
        .TextMatrix(0, 5) = "Req Date"

        c = 1
        Do Until RSRead.EOF
            'this first condition checks for duplicate job numbers and skip the record
            If c <> 1 And .TextMatrix(c - 1, 1) = RSRead.Fields(0) Then
                RSRead.MoveNext
            Else
            .TextMatrix(c, 0) = c
            .TextMatrix(c, 1) = RSRead.Fields(0)
            .TextMatrix(c, 2) = RSRead.Fields(4)
            .TextMatrix(c, 3) = RSRead.Fields(7)
            .TextMatrix(c, 4) = RSRead.Fields(5)
            .TextMatrix(c, 5) = RSRead.Fields(1)
            RSRead.MoveNext
            c = c + 1
            .Rows = .Rows + 1
            End If
        Loop
        'remove blank row
        .Rows = .Rows - 1
    End With
    RSRead.Close
    CNVantage.Close
    KANBAN.MousePointer = 0
    
End Sub


Private Sub KANcancelbut_Click(Index As Integer)
    Unload KANBAN
End Sub


Private Sub KANrefbut_Click()
'refresh the form
    KANGRID.Rows = 0
   Call Form_Load

End Sub
