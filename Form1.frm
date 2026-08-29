VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MASTER 
   Caption         =   "Planning All in One"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   3000
      Value           =   2  'Grayed
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog prindia 
      Left            =   1200
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2025
      Value           =   2  'Grayed
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   1470
      Value           =   2  'Grayed
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox upchk 
      Caption         =   "Check1"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   19
      Top             =   600
      Value           =   2  'Grayed
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton jobholdbut 
      Caption         =   "Open"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox purdirchk 
      Caption         =   "Check1"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   16
      Top             =   5370
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox matdecchk 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   15
      Top             =   1020
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox incomchk 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   2415
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton incombutt 
      Caption         =   "Open"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton kanbanbut 
      Caption         =   "Open"
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton mcancel 
      Caption         =   "Exit"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton reqbpobut 
      Caption         =   "Open"
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton matdecbut 
      Caption         =   "Open"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton pudirbut 
      Caption         =   "Open"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton upcombut 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Jobs On Hold:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   17
      Top             =   2655
      Width           =   1575
   End
   Begin VB.Label excellab 
      Alignment       =   2  'Center
      Caption         =   "Save Items"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Incompletes:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label kanbanlab 
      Caption         =   "Schedule KanBan:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label reqbpolab 
      Caption         =   "Req BPO Jobs:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label matdeclab 
      Caption         =   "Partial Materials:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label purdirlab 
      Caption         =   "Purchase Direct:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   765
      Width           =   1815
   End
   Begin VB.Label upcomlab 
      Caption         =   "Upcoming Jobs:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   285
      Width           =   1695
   End
End
Attribute VB_Name = "MASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'this sub controls the saving feature. It isn't finished and may not be.

'    MsgBox "This will open all checked items. Are you sure you want to save the checked items to the database? This may take several minutes", vbOKCancel, "Warning"
'    If vbOK Then
'
'        Dim dbcon As New ADODB.Connection
'        Dim purdirrs, incomoprs, incommatrs, incomqtyrs, matdecrs As New ADODB.Recordset
'        Dim incomopstr, incommatstr, incomqtystr, purdirstr, matdecstr As String
'
'        Set dbcon = New Connection
'        dbcon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\allinonedb.accdb; Persist Security Info=False;")
'
'        Set purdirrs = New Recordset
'
'        If Me.purdirchk = 2 Then
'            PURDIR.Show
'            For c = 1 To purdirrs.RecordCount - 1
'                purdirstr = "INSERT INTO PURDIR(JobNum, AssemblySeq, PartNum, BuyIt, Direct, WarehouseCode) "
'                purdirstr = purdirstr & "VALUES('" & PURDIR.PURDIRgrid.TextMatrix(c, 1) & "', "
'                purdirstr = purdirstr & "" & PURDIR.PURDIRgrid.TextMatrix(c, 2) & ", "
'                purdirstr = purdirstr & "'" & PURDIR.PURDIRgrid.TextMatrix(c, 3) & "', "
'                purdirstr = purdirstr & "" & PURDIR.PURDIRgrid.TextMatrix(c, 4) & ", "
'               purdirstr = purdirstr & "" & PURDIR.PURDIRgrid.TextMatrix(c, 5) & ", "
'               purdirstr = purdirstr & "'" & PURDIR.PURDIRgrid.TextMatrix(c, 6) & "')"
'                Set purdirrs = dbcon.Execute(purdirstr)
'            Next
'        End If
'        If incomchk = vbChecked Then
'            INCOMPLETE.Show
'            For c = 1 To incomoprs.RecordCount - 1
'                incomopstr = "INSERT INTO INCOMOP(JobNum, AssemblySeq, OpSeq, WCCode, OpCode) "
'                incomopstr = incomopstr & "VALUES('" & INCOMPLETE.inopgrid.TextMatrix(c, 1) & "', "
'                incomopstr = incomopstr & "" & INCOMPLETE.inopgrid.TextMatrix(c, 2) & ", "
'                incomopstr = incomopstr & "" & INCOMPLETE.inopgrid.TextMatrix(c, 3) & ", "
'                incomopstr = incomopstr & "'" & INCOMPLETE.inopgrid.TextMatrix(c, 4) & "', "
'                incomopstr = incomopstr & "'" & INCOMPLETE.inopgrid.TextMatrix(c, 5) & "') "
'                Set incomoprs = dbcon.Execute(incomopstr)
'            Next
'
'            For c = 1 To incommatrs.RecordCount - 1
'                incommatstr = "INSERT INTO INCOMMAT(JobNum, AssemblySeq, MtlSeq, PartNum) "
'                incommatstr = incommatstr & "VALUES('" & INCOMPLETE.inmatgrid.TextMatrix(c, 1) & "', "
'                incommatstr = incommatstr & "" & INCOMPLETE.inmatgrid.TextMatrix(c, 2) & ", "
'                incommatstr = incommatstr & "" & INCOMPLETE.inmatgrid.TextMatrix(c, 3) & ", "
'                incommatstr = incommatstr & "'" & INCOMPLETE.inmatgrid.TextMatrix(c, 4) & "') "
'                Set incommatrs = dbcon.Execute(incommatstr)
'            Next
'
'            For c = 1 To incomqtyrs.RecordCount - 1
'                incomqtystr = "INSERT INTO INCOMQTY(JobNum, ProdQty, CompletedQty, Shipped, RevievedQty) "
'                incomqtystr = incomqtystr & "VALUES('" & INCOMPLETE.inshogrid.TextMatrix(c, 1) & "', "
'                incomqtystr = incomqtystr & "" & INCOMPLETE.inshogrid.TextMatrix(c, 2) & ", "
'                incomqtystr = incomqtystr & "" & INCOMPLETE.inshogrid.TextMatrix(c, 3) & ", "
'                incomqtystr = incomqtystr & "" & INCOMPLETE.inshogrid.TextMatrix(c, 4) & ", "
'                incomqtystr = incomqtystr & "" & INCOMPLETE.inshogrid.TextMatrix(c, 5) & ") "
'                Set incomqtyrs = dbcon.Execute(incomqtystr)
'            Next
'        End If
'        If matdecchk = vbChecked Then
'            For c = 2 To matdecrs.RecordCount - 1
'            MATDEC.Show
'                matdecstr = "INSERT INTO matdec(JobNum, AssemblySeq, MtlSeq, PartNum, RequiredQty) "
'                matdecstr = matdecstr & "VALUES('" & MATDEC.matdecgrid.TextMatrix(c, 1) & "', "
'                matdecstr = matdecstr & "" & MATDEC.matdecgrid.TextMatrix(c, 2) & ", "
'                matdecstr = matdecstr & "" & MATDEC.matdecgrid.TextMatrix(c, 3) & ", "
'                matdecstr = matdecstr & "'" & MATDEC.matdecgrid.TextMatrix(c, 4) & "', "
'                matdecstr = matdecstr & "" & MATDEC.matdecgrid.TextMatrix(c, 5) & ") "
'
'            Next
'         End If
'    Else
'    Exit Sub
'    End If
'
'    purdirrs.Close
'    incomoprs.Close
'    incomqtyrs.Close
'    incommatrs.Close
'    matdecrs.Close
'    dbcon.Close
End Sub


Private Sub incombut_Click(Index As Integer)
    INCOMPLETE.Top = MASTER.Top
    INCOMPLETE.Left = MASTER.Left + MASTER.Width
    INCOMPLETE.Show
    MsgBox "Feature not finished"
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'kill all forms upon master form exit
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub incombutt_Click(Index As Integer)
    INCOMPLETE.Top = MASTER.Top
    INCOMPLETE.Left = MASTER.Left + MASTER.Width
    INCOMPLETE.Show
End Sub

Private Sub jobholdbut_Click(Index As Integer)
    JOBHOLDS.Top = MASTER.Top
    JOBHOLDS.Left = MASTER.Left + MASTER.Width
    JOBHOLDS.Show
End Sub

Private Sub kanbanbut_Click(Index As Integer)
    KANBAN.Top = MASTER.Top
    KANBAN.Left = MASTER.Left + MASTER.Width
    KANBAN.Show
End Sub

Private Sub matdecbut_Click(Index As Integer)
    'MATDEC.Top = MASTER.Top
    'MATDEC.Left = MASTER.Left + MASTER.Width
    'MATDEC.Show
    MsgBox "This function has not been implemented", vbExclamation, "Form unavailable"
End Sub

Private Sub mcancel_Click(Index As Integer)
    'kill all forms upon exit
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub pudirbut_Click(Index As Integer)
    PURDIR.Top = MASTER.Top
    PURDIR.Left = MASTER.Left + MASTER.Width
    PURDIR.Show
End Sub

Private Sub reqbpobut_Click(Index As Integer)
    BPO.Top = MASTER.Top
    BPO.Left = MASTER.Left + MASTER.Width
    BPO.Show
End Sub

Private Sub upcombut_Click()
    UPCOMING.Top = MASTER.Top
    UPCOMING.Left = MASTER.Left + MASTER.Width
    UPCOMING.Show
    MsgBox "This feature is in the testing phase."
End Sub
