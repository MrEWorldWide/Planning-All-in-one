VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form INCOMPLETE 
   Caption         =   "Incomplete Materials Jobs, and Operations"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid inshogrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid inmatgrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton incomprintbut 
      Caption         =   "Print"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   7440
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid inopgrid 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      WordWrap        =   -1  'True
   End
   Begin VB.CommandButton incomcopybut 
      Caption         =   "Copy"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton incomcanbut 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton incomrefbut 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity errors:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4995
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Incomplete Materials:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2595
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Incomplete Operations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   45
      Width           =   2415
   End
End
Attribute VB_Name = "INCOMPLETE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub incomcanbut_Click()

    Unload INCOMPLETE
    
End Sub


Private Sub Form_Load()
    DoEvents
    Dim c As Long
    
    SetupConnection
    
'**************************'
'Operation grid'

    Set RSRead = New ADODB.Recordset
    SQLSelect = "SELECT  Pub.JobHead.Candidate, Pub.JobOper.JobNum, Pub.JobOper.OprSeq, Pub.JobOper.AssemblySeq, "
    SQLSelect = SQLSelect + "Pub.JobOper.OpComplete, Pub.JobOper.WCCode, Pub.JobOper.OpCode, Pub.JobHead.ReqDueDate, Pub.JobHead.JobClosed FROM Pub.JobHead INNER JOIN Pub.JobOper "
    SQLSelect = SQLSelect + "ON Pub.JobHead.JobNum=Pub.JobOper.JobNum WHERE Pub.JobOper.OpComplete=0 AND Pub.JobHead.Candidate=1 AND Pub.JobHead.ReqDueDate>='01/01/2012' AND Pub.JobHead.JobClosed=0"
    
    INCOMPLETE.MousePointer = 11
    RSRead.Open (SQLSelect), CNVantage
   
   'if no records then notify
    If RSRead.EOF Then
        MsgBox "There are no incomplete operations.", vbExclamation, "No Records"
    Else
        With inopgrid
            .Rows = 2
            .Cols = 6
            .FixedCols = 1
            .ColWidth(0) = 350
            .ColWidth(1) = 1000
            .ColWidth(2) = 450
            .ColWidth(3) = 450
            .ColWidth(4) = 950
            .ColWidth(5) = 786
            
            .FixedRows = 1
            .RowHeight(0) = 500
            .ColAlignment(0) = flexAlignRightCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignRightCenter
            .ColAlignment(5) = flexAlignRightCenter
    
            
            .TextMatrix(0, 1) = "Job Num"
            .TextMatrix(0, 2) = "ASM"
            .TextMatrix(0, 3) = "OP Seq"
            .TextMatrix(0, 4) = "WkCenter"
            .TextMatrix(0, 5) = "Op"
            
            c = 1
            Do Until RSRead.EOF Or .Rows = 1100
                .TextMatrix(c, 1) = RSRead.Fields(1)
                .TextMatrix(c, 2) = RSRead.Fields(3)
                .TextMatrix(c, 3) = RSRead.Fields(2)
                .TextMatrix(c, 4) = RSRead.Fields(5)
                .TextMatrix(c, 5) = RSRead.Fields(6)
                .Rows = .Rows + 1
                c = c + 1
                RSRead.MoveNext
           Loop
           .Rows = .Rows - 1
            .Col = 1
            .ColSel = 1
            .Sort = flexSortStringAscending
            'add ID numbers to rows
            For c = 1 To .Rows - 1
                If c = .Rows Then Exit For
                .TextMatrix(c, 0) = c
            Next
            
        End With
    End If

    

'**************************'
'Material grid'

    Set RSRead = New ADODB.Recordset
    SQLSelect = "SELECT Pub.JobMtl.AssemblySeq, Pub.JobMtl.IssuedComplete, Pub.JobMtl.IssuedQty, Pub.Jobhead.Candidate, "
    SQLSelect = SQLSelect & "Pub.JobMtl.MtlSeq, Pub.JobMtl.PartNum, Pub.JobMtl.JobNum, Pub.JobHead.JobClosed, Pub.JobHead.ReqDueDate FROM Pub.JobMtl "
    SQLSelect = SQLSelect & "INNER JOIN Pub.JobHead ON Pub.JobMtl.JobNum=Pub.JobHead.JobNum "
    SQLSelect = SQLSelect & "WHERE Pub.JobMtl.IssuedComplete=0 AND Pub.JobHead.JobClosed=0 AND Pub.JobHead.Candidate=1 AND Pub.JobHead.ReqDueDate>='01/01/2012' "
    SQLSelect = SQLSelect & "AND Pub.JobMtl.PartNum NOT LIKE 'SHP%' AND Pub.JobMtl.PartNum NOT LIKE 'PACK%'"
    SQLSelect = SQLSelect & "AND Pub.JobMtl.PartNum NOT LIKE 'PL%' AND Pub.JobMtl.PartNum NOT LIKE 'PC%' AND Pub.JobMtl.PartNum NOT LIKE 'A-LABEL%'"
    SQLSelect = SQLSelect & "ORDER BY Pub.JobHead.JobNum "
    
    RSRead.Open (SQLSelect), CNVantage
   'if no records then notify
    If RSRead.EOF Then
        MsgBox "There are no incomplete materials.", vbExclamation, "No Records"
    Else
        With inmatgrid
            .WordWrap = True
            .Rows = 2
            .Cols = 5
            .FixedCols = 1
            .ColWidth(0) = 350
            .ColWidth(1) = 1000
            .ColWidth(2) = 450
            .ColWidth(3) = 450
            .ColWidth(4) = 1730
            
            .FixedRows = 1
            .RowHeight(0) = 500
            .ColAlignment(0) = flexAlignRightCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignLeftCenter
            
            .TextMatrix(0, 1) = "Job Num"
            .TextMatrix(0, 2) = "ASM"
            .TextMatrix(0, 3) = "Mtl Seq"
            .TextMatrix(0, 4) = "Part Num"
            
            
            c = 1
            Do Until RSRead.EOF
                .TextMatrix(c, 1) = RSRead.Fields(6)
                .TextMatrix(c, 2) = RSRead.Fields(0)
                .TextMatrix(c, 3) = RSRead.Fields(4)
                .TextMatrix(c, 4) = RSRead.Fields(5)
                RSRead.MoveNext
                c = c + 1
                .Rows = .Rows + 1
           Loop
           'row IDs need to be inserted after the sort
           .Rows = .Rows - 1
            .Col = 1
            .ColSel = 1
            .Sort = flexSortNumericAscending
            For c = 1 To .Rows
                If c = .Rows Then Exit For
                .TextMatrix(c, 0) = c
            Next
            
        End With
    End If


'**************************'
'short Quantities grid'
    SetupConnection
    Set RSRead = New Recordset
    SQLSelect = "SELECT Pub.JobHead.JobComplete, Pub.JobHead.PartNum, Pub.JobHead.StockQty, Pub.JobHead.ShippedQty, "
    SQLSelect = SQLSelect & "Pub.JobHead.JobNum, Pub.JobHead.ReceivedQty, Pub.Part.WarrantyCode, Pub.JobHead.Candidate, "
    SQLSelect = SQLSelect & "Pub.JobHead.ReqDueDate FROM Pub.JobHead INNER JOIN Pub.Part ON "
    SQLSelect = SQLSelect & "Pub.JobHead.PartNum=Pub.Part.PartNum "
    SQLSelect = SQLSelect & "WHERE Pub.JobHead.Candidate=1 AND Pub.Part.WarrantyCode<>'BPO' AND Pub.JobHead.ReqDueDate>='01/01/2012'"
    SQLSelect = SQLSelect & "ORDER BY Pub.JobHead.JobNum"
    Debug.Print SQLSelect
    RSRead.Open (SQLSelect), CNVantage
   'if no records then notify
    If RSRead.EOF Then
        MsgBox "There are no shipping, inventory, short job errors.", vbExclamation, "No Records"
    Else
         With inshogrid
            .Rows = 2
            .Cols = 6
            .FixedCols = 1
            .ColWidth(0) = 400
            .ColWidth(1) = 1000
            .ColWidth(2) = 500
            .ColWidth(3) = 500
            .ColWidth(4) = 500
            .ColWidth(5) = 500
    
            
            .FixedRows = 1
            .RowHeight(0) = 500
            .ColAlignment(0) = flexAlignRightCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignRightCenter
    
            .WordWrap = True
            .TextMatrix(0, 1) = "Job Num"
            .TextMatrix(0, 2) = "Prod Qty"
            .TextMatrix(0, 3) = "Cmp Qty"
            .TextMatrix(0, 4) = "Ship Qty"
            .TextMatrix(0, 5) = "Rec Qty"
            
            c = 1

            Do Until RSRead.EOF
                If RSRead.Fields(2) - RSRead.Fields(5) - RSRead(3) <> 0 Or RSRead.Fields(6) = "BPO" Then
                    RSRead.MoveNext
                Else
                .TextMatrix(c, 1) = RSRead.Fields(4)
                .TextMatrix(c, 2) = RSRead.Fields(2)
                .TextMatrix(c, 3) = RSRead.Fields(8)
                .TextMatrix(c, 4) = RSRead.Fields(3)
                .TextMatrix(c, 5) = RSRead.Fields(6)
                .Rows = .Rows + 1
                RSRead.MoveNext
                c = c + 1
                End If
           Loop
           RSRead.Close
           'add final completed qty to matrix
            For c = 1 To .Rows - 1
                SQLSelect = "SELECT  Pub.JobOper.QTYCompleted FROM Pub.JobOper "
                SQLSelect = SQLSelect & "WHERE Pub.JobOper.JobNum='" & .TextMatrix(c, 1) & "' AND Pub.JobOper.AssemblySeq=0"
                RSRead.Open (SQLSelect), CNVantage
                Do Until RSRead.EOF
                    .TextMatrix(c, 3) = RSRead.Fields(0)
                    RSRead.MoveNext
                Loop
                c = c + 1
                RSRead.Close
            Next
           
           
           
           
           
           
           
           'ID the rows
           .Rows = .Rows - 1
            .Col = 1
            .ColSel = 1
            .Sort = flexSortStringAscending
            
            For c = 1 To .Rows - 1
                If c = .Rows Then Exit For
                .TextMatrix(c, 0) = c
            Next
        End With
   End If
'end connection
    RSRead.Close
    CNVantage.Close
    INCOMPLETE.MousePointer = 0
    
End Sub

Private Sub incomcopybut_Click()
    Clipboard.Clear
    Dim copystr As String
    Dim c, c1, c2 As Long
    
    copystr = vbCr + vbCr
    'copy operation grid
    copystr = "Incomplete Items     " & Date & vbCr
    'copy material grid
    With inmatgrid
        If .TextMatrix(1, 1) <> "" Then
            copystr = copystr & vbCr & "Operations:" & vbCrLf
            copystr = copystr + .TextMatrix(0, 1) + "   " + .TextMatrix(0, 2)
            copystr = copystr + "  " + .TextMatrix(0, 3) + "    " + .TextMatrix(0, 4)
            copystr = copystr + vbCr
            For c = 1 To .Rows - 1
                copystr = copystr + "  " + .TextMatrix(c, 1)
                For c2 = 1 To 15 - Len(.TextMatrix(c, 1))
                    copystr = copystr + " "
                Next
                copystr = copystr + .TextMatrix(c, 2)
                For c2 = 1 To 10 - Len(.TextMatrix(c, 2))
                    copystr = copystr + " "
                Next
                copystr = copystr + .TextMatrix(c, 3)
                For c2 = 1 To 10 - Len(.TextMatrix(c, 3))
                    copystr = copystr + " "
                Next
                copystr = copystr + .TextMatrix(c, 4)
                For c2 = 1 To 20 - Len(.TextMatrix(c, 4))
                    copystr = copystr + " "
                Next
                copystr = copystr & vbCr
             Next
        End If
        
    End With
    
    'copy material grid
    With inmatgrid
        If .TextMatrix(1, 1) <> "" Then
            copystr = copystr & vbCr & "Materials:" & vbCrLf
            copystr = copystr + .TextMatrix(0, 1) + "   " + .TextMatrix(0, 2)
            copystr = copystr + "  " + .TextMatrix(0, 3) + "    " + .TextMatrix(0, 4)
            copystr = copystr + vbCr
            For c = 1 To .Rows - 1
                copystr = copystr + "  " + .TextMatrix(c, 1)
                For c2 = 1 To 15 - Len(.TextMatrix(c, 1))
                    copystr = copystr + " "
                Next
                copystr = copystr + .TextMatrix(c, 2)
                For c2 = 1 To 10 - Len(.TextMatrix(c, 2))
                    copystr = copystr + " "
                Next
                copystr = copystr + .TextMatrix(c, 3)
                For c2 = 1 To 10 - Len(.TextMatrix(c, 3))
                    copystr = copystr + " "
                Next
                copystr = copystr + .TextMatrix(c, 4)
                For c2 = 1 To 20 - Len(.TextMatrix(c, 4))
                    copystr = copystr + " "
                Next
                copystr = copystr + vbCr
             Next
        End If
    End With
    
    
    
    'copy short job grid
    With inshogrid
    
        If .TextMatrix(1, 1) <> "" Then
        copystr = copystr & vbCr & "Job Quantities:" & vbCr
        copystr = copystr + vbCr + .TextMatrix(0, 1) + "   " + .TextMatrix(0, 2)
        copystr = copystr + "  " + .TextMatrix(0, 3) + "    " + .TextMatrix(0, 4)
        copystr = copystr + .TextMatrix(0, 5) + vbCr
        For c = 1 To .Rows - 1
            copystr = copystr + "  " + .TextMatrix(c, 1)
            For c2 = 1 To 15 - Len(.TextMatrix(c, 1))
                copystr = copystr + " "
            Next
                copystr = copystr + .TextMatrix(c, 2)
            For c2 = 1 To 15 - Len(.TextMatrix(c, 2))
                copystr = copystr + " "
            Next
                copystr = copystr + .TextMatrix(c, 3)
            For c2 = 1 To 15 - Len(.TextMatrix(c, 3))
                copystr = copystr + " "
            Next
                copystr = copystr + .TextMatrix(c, 4)
            For c2 = 1 To 15 - Len(.TextMatrix(c, 4))
                copystr = copystr + " "
            Next
                copystr = copystr + .TextMatrix(c, 5)
                copystr = copystr + vbCr
            Next
        End If
    End With
         
    Clipboard.SetText copystr

End Sub

Private Sub incomprintbut_Click()
    
    'line counter and
    Dim lc, i As Long
    Dim title, header As String
    'col vars are the alignments to the page
    Dim col1x, col2x, col3x, col4x, col5x, col6x, ppage As Integer
    
    'Page Name
    title = "Operation, Material, and Quantity issues"
    
    'start page counter at 1
    ppage = 1
    
    header = vbCrLf & "  Page:" & ppage & "  " & Now() & "      Origin: " & Environ("USERNAME") & vbTab & vbTab & title
    header = header & "-----------------------------------------------------------------------"
    
    
    Printer.FontName = "Verdana"
    Printer.FontSize = 11
    
    Printer.Print
    Printer.Print header
    Printer.Print
    Printer.Print
    'start the line counter
    lc = 0
    
'*****************************
'Print operation discrepancies
'*****************************

    If inopgrid.TextMatrix(0, 1) <> "" Then
        col1x = 300
        col2x = 900
        col3x = 3000
        col4x = 4050
        col5x = 5000
        col6x = 6000
        
        Printer.FontSize = 14
        Printer.CurrentX = col1x
        Printer.Print vbCrLf
        Printer.Print "   Incomplete Operations:"
        Printer.Print "   -------------------------"
        Printer.FontSize = 11
        
        Printer.CurrentX = col2x
        Printer.Print "JobNum";
        
        Printer.CurrentX = col3x
        Printer.Print "ASM";
        
        Printer.CurrentX = col4x
        Printer.Print "OPSeq";
        
        Printer.CurrentX = col5x
        Printer.Print "WkCtr";
        
        Printer.CurrentX = col6x
        Printer.Print "OP"
        
        For i = 1 To inopgrid.Rows - 1
            
            Printer.CurrentX = col1x
            Printer.Print inopgrid.TextMatrix(i, 0); ")";
            Printer.CurrentX = col2x
            Printer.Print inopgrid.TextMatrix(i, 1);
            Printer.CurrentX = col3x
            Printer.Print inopgrid.TextMatrix(i, 2);
            Printer.CurrentX = col4x
            Printer.Print inopgrid.TextMatrix(i, 3);
            Printer.CurrentX = col5x
            Printer.Print inopgrid.TextMatrix(i, 4);
            Printer.CurrentX = col6x
            Printer.Print inopgrid.TextMatrix(i, 5)
            
            lc = lc + 1
            If lc >= 54 Then
                Printer.NewPage
                ppage = ppage + 1
                Printer.Print header
                lc = 0
            End If
        Next
        Printer.Print
    End If
    
'****************************
'Print material discrepancies
'****************************
    
    If inmatgrid.TextMatrix(0, 1) <> "" Then
        col1x = 300
        col2x = 900
        col3x = 3000
        col4x = 4000
        col5x = 500

        Printer.CurrentX = col1x
        Printer.FontSize = 14
        Printer.Print "   Incomplete Materials:"
        Printer.Print "   ---------------------"
        Printer.FontSize = 11
        
        Printer.CurrentX = col2x
        Printer.Print "JobNum";
        
        Printer.CurrentX = col3x
        Printer.Print "ASM";
        
        Printer.CurrentX = col4x
        Printer.Print "MtlSeq";
        
        Printer.CurrentX = col5x
        Printer.Print "Partnum";
        
        For i = 1 To inmatgrid.Rows - 1
            
            Printer.CurrentX = col1x
            Printer.Print inmatgrid.TextMatrix(i, 0); ")";
            Printer.CurrentX = col2x
            Printer.Print inmatgrid.TextMatrix(i, 1);
            Printer.CurrentX = col3x
            Printer.Print inmatgrid.TextMatrix(i, 2);
            Printer.CurrentX = col4x
            Printer.Print inmatgrid.TextMatrix(i, 3);
            Printer.CurrentX = col5x
            Printer.Print inmatgrid.TextMatrix(i, 4)
            
            lc = lc + 1
            If lc >= 54 Then
                Printer.NewPage
                ppage = ppage + 1
                Printer.Print header
                lc = 0
            End If
        Next
        Printer.Print
    End If
    
'****************************
'Print quantity discrepancies
'****************************

    If inshogrid.TextMatrix(0, 1) <> "" Then
        col1x = 300
        col2x = 900
        col3x = 3000
        col4x = 4000
        col5x = 5000

        Printer.CurrentX = col1x
        Printer.FontSize = 14
        Printer.Print "   Quantity Discrepencies:"
        Printer.Print "   -----------------------"
        Printer.FontSize = 11
        
        Printer.CurrentX = col2x
        Printer.Print "JobNum";
        
        Printer.CurrentX = col3x
        Printer.Print "ProdQty";
        
        Printer.CurrentX = col4x
        Printer.Print "ShipQty";
        
        Printer.CurrentX = col5x
        Printer.Print "RcvQty";
        
        For i = 1 To inshogrid.Rows - 1
            
            Printer.CurrentX = col1x
            Printer.Print inshogrid.TextMatrix(i, 0); ")";
            Printer.CurrentX = col2x
            Printer.Print inshogrid.TextMatrix(i, 1);
            Printer.CurrentX = col3x
            Printer.Print inshogrid.TextMatrix(i, 2);
            Printer.CurrentX = col4x
            Printer.Print inshogrid.TextMatrix(i, 3);
            Printer.CurrentX = col5x
            Printer.Print inshogrid.TextMatrix(i, 4)
            
            lc = lc + 1
            If lc >= 54 Then
                Printer.NewPage
                ppage = ppage + 1
                Printer.Print header
                lc = 0
            End If
        Next
    End If
        
    Printer.EndDoc
     
End Sub

Private Sub incomrefbut_Click()
    inopgrid.Rows = 0
    inshogrid.Rows = 0
    inmatgrid.Rows = 0
    Call Form_Load
End Sub

