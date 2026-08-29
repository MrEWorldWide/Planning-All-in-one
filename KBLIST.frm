VERSION 5.00
Begin VB.Form KBLIST 
   Caption         =   "Manage the KanBan list"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exitbut 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton addbut 
      Caption         =   "Add"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox partlist 
      Height          =   2595
      ItemData        =   "KBLIST.frx":0000
      Left            =   120
      List            =   "KBLIST.frx":0002
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "KBLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addbut_Click()
    
    
    
    Dim str, SQLInsert As String
    
    str = InputBox("Enter a part number:", "Add Part")
    If str = "" Then
        Exit Sub
    Else
        LocalConnection
        SQLInsert = "INSERT INTO Parts(PartNum) VALUES('" & str & "')"
        RSRead.Open (SQLInsert), CNACCESS
        Debug.Print SQLInsert
        partlist.Clear
        CNACCESS.Close
        Call Form_Load
        
    End If
    
End Sub

Private Sub Command1_Click()
    
    LocalConnection
    Dim sqldelete As String
    sqldelete = "DELETE FROM Parts WHERE PartNum='" & partlist.Text & "'"
    With partlist
        RSRead.Open (sqldelete), CNACCESS
    End With
    partlist.RemoveItem (partlist.ListIndex)
    CNACCESS.Close
    Call Form_Load
    
End Sub

Private Sub exitbut_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    DoEvents
    
    LocalConnection
    
    Dim c, i As Long
    SQLSelect = "SELECT PartNum FROM Parts ORDER BY PartNum ASC"
    RSRead.Open (SQLSelect), CNACCESS

    'Do Until RSRead.EOF
     '   RSRead.MoveNext
    '    i = i + 1
   ' Loop

    
    
    If RSRead.EOF Then
        partlist.AddItem ("None")
    Else
        Do Until RSRead.EOF

            partlist.AddItem RSRead!PartNum
            RSRead.MoveNext
        Loop
        partlist.Refresh
        
    End If
    

    RSRead.Close
    CNACCESS.Close
    
End Sub
