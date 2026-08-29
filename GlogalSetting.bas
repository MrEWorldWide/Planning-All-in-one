Attribute VB_Name = "GlogalSetting"
Global Const MF_BYPOSITION = &H400&

Global ReadyToClose As Boolean
Global CNACCESS As New ADODB.Connection
Global CNBACCESS As New ADODB.Connection
Global CNASetup As New ADODB.Connection
Global CNVantage As New ADODB.Connection
Global CNAS400 As New ADODB.Connection
Global CNAUSER As New ADODB.Connection
Global RSAccess As New ADODB.Recordset
Global RSAdd As New ADODB.Recordset
Global RSBACCESS As New ADODB.Recordset
Global RSReport As New ADODB.Recordset
Global RSAmountMet As New ADODB.Recordset
Global RSGen As New ADODB.Recordset
Global RSRead As New ADODB.Recordset
Public CfgFileName As String
Global StartTime As String
Global RetensionDays As String
Global Directory As String
Global ReplaceFileName As String
Global ReplaceBFileName As String
Global BackupFileName As String
Public SQLSelect As String
Public RSDriveConn As New ADODB.Recordset
Public RSConfig As New ADODB.Recordset
Public RSTime As New ADODB.Recordset
Public SaveDay As Integer
Public WildCardCounter As Integer
Public WildCardCounter2 As Integer
Public FS As Object
Public TS As Object
Public PN As Object
Public FirstRead As Boolean
Public DisplaySel As String

Global txtFileName As String
Global txtInDate As String
Global lblFileName As String
Public SaveReport As String
Global ResetHeadings As Boolean
Global ResetEXHeadings As Boolean
Global ResetReturn As Boolean
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

'
' API Types
'
' RECT is used to get the size of the panel we're inserting into
'
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'
' API Messages
'
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFFFFFF
Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetProfileString Lib "kernel32.dll" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public DefaultPrinter As String
Public WSNet As Object
Public PrinterList As Object
Public txtRetension As String


Public Sub Delay(Seconds As Long)
Dim n As Double

n = Timer + Seconds
Do While Timer < n
    DoEvents
Loop

End Sub
Function SetupConnection()
Dim SetupDBName As String
'Dim db As Connection

'Set db = New Connection
'db.CursorLocation = adUseClient
'db.Open "PROVIDER=MSDASQL;DRIVER={MERANT 3.60 32-BIT Progress SQL92 v9.1C};UID=SYSPROGRESS;DB=Vantage;PORT=3550;HOST=//Server;PWD=******"

On Error GoTo localerror
SetupDBName = "PROVIDER=MSDASQL;DRIVER={MERANT 3.60 32-BIT Progress SQL92 v9.1C};UID=odbcread;DB=Vantage;PORT=3550;HOST=192.168.100.124;PWD=******"
CNVantage.CursorLocation = adUseClient
CNVantage.Open SetupDBName
Exit Function
localerror:
Exit Function

End Function
Function LocalConnection()
Dim DBconnect As String

On Error GoTo localerror
FileName = App.Path & "V:\main\Users\Materials\KBLISTDB\kblist.accdb"
DBconnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Persist Security Info=False"
CNACCESS.Open DBconnect
Exit Function
localerror:
Exit Function

End Function

Public Function RightPad(StringIn As String, StrLen As Integer) As String
Dim LenCount As Integer

StringIn = Trim(StringIn)

While Len(StringIn) < StrLen
    StringIn = StringIn & " "
Wend
RightPad = StringIn
End Function

Function CountString(sText As String, sSearchFor As String, Optional bIgnoreCase As Boolean = True, Optional sIgnoreText As String) As Long
    Dim asItems() As String, lThisItem As Long
    On Error GoTo ErrFailed
    
    If bIgnoreCase Then
        asItems = Split(UCase$(sText), UCase$(sSearchFor))
        CountString = UBound(asItems)
    Else
        asItems = Split(sText, sSearchFor)
        CountString = UBound(asItems)
    End If
    
    If Len(sIgnoreText) Then
        'Deduct any items which contain the specified "sIgnoreText"
        For lThisItem = 0 To UBound(asItems) - 1
            If asItems(lThisItem) = sIgnoreText Then
                'Deduct this item
                CountString = CountString - 1
            End If
        Next
    End If
    WildCardCounter = CountString
    Exit Function

ErrFailed:
    'Error occurred
    Debug.Print "Error in CountString " & Err.Description
    Debug.Assert False
    CountString = 0

End Function

Public Function SQLTextPrep(strStringIn As String, strStringOut As String)
Dim strFind As String
Dim strReplace As String

  On Error GoTo HandleError:

  strFind = "'"
  strReplace = Chr(146)
  strStringOut = Trim(Replace(strStringIn, strFind, strReplace))
  
  Exit Function

HandleError:
   MsgBox "FUNCTION ERROR: SQLTextPrep " & Err.Number & " " & Err.Description
End Function
Public Function ChangeSingleQuote(strIn As String, strOut As String)

On Error GoTo SingleQuoteError

Dim StartNum As Integer, NextNum As Integer
Dim Str1 As String, Str2 As String
Dim StrQuote As String

StartNum = 1
Str1 = strIn
' CHARACTER CODE 34 = "
' CHARACTER CODE 39 = '

If InStr(Str1, Chr(39)) > 0 Then
    While InStr(NextNum + 1, Str1, Chr(39), 1)
        NextNum = InStr(StartNum, Str1, Chr(39), 1)
        Str2 = Str2 & Mid(Str1, StartNum, NextNum - StartNum) & Chr(39) & Chr(39)
        StartNum = NextNum + 1
    Wend
    Len1 = Len(Str1) - NextNum
    Str1 = Str2 & Right(Str1, Len1)
    Str2 = ""
End If

strOut = Str1
If Len(Str1) > 251 Then
    MsgBox ("This Description has a ( ' ) in it. " & (Chr(13) & Chr(10)) & "To Process it correctly, this process will add another ( ' ) for each of these used in the description, which will make it more than 30 characters." & (Chr(13) & Chr(10)) & "Each time you use a ( ' ) you will need to subtract 1 more character for each ( ' ) used in the description, please correct.")
End If
Exit Function
SingleQuoteError:
Debug.Print Err.Number
Debug.Print Err.Description

End Function
Private Function fnCheckQuotes(str As String) As Boolean
   On Error GoTo LOCALERRORHANDLER
   'Checks the supplied string to see if it contains
   'both single and double quotes

   fnCheckQuotes = False
   
   'Contains a double quote
   If InStr(1, str, Chr(34), vbTextCompare) <> 0 Then
         'Contains a single quote
      If InStr(1, str, Chr(39), vbTextCompare) <> 0 Then
         fnCheckQuotes = True
      End If
   End If
   Exit Function

LOCALERRORHANDLER:
End Function


Function GetDefaultPrinter() As String
  Dim strBuffer As String * 254
  Dim iRetValue As Long
  Dim strDefaultPrinterInfo As String
  Dim tblDefaultPrinterInfo() As String
  
  ' Retreive current default printer information
  iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
  strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
  tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
  GetDefaultPrinter = tblDefaultPrinterInfo(0)
End Function

