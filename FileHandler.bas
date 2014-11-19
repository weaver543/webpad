Attribute VB_Name = "FileHandler"
'coloring at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=40401&lngWId=1
'this requires richtx32.ocx 5/7/99 v6.00.8418 with size=204296)

Option Explicit

Public ColorColl As Collection

Public Const VERSION = "© 5/10/14 (Kevin Ryan)"
'Public Const VERSION = "© 6/13/13 (Kevin Ryan)" ' fixed bug at frmmain line 1529
'Public Const VERSION = "© 12/15/10 (Kevin Ryan)"
Public Const APP_NAME = "Webpad"
Public Const INI_FILE = "webpad.ini"
Public Const SEARCH_FILE = "search.ini"

Public Const WM_PASTE = &H302
Public Const VK_LMENU = &HA4

Public Const MAX_SEARCHES = 14
Public Const MAX_KEYS = 50

Public bookmarkstr(10) As String
Public numkeys As Integer
Public control_forced_up As Boolean
Public control_down As Boolean
Public shift_down As Boolean
Public alt_forced_up As Boolean

Public Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public keystate(256) As Byte

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const MAX_LOG_CHANGES = 300
Public Const MAX_KEYFUNCTIONS = 16
Public key(MAX_KEYS) As Integer
Public keyfunc(MAX_KEYS) As String
Public descrip(MAX_KEYS) As String

Public param
Public strr As String
Public i As Integer
Public text_changed As Boolean
Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Public Const EM_GETSELTEXT = &H43E

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (Prop As SHELLEXECUTEINFO) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long
Public Const CB_FINDSTRINGEXACT = &H158
Public Const EM_SCROLL = &HB5
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public FileChanged As Boolean
Public ChangeState As Boolean
Public NoStatusUpdate  As Boolean
Public newsearch As Boolean
Public scrollbars As Boolean
Public txtvisible As Boolean
Public menuvisible As Boolean

Public Function SpecialFolder(ByVal CSIDL As Long) As String
'used to locate 'Send to'
Dim r As Long
Dim sPath As String
Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260
r = SHGetSpecialFolderLocation(frmMain.hwnd, CSIDL, IDL)
If r = NOERROR Then
    sPath = Space$(MAX_LENGTH)
    r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    If r Then
        SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End If
End Function

Public Function FileExists(sSource As String) As Boolean
    'does this file exist ?
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
'Binary file writing
Public Sub FileSave(Text As String, filepath As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
'MsgBox "filesaves filepath: " & filepath
    Open filepath For Binary As #f
    Put #f, , Text
    Close #f
    frmMain.changelog = ""
End Sub
Public Function PathOnly(ByVal filepath As String) As String
    Dim temp As String
    temp = Mid$(filepath, 1, InStrRev5(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev5(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev5(filepath, ".") + 1)
    If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = Mid$(filepath, 1, InStrRev5(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Public Function GetFileSize(zLen As Long) As String
    'just returns a user friendly string of the filesize
    Dim tmp As String
    Const KB As Double = 1024
    Const MB As Double = 1024 * KB
    Const GB As Double = 1024 * MB
    If zLen < KB Then
        tmp = Format(zLen) & " bytes"
    ElseIf zLen < MB Then
        tmp = Format(zLen / KB, "0.00") & " KB"
    Else
        If zLen / MB > 1024 Then
            tmp = Format(zLen / GB, "0.00") & " GB"
        Else
            tmp = Format(zLen / MB, "0.00") & " MB"
        End If
    End If
    GetFileSize = Chr(32) + tmp + Chr(32)
End Function

Public Sub SetScrollPos(mPos As Long, mRTF As RichTextBox)
  Dim CurLineCount As Long, curvl As Long, lastvl As Long
  CurLineCount = SendMessage(mRTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
  curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
  'first use pageup/pagedown to get close to the
  'position quickly
  If mPos < curvl Then
    Do Until curvl < mPos
        SendMessage mRTF.hwnd, EM_SCROLL, 2, 0
        curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
        If curvl = 0 Or curvl = CurLineCount Then Exit Do
    Loop
  Else
    Do Until curvl > mPos
        SendMessage mRTF.hwnd, EM_SCROLL, 3, 0
        curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
        If curvl = 0 Or curvl = CurLineCount Or lastvl = curvl Then Exit Do
        lastvl = curvl
    Loop
  End If
  'do fine adjustment to get position exact
  Do Until curvl = mPos
    If mPos < curvl Then
        SendMessage mRTF.hwnd, EM_SCROLL, 0, 0
    Else
        SendMessage mRTF.hwnd, EM_SCROLL, 1, 0
    End If
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If curvl = 0 Or curvl = CurLineCount Or lastvl = curvl Then Exit Do
    lastvl = curvl
  Loop
End Sub
Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, " ")
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function

'vb5 implementation of split() in vb6
Public Function Split5(ByVal sIn As String, _
  Optional sDelim As String, Optional nLimit As Long = -1, _
  Optional bCompare As VbCompareMethod = vbBinaryCompare) As Variant
  Dim sOut() As String
  Dim sRead As String, nC As Integer

  If sDelim = "" Then sDelim = " "
   
  If InStr(sIn, sDelim) = 0 Then
    ReDim sOut(0) As String
    sOut(0) = sIn
    Split5 = sOut
    Exit Function
  End If

  sRead = ReadUntil(sIn, sDelim, bCompare)

  Do
    ReDim Preserve sOut(nC)
    sOut(nC) = sRead
    nC = nC + 1
    If nLimit <> -1 And nC >= nLimit Then Exit Do
    sRead = ReadUntil(sIn, sDelim)
  Loop While sRead <> "~TWA"
  
  ReDim Preserve sOut(nC)
  sOut(nC) = sIn
  Split5 = sOut
End Function
' used by split5()
Private Function ReadUntil(ByRef sIn As String, sDelim As String, Optional bCompare As VbCompareMethod = vbBinaryCompare) As String
  Dim nPos As Long
  nPos = InStr(1, sIn, sDelim, bCompare)
  If nPos > 0 Then
     ReadUntil = Left(sIn, nPos - 1)
     sIn = Mid(sIn, nPos + Len(sDelim))
  Else
     ReadUntil = "~TWA"
  End If
End Function

Public Function Inc(ByRef i As Integer) As Integer
  Inc = i
  i = i + 1
End Function


' the following function code is from Microsoft article Q188007
Public Function Join5(Source() As String, Optional _
      sDelim As String = " ") As String
Dim sOut As String, iC As Integer
On Error GoTo errh:
    For iC = LBound(Source) To UBound(Source) - 1
        sOut = sOut & Source(iC) & sDelim
    Next
    sOut = sOut & Source(iC)
    Join5 = sOut
    Exit Function
errh:
    Err.Raise Err.Number
End Function


Public Function StrReverse5(ByVal sIn As String) As String
  Dim nC As Integer, sOut As String
  For nC = Len(sIn) To 1 Step -1
  sOut = sOut & Mid(sIn, nC, 1)
  Next
  StrReverse5 = sOut
End Function

Public Function InStrRev5(ByVal sIn As String, sFind As String, _
 Optional nStart As Long = 1, Optional bCompare As _
      VbCompareMethod = vbBinaryCompare) As Long
    Dim nPos As Long
    sIn = StrReverse5(sIn)
    sFind = StrReverse5(sFind)
    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        InStrRev5 = 0
    Else
        InStrRev5 = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function

  Public Function Replace5(sIn As String, sFind As String, _
    sReplace As String, Optional nStart As Long = 1, _
    Optional nCount As Long = -1, Optional bCompare As _
    VbCompareMethod = vbBinaryCompare) As String

    Dim nC As Long, nPos As Integer, sOut As String
    sOut = sIn
    nPos = InStr(nStart, sOut, sFind, bCompare)
    If nPos = 0 Then GoTo EndFn:
    Do
      nC = nC + 1
      sOut = Left(sOut, nPos - 1) & sReplace & _
         Mid(sOut, nPos + Len(sFind))
      If nCount <> -1 And nC >= nCount Then Exit Do
      nPos = InStr(nStart, sOut, sFind, bCompare)
    Loop While nPos > 0
EndFn:
    Replace5 = sOut
  End Function

Public Function strUnQuoteString(ByVal strQuotedString As String)
    'pulled this from the P&D Wizard source
    strQuotedString = Trim$(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = Chr(34) Then
        If Right$(strQuotedString, 1) = Chr(34) Then
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function


Public Sub HighLightSelection(mForm As Form, mRTF As RichTextBox, mHighLightColor As Long, Optional DontLock As Boolean)
    'This is trickier than the other Highlight functions because
    'we have to allow for existing highlighting in various colors
    
    Dim TempRTF As String
    Dim SelStart As Long
    Dim SelEnd As Long
    Dim SelectedText As String
    Dim BeforeHL As String
    Dim AfterHL As String
    Dim FirstSelHL As String
    Dim LastSelHL As String
    Dim StartReplaceHL As String
    Dim EndReplaceHL As String
    Dim TempNum As String
    Dim z As Long
    Dim st As Long
    Dim Found As Long
    Dim HLNum As Long
    Dim RepairCtbl As Boolean
    Dim OldCol As Long
    If mRTF.SelLength = 0 Then Exit Sub
    st = mRTF.SelStart
    Found = mRTF.SelLength
'    If Not DontLock Then LockWindowUpdate mForm.hWnd
    'Locate the chosen color in the Colortable
    GetColorTable mRTF
    For z = 1 To ColorColl.count
        If ColorColl(z) = mHighLightColor Then
            HLNum = z - 1
            Exit For
        End If
    Next
    'If we didn't find it then modify the content
    'to place the color in the Colortable
    If HLNum = 0 Then
        mRTF.SelStart = st
        mRTF.SelLength = 1
        OldCol = mRTF.SelColor
        mRTF.SelColor = mHighLightColor
        GetColorTable mRTF
        For z = 1 To ColorColl.count
            If ColorColl(z) = mHighLightColor Then
                HLNum = z - 1
                Exit For
            End If
        Next
        RepairCtbl = True
    End If
    mRTF.SelStart = st
    mRTF.SelLength = 0
    'Place markers around the selection
    mRTF.SelText = "%%%ZSTART%%%"
    mRTF.SelStart = st + Found + 12
    mRTF.SelText = "%%%ZENDBB%%%"
    TempRTF = mRTF.TextRTF
    SelStart = InStr(1, TempRTF, "%%%ZSTART%%%")
    SelEnd = InStr(1, TempRTF, "%%%ZENDBB%%%") + 12
    'Place the selected text RTF code in a variable
    SelectedText = Mid(TempRTF, SelStart, SelEnd - SelStart)
    
    'inspect the preceding RTF code for any highlighting
    z = InStrRev5(TempRTF, "\highlight", SelStart)
    'If there's highlighting, record its number(color index)
    If z <> 0 Then BeforeHL = Mid(TempRTF, z + 10, 1)
    
    'inspect the RTF code after the selection for any highlighting
    z = InStr(SelEnd, TempRTF, "\highlight")
    'If there's highlighting, record its number(color index)
    If z <> 0 Then AfterHL = Mid(TempRTF, z + 10, 1)
    
    'inspect the RTF code of the selection for any highlighting
    'find the first highlighting entry in the selection
    z = InStr(1, SelectedText, "\highlight")
    'If there's highlighting, record the first highlighting entry's number(color index)
    If z <> 0 Then FirstSelHL = Mid(SelectedText, z + 10, 1)
    'find the last highlighting entry in the selection
    z = InStrRev5(SelectedText, "\highlight")
    If z <> 0 Then
        'if found record it's number(color index)
        LastSelHL = Mid(SelectedText, z + 10, 1)
        'Ok, we've got all the selections highlighting recorded
        'now we remove ALL highlighting from the selection
        Do
            TempNum = Mid(SelectedText, z + 10, 1)
            SelectedText = Replace5(SelectedText, "\highlight" & TempNum & " ", "", , 1)
            z = InStr(1, SelectedText, "\highlight")
            If z = 0 Then Exit Do
        Loop
        'retuen the altered seleted RTF code back to the entire RTF code
        TempRTF = Left(TempRTF, SelStart - 1) & SelectedText & Right(TempRTF, Len(TempRTF) - SelEnd + 1)
    Else
        'If there was no highlighting in the selection then
        'use any highlighting data from BEFORE the selection
        If BeforeHL <> "" And BeforeHL <> "0" Then
            LastSelHL = BeforeHL
        End If
    End If
    
    'Now to replace our markers with the appropriate RTF tags according to
    'the highlighting tags found before/in/after the selection
    
    'Prepare the replacement strings
    StartReplaceHL = IIf(BeforeHL = "0" Or BeforeHL = "", "\highlight" & HLNum & " ", "\highlight0 " & "\highlight" & HLNum & " ")
    EndReplaceHL = IIf(LastSelHL = "0" Or LastSelHL = "", "\highlight0 ", "\highlight0 " & "\highlight" & LastSelHL & " ")
    'Do the replacing
    TempRTF = Replace5(TempRTF, "%%%ZSTART%%%", StartReplaceHL)
    TempRTF = Replace5(TempRTF, "%%%ZENDBB%%%", EndReplaceHL)
    'return the RTF code to the richtextbox
    mRTF.TextRTF = TempRTF
    'Return any adjustments back to what it was
    If RepairCtbl Then
        mRTF.SelStart = st
        mRTF.SelLength = 1
        mRTF.SelColor = OldCol
        mRTF.SelStart = 0
    End If
    mRTF.SelStart = st
    mRTF.Refresh
 '   If Not DontLock Then LockWindowUpdate 0
End Sub

Private Sub GetColorTable(mRTF As RichTextBox)
    Dim z As Long, z1 As Long, temp As String, tmp() As String, tmpCol() As String
    Set ColorColl = New Collection
    ColorColl.Add 0
    'Parse the RTF code to extract the Colortable
    z = InStr(1, mRTF.TextRTF, "{\colortbl")
    If z = 0 Then
        Exit Sub
    Else
        'Parse the Colortable to extract the colors used
        z1 = InStr(z, mRTF.TextRTF, "}")
        If z1 = 0 Then
            Exit Sub
        Else
            temp = Mid(mRTF.TextRTF, z, z1 - z + 1)
            'this broke with split5!  tmp = Split5(temp, ";")
            For z = 1 To UBound(tmp) - 1
                If tmp(z) <> "" Then
                    If Left(tmp(z), 1) = "\" Then tmp(z) = Right(tmp(z), Len(tmp(z)) - 1)
                    ' this broke with split5!  tmpCol = Split5(tmp(z), "\")
                    'Dump the colors found into a collection
                    ColorColl.Add RGB(Val(Right(tmpCol(0), Len(tmpCol(0)) - 3)), Val(Right(tmpCol(1), Len(tmpCol(1)) - 5)), Val(Right(tmpCol(2), Len(tmpCol(2)) - 4)))
                End If
            Next
        End If
    End If
End Sub

'Binary file reading
Public Function OneGulp(Src As String) As String
  On Error Resume Next
  Dim f As Integer, temp As String
  f = FreeFile
  DoEvents
  Open Src For Binary As #f
  temp = String(LOF(f), Chr$(0))
  Get #f, , temp
  Close #f
  'check for unicode - some older .reg files for example
  If Left(temp, 2) = "ÿþ" Or Left(temp, 2) = "þÿ" Then
    MsgBox "kwr error at 2838"  ' since no replace command in vb5!
    'temp = Replace(Right(temp, Len(temp) - 2), Chr(0), "")
  End If
  OneGulp = temp
End Function


