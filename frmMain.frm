VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "WebPad"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8385
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox rtf2 
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   0
      Width           =   1215
      Begin RichTextLib.RichTextBox RTF 
         Height          =   615
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         _Version        =   393217
         ScrollBars      =   3
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"frmMain.frx":094D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Left            =   7560
      Top             =   240
   End
   Begin VB.Timer PasteTimer 
      Interval        =   100
      Left            =   7560
      Top             =   720
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   5475
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8123
            Text            =   "File not saved"
            TextSave        =   "File not saved"
            Object.ToolTipText     =   "File path"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Cursor position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Text            =   "0"
            TextSave        =   "0"
            Object.ToolTipText     =   "Selection length"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "2 bytes"
            TextSave        =   "2 bytes"
            Object.ToolTipText     =   "File size"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEditBase 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuCopyrightpane 
         Caption         =   "Copy right pane"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu next_menu 
         Caption         =   "Find Next"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Options"
      Begin VB.Menu mnu_shortcut 
         Caption         =   "Customize shortcut keys"
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "&Split Pane (Alt-S)"
      End
      Begin VB.Menu ontop 
         Caption         =   "Always On Top"
      End
      Begin VB.Menu mnuFormatWordwrap 
         Caption         =   "Wordwrap"
      End
      Begin VB.Menu mnuFormatSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuResetFont 
         Caption         =   "Reset Font"
      End
      Begin VB.Menu mnuFormatBackcolor 
         Caption         =   "Backcolor"
      End
      Begin VB.Menu mnuFormatSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRememberPos 
         Caption         =   "Remember last position"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuchangeicon 
         Caption         =   "Change Icon"
      End
      Begin VB.Menu mnu_setdefaults 
         Caption         =   "Set defaults for this file"
      End
      Begin VB.Menu mnuFormatStats 
         Caption         =   "Document Statistics"
      End
      Begin VB.Menu mnuFilterRTF 
         Caption         =   "Filter RTF"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_changelog 
         Caption         =   "Show change log"
      End
      Begin VB.Menu mnuMenubar 
         Caption         =   "Menubar (Alt-M)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuScrollbars 
         Caption         =   "Horiz Scrollbar (Alt-L)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewTitlebar 
         Caption         =   "Compact mode (Alt-C)"
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "Statusbar (Alt-T)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnu_help 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "right click menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAngelfire 
         Caption         =   "Open Web Site"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private junktmp As Integer

Private b_left As Integer
Private b_top As Integer
Private b_width As Integer
Private b_height As Integer

Private L As Long
Private DOSfiletype As Boolean
Private reminded As Boolean
Private compact_mode As Boolean
Private lStyle As Long
Public changelog As String
Private selecting_down As Boolean, selecting_up As Boolean
Private dont_save_settings As Boolean ' dont save when loading file that has defaults in ini file
Private settings_string As String
Private col As Long
Private newstart As Long
Private orig_start As Long
Public macro_function As Integer
Public current_file As String
Public remap_key As Integer
Public searching As Boolean
Public separator As String
Private Const CURSORWIDTH = 6
Private Const CURSORHEIGHT = 11
Private Const SW_SHOW = 5
Private Const SW_SHOWNORMAL = 1
Private Const EM_SCROLL& = &HB5
Private Const SB_LINEDOWN& = 1
Private Const SB_LINEUP& = 0
Private Const EM_LINESCROLL& = &HB6
Private Const EM_SCROLLCARET& = &HB7
Private Const SB_PAGEDOWN& = 3
Private Const SB_PAGEUP& = 2
Private Const EM_GETFIRSTVISIBLELINE& = &HCE
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Const VK_CONTROL = &H11
Private Const EM_LINEFROMCHAR = &HC9
Private Const MAX_SAVSTRINGS = 5
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_PASTE = &H302

Private Const WM_VSCROLL = &H115

'Private Const SB_LINEDOWN = 1
'Private Const SB_LINEUP = 0
'Call SendMessage(Text1.hwnd, WM_VSCROLL, SB_LINEDOWN, 0)


Dim savstring(MAX_SAVSTRINGS) As String
Private num_savstrings As Integer

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
   String, ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
   "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
   String, ByVal lpResult As String) As Long

Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, _
    ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, _
    ByVal wBar As Long, ByVal bShow As Long) As Long
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_BOTH = 3

Private Const GWL_STYLE = -16&
Private Const WS_CAPTION = &HC00000
Private Const WS_BORDER = &H800000
Private Const WS_SIZEBOX = &H40000

Public WithEvents Undo As clsUndo 'heavily modified version of a class by Sebastian Thomschke
Attribute Undo.VB_VarHelpID = -1
Public onlyLoading As Boolean 'indicates form load complete
Dim mTStop() As Boolean 'allows the use of 'tab' within the richtextbox
'rather than moving focus to the next control
Dim myCommand As String

Sub adjust_cursor()
  Exit Sub ' until I figure out best place to put this routine
  Call CreateCaret(RTF.hwnd, 0, CURSORWIDTH, CURSORHEIGHT)
  Call ShowCaret(RTF.hwnd)
End Sub

Private Sub mnu_changelog_Click()
  MsgBox changelog
End Sub

Private Sub mnu_setdefaults_Click()
  dont_save_settings = True
  For i = 0 To num_savstrings - 1
    param = Split5(savstring(i), "=")
    param = Split5(param(1), "|")

    If param(0) = current_file Then
      savstring(i) = "default=" & current_file & "|" & _
        param(1) & "|" & CStr(Me.Left) & "|" & CStr(Me.Top) & "|" & _
        CStr(Me.Width) & "|" & CStr(Me.Height) & "|" & param(6)
      Exit For
    End If
  Next
End Sub

Private Sub mnu_shortcut_Click()
  frmCustomize.Show
End Sub

Private Sub mnuchangeicon_Click()
  With cmndlg
    .filefilter = "Icons (*.ico)"
    OpenFile
    If Len(.filename) <> 0 Then Set Me.Icon = LoadPicture(.filename)
  End With
End Sub

Private Sub mnuCopy_Click()
    SendMessage RTF.hwnd, WM_COPY, 0, 0
End Sub

Private Sub mnuCopyrightpane_Click()
    SendMessage rtf2.hwnd, WM_COPY, 0, 0
End Sub

Private Sub mnuCut_Click()
    SendMessage RTF.hwnd, WM_CUT, 0, 0
    'RTF.Text = ""
End Sub

Private Sub mnuFiles_Click(Index As Integer)
  load_new_file (mnuFiles(Index).Caption)  ' menu to load preset file
End Sub

Private Sub mnuFilterRTF_Click()
  mnuFilterRTF.Checked = Not mnuFilterRTF.Checked
End Sub


Private Sub mnuMenubar_Click()
  menuvisible = Not menuvisible
  mnuMenubar.Checked = menuvisible
  mnuFile.Visible = menuvisible
  mnuEditBase.Visible = menuvisible
  mnuFormat.Visible = menuvisible
  mnuhelp.Visible = menuvisible
End Sub

Private Sub mnuPaste_Click()
    SendMessage RTF.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub mnuRememberPos_Click()
  mnuRememberPos.Checked = Not mnuRememberPos.Checked
End Sub

Private Sub mnuResetFont_Click()
  st = RTF.SelStart
  RTF.SelStart = 0
  RTF.SelLength = Len(RTF.Text)
  
  RTF.SelFontName = "Lucida Console"
  RTF.SelFontSize = 10
  
  RTF.SelStart = st
  RTF.SelLength = 0
End Sub

Private Sub mnuScrollbars_Click()
 mnuScrollbars.Checked = Not mnuScrollbars.Checked
' scrollbars = Not scrollbars
 ShowScrollBar RTF.hwnd, SB_HORZ, mnuScrollbars.Checked
End Sub

Private Sub resize_rtf2()
Debug.Print "kwrr 1 rtf1 width: " + Str(RTF.Width) + " rtf2: " + Str(RTF.Width)
  midpoint = RTF.Width / 2
  RTF.Width = midpoint
  rtf2.Left = midpoint + 1
  rtf2.Width = RTF.Width
  rtf2.Height = RTF.Height
Debug.Print "kwrr 2 rtf1 width: " + Str(RTF.Width) + " rtf2: " + Str(RTF.Width)
End Sub

Private Sub mnuSplit_Click()
  rtf2.Visible = Not rtf2.Visible
  If rtf2.Visible Then
    resize_rtf2
  Else
    RTF.Width = PicLeft.Width - 120
  End If
End Sub

Private Sub mnuViewTitlebar_Click()
  
  compact_mode = Not compact_mode
  
'  SB.Visible = Not compact_mode
'  mnuFile.Visible = Not compact_mode
'  mnuEditBase.Visible = Not compact_mode
'  mnuFormat.Visible = Not compact_mode
'  mnuhelp.Visible = Not compact_mode
  
  If compact_mode Then
    ' turn off menu first
    menuvisible = True
    mnuMenubar_Click
    
    ' turn off statusbar
    mnuViewStatusbar.Checked = True
    mnuViewStatusbar_Click
    
    oriStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    lStyle = oriStyle And (Not WS_CAPTION)
    SetWindowLong Me.hwnd, GWL_STYLE, lStyle
  Else
    menuvisible = False
    mnuMenubar_Click
    mnuViewStatusbar_Click
    Me.Hide

    lStyle = oriStyle Or WS_CAPTION Or WS_SIZEBOX
    SetWindowLong Me.hwnd, GWL_STYLE, lStyle
    Me.Show
  End If

End Sub

Private Sub ontop_Click()
  ontop.Checked = Not ontop.Checked
  If ontop.Checked Then
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SB.Panels(1) = "Always on top"
    Set frmFind.Icon = Me.Icon ' save old icon to restore
    Set Me.Icon = frmCustomize.Icon
  Else
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SB.Panels(1) = "No longer always on top"
    Set Me.Icon = frmFind.Icon
  End If
End Sub

Private Sub PicLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReleaseCapture
  SendMessage hwnd, &H112, &HF012&, 0&
End Sub


Private Sub RTF_Click()
  force_control_up
End Sub

Private Sub RTF_DblClick()
  Call test_url
End Sub

Function test_url() ' see if url is at cursor
  Dim lOrigin As Long, lFinal As Long, filelength
  On Error GoTo error:
  filelength = Len(RTF.Text)
  
  test_url = False
  SB.Panels(1) = "checking url..."

  ' get positions of beginning & end of word (space delimited)
  lOrigin = RTF.SelStart
  lFinal = lOrigin + 1
  Do While (lOrigin > 1 And Mid(RTF.Text, lOrigin, 1) <> " " And Asc(Mid(RTF.Text, lOrigin, 1)) <> 10)
    lOrigin = lOrigin - 1
  Loop
  If lOrigin > 1 Then lOrigin = lOrigin + 1
  Do While (Mid(RTF.Text, lFinal, 1) <> " " And Asc(Mid(RTF.Text, lFinal, 1)) <> 13 And lFinal < filelength)
    lFinal = lFinal + 1
  Loop
  
  If lFinal <> filelength Then lFinal = lFinal - 1

  SB.Panels(1) = ""
  If (lFinal - lOrigin > 3) Then
    If Mid(RTF.Text, lOrigin, 4) = "http" Or Mid(RTF.Text, lOrigin, 3) = "www" Or _
        Mid(RTF.Text, lOrigin, 5) = "<http" Then
        'Mid(RTF.Text, lOrigin, 7) = "telnet:" Then
      open_url (Mid(RTF.Text, lOrigin, lFinal - lOrigin + 1))
      test_url = True
    ElseIf Mid(RTF.Text, lOrigin + 1, 2) = ":\" Then
      Shell Mid(RTF.Text, lOrigin, lFinal - lOrigin + 1), vbNormalFocus
      'Call ShellExecute(Me.hwnd, "Open", "c:\\junk.txt", 0&, 0&, SW_SHOWMAXIMIZED)

      'Call ShellExecute(Me.hwnd, "open", BrowserExec, Mid(RTF.Text, lOrigin, lFinal - lOrigin + 1), Dummy, SW_SHOWNORMAL)
    ElseIf InStr(Mid(RTF.Text, lOrigin, lFinal - lOrigin + 1), "@") Then
      ShellExecute 0, vbNullString, "mailto:" & Mid(RTF.Text, lOrigin, lFinal - lOrigin + 1), vbNullString, vbNullString, vbNormalFocus
      test_url = True
    ElseIf Mid(RTF.Text, lOrigin, 7) = "TodayIs" Then
      RTF.SelStart = lOrigin - 1
      todaystr = Choose(Weekday(Now()), "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
      tmp_str = "TodayIs " & todaystr & " " & Date
      i = 0
      Do While (Asc(Mid(RTF.Text, lOrigin + i, 1)) <> 13)
        i = i + 1
      Loop
      RTF.SelLength = i
      RTF.SelText = tmp_str
      test_url = True
    End If
  End If
  Return
error:
  
End Function


Private Sub say(strr As String)
  RTF.SelStart = Len(RTF.Text)
  RTF.Text = RTF.Text + strr + vbCrLf
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
  
  Call adjust_cursor
  
  SB.Panels(1) = ""
  If Len(changelog) < MAX_LOG_CHANGES Then changelog = changelog + Chr(KeyAscii)
End Sub


Public Sub force_alt_up()
  ' release the alt key state from "pressed"
  GetKeyboardState keystate(0)
  keystate(VK_LMENU) = keystate(VK_LMENU) And &H7F
  SetKeyboardState keystate(0)
  alt_forced_up = True
End Sub

Public Sub resume_alt_down()
  GetKeyboardState keystate(0)
  keystate(VK_LMENU) = keystate(VK_LMENU) Or &H80
  SetKeyboardState keystate(0)
  alt_forced_up = False
End Sub
Public Sub force_control_up()
  ' release the control key state from "pressed"
  GetKeyboardState keystate(0)
  keystate(VK_CONTROL) = keystate(VK_CONTROL) And &H7F
  SetKeyboardState keystate(0)
  control_forced_up = True
End Sub

Public Sub resume_control_down()
  GetKeyboardState keystate(0)
  keystate(VK_CONTROL) = keystate(VK_CONTROL) Or &H80
  SetKeyboardState keystate(0)
  control_forced_up = False
End Sub

Private Sub force_shift_up()
  GetKeyboardState keystate(0)
  keystate(VK_SHIFT) = keystate(VK_SHIFT) And &H7F
  SetKeyboardState keystate(0)
  shift_forced_up = True
End Sub

Private Sub resume_shift_down()
  GetKeyboardState keystate(0)
  keystate(VK_SHIFT) = keystate(VK_SHIFT) Or &H80
  SetKeyboardState keystate(0)
  shift_forced_up = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  Call adjust_cursor
  
  If Shift = 4 And KeyCode = 18 Then force_control_up
  
  If KeyCode = vbKeyControl Then
    control_down = False
    tmp_str = SB.Panels(5)
    Mid(tmp_str, 1, 1) = " "
    SB.Panels(5) = tmp_str
  End If
  
  If KeyCode = vbKeyShift Then
    shift_down = False
    tmp_str = SB.Panels(5)
    Mid(tmp_str, 2, 1) = " "
    SB.Panels(5) = tmp_str
  End If
End Sub

Private Sub mnu_help_Click()
  Dim helptext As String
  helptext = "if shift/alt/number will assign numeric key to bookmark (last search string)" & vbCrLf & _
    "Alt-number: jump to bookmark" & vbCrLf & _
    "Alt-V: Paste under cursor" & vbCrLf & _
    "Alt-N to switch to Normal size" & vbCrLf & _
    "Alt-S to switch to small size" & vbCrLf & _
    "Alt-A toggles 'always on top'" & vbCrLf & _
    "Alt-M to toggle statusbar/menubar" & vbCrLf & _
    "Alt-W toggles wordwrap" & vbCrLf & _
    "Alt-Q: save and exit" & vbCrLf & _
    "Alt-X to exit without changes" & vbCrLf

MsgBox helptext
End Sub

Private Sub goto_bol()
  For L = RTF.SelStart To 1 Step -1
    If Asc(Mid(RTF.Text, L, 1)) = 10 Then Exit For
  Next

  RTF.SelStart = L
End Sub

Private Sub goto_eol()
  For L = RTF.SelStart To Len(RTF.Text)
    If (Asc(Mid(RTF.Text, L, 1)) = 13) Then Exit For
  Next
  RTF.SelStart = L
End Sub

Private Function InStrMult(start As Integer, fullstring As String, _
    expr1 As String, expr2 As String)
  ret1 = InStr(start, fullstring, expr1, 0)
  ret2 = InStr(start, fullstring, expr2, 0)
  If ret1 < ret2 Then
    InStrMult = ret1
  Else
    InStrMult = ret2
  End If
End Function


Private Function orig_goto_next_word() ' slower since it does NOT use InStr
  If Not DOSfiletype Then
    goto_next_word = False
    Return
  End If
    
  text_reached = False
  For L = RTF.SelStart + 1 To Len(RTF.Text) - 1
    If Asc(Mid(RTF.Text, L, 1)) <> 32 Then text_reached = True
    
    If (Asc(Mid(RTF.Text, L, 1)) = 13) Then
      L = L - 1
      Exit For ' stop at CR
    End If
    If text_reached And (Asc(Mid(RTF.Text, L, 1)) = 32) And (Asc(Mid(RTF.Text, L + 1, 1)) <> 32) Then Exit For
  Next
  
  If RTF.SelStart = L Then
    goto_next_word = False
  Else
    goto_next_word = True ' indicate that it's advanced
    RTF.SelStart = L
  End If
End Function


Private Sub RTF_KeyDown(KeyCode As Integer, Shift As Integer)

  ' ctrl-A to select all
  If (Shift = 2) And KeyCode = vbKeyA Then
    RTF.SelStart = 0
    RTF.SelLength = Len(RTF.Text)
    Exit Sub
  End If

  Call adjust_cursor

  If KeyCode = 13 And Len(changelog) < MAX_LOG_CHANGES Then changelog = changelog + vbCrLf
  
  SelFont
  SB.Panels(2) = RTF.SelStart
  SB.Panels(3) = RTF.SelLength

  If (DOSfiletype And (Shift = 1) And ((KeyCode = vbKeyDown) Or (KeyCode = vbKeyUp))) Then
    
    If selecting_down Then
      If KeyCode = vbKeyDown Then
Debug.Print "origstart: " + Str(RTF.SelStart) + " sellength: " + Str(RTF.SelLength)
        L = InStr(RTF.SelStart + RTF.SelLength + 1, RTF.Text, Chr(13), 0) - 1
        If L < 1 Then Exit Sub
        RTF.SelLength = L - RTF.SelStart + 2
Debug.Print "  L: " + Str(L) + " newstart: " + Str(RTF.SelStart) + " sellength: " + Str(RTF.SelLength)
      Else ' up arrow so user is subtracting from bottom of selection
        L = InStrRev(RTF.Text, Chr(10), RTF.SelStart + RTF.SelLength - 2, 0)
        If L >= RTF.SelStart Then RTF.SelLength = L - RTF.SelStart
      End If
    ElseIf selecting_up Then
      If KeyCode = vbKeyUp Then
        L = InStrRev(RTF.Text, Chr(10), RTF.SelStart - 1, 0)
        newlength = RTF.SelLength + RTF.SelStart - L
        If L >= 0 Then RTF.SelStart = L
        If newlength >= 0 Then RTF.SelLength = newlength
      Else ' down arrow so user is subtracting from top of selection
        L = InStr(RTF.SelStart + 2, RTF.Text, Chr(13), 0)
        newlength = RTF.SelLength - (L - RTF.SelStart) - 1
        RTF.SelStart = L + 1
        If newlength >= 0 Then RTF.SelLength = newlength
      End If
    Else
      goto_bol

      L = InStr(RTF.SelStart + 1, RTF.Text, Chr(13), 0)

      If L = 0 Then Exit Sub
      
      RTF.SelLength = L - RTF.SelStart + 1
      
      If KeyCode = vbKeyDown Then
        selecting_down = True
      Else
        selecting_up = True
      End If
    End If
    
    KeyCode = 0
  Else
    If KeyCode <> 16 Then ' if not shift down
      selecting_down = False
      selecting_up = False
    End If
  End If
  
End Sub

Private Sub daily_reminder()
  param = Split5(Date, "/200")
  reminded = True
  force_control_up
  control_down = False
  
  ' to be detected, date must have leading & trailing space (like " 9/11 ")
  ' and be in first 1000 lines of file
  Found = frmMain.RTF.Find(" " & param(0) & " ", 0, 1000, mWholeword Or mmatchCase)
  If Found > -1 Then
    SendKeys "+{END}", True
    j = Len(RTF.SelText) \ 2 ' use to center heading
    i = MsgBox(Space(j) & "REMINDER:" + vbCrLf + vbCrLf + RTF.SelText + vbCrLf + Space(j) & "Turn off?", vbYesNo)
    If i = 6 Then  ' 6=yes  7=no
      SendKeys "{HOME}{RIGHT}*", True ' insert * so no leading space anymore
      FileChanged = True
    End If
  End If
  SendKeys "^", True
  'force_control_up
  'resume_control_down
End Sub

Private Sub check_byte()
  L = L + 1
  Debug.Print Mid(RTF.Text, L, 1), Asc(Mid(RTF.Text, L, 1))
End Sub

Private Sub set_filetype()
  On Error GoTo WOOPS2
  
  For L = 1 To Len(RTF.Text)
    If Asc(Mid(RTF.Text, L, 1)) = 10 Then Exit For
  Next
  
  If Asc(Mid(RTF.Text, L - 1, 1)) = 13 Then
    DOSfiletype = True
  Else
    DOSfiletype = False
  End If

WOOPS2:
End Sub
Private Function goto_next_word()
  Dim cr_loc As Long
'Debug.Print vbCrLf & vbCrLf & "at 1; sel: " & Str(RTF.SelStart) & " is " & Mid(RTF.Text, RTF.SelStart + 1, 1)

  If Not DOSfiletype Then
    goto_next_word = False
    Return
  End If
  text_reached = False
              
              
  ' advance past spaces
  For L = RTF.SelStart + 1 To Len(RTF.Text) - 1
'    If Asc(Mid(RTF.Text, L, 1)) <> 32 Then
    If Mid(RTF.Text, L, 1) <> " " Then
        L = L - 1
        Exit For 'text_reached = True
    End If
  Next

  If L > RTF.SelStart Then ' cursor has traveled over spaces before reaching word
    RTF.SelStart = L
    Exit Function
  End If

  cr_loc = InStr(L + 1, RTF.Text, Chr(13), 0)
  If cr_loc = 0 Then cr_loc = Len(RTF.Text)
  L = InStr(L + 1, RTF.Text, " ", 0) ' advance to next space
  If L > cr_loc Then L = cr_loc
  
  While Mid(RTF.Text, L + 1, 1) = " " And L < cr_loc
    L = L + 1
  Wend
  If L = cr_loc Then L = cr_loc - 1

  If RTF.SelStart = L Then
    goto_next_word = False
  Else
    goto_next_word = True ' indicate that it's advanced
    RTF.SelStart = L
  End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim url_there As Boolean

  On Error Resume Next
  
  If KeyCode > 48 And KeyCode < 59 And Shift = 2 Then
    If (KeyCode - 49) < mnuFiles.UBound Then
      ' now close file and open mnuFiles(KeyCode - 48).Caption
      load_new_file (mnuFiles(KeyCode - 48).Caption)
    End If
  End If

  Call adjust_cursor

  If KeyCode = 122 Then L = RTF.SelStart ' F11
  If KeyCode = 123 Then MsgBox RTF.SelStart ' check_byte ' F12
  
  'shft=1 ctrl=2 alt=4
  If KeyCode = vbKeyControl Then
    control_down = True
    tmp_str = SB.Panels(5)
    Mid(tmp_str, 1, 1) = "C"
    SB.Panels(5) = tmp_str
  End If
  
  If Shift = 6 Then
    control_down = False
    tmp_str = SB.Panels(5)
    Mid(tmp_str, 1, 1) = " "
    SB.Panels(5) = tmp_str
  End If
  
  ' if shift/alt is pressed, assign numeric key to bookmark
  If Shift = 5 Then
    If KeyCode > 47 And KeyCode < 58 Then
      bookmarkstr(KeyCode - 48) = RTF.SelText
      SB.Panels(1) = "bookmark " & CStr(KeyCode - 48) & " set"
      KeyCode = 0
    End If
  End If
  
  If KeyCode = vbKeyShift Then
    shift_down = True
    tmp_str = SB.Panels(5)
    Mid(tmp_str, 2, 1) = "S"
    SB.Panels(5) = tmp_str
  End If
  
    ' hardcode ^A to select all
'  If (Shift = 2) And KeyCode = vbKeyA Then
'    RTF.SelStart = 0
'    RTF.SelLength = Len(RTF.Text)
'  End If

  If Shift = 2 And KeyCode = vbKeyF Then control_down = True

  If control_down And shift_down Then
    If KeyCode = vbKeyE Then ' top
      KeyCode = vbKeyHome
      force_shift_up
      resume_control_down
    ElseIf KeyCode = vbKeyD Then ' bottom
      KeyCode = vbKeyEnd
      resume_control_down
      force_shift_up
    ElseIf KeyCode = vbKeyH Then ' word left
      'If DOSfiletype Then ' use new word move
      '  text_reached = False
        'L = InStr(Mid(RTF.Text, lOrigin, lFinal - lOrigin + 1), "@") then
       
        'For L = RTF.SelStart To 2 Step -1
        '  If Asc(Mid(RTF.Text, L, 1)) <> 32 Then text_reached = True
        '  If (Asc(Mid(RTF.Text, L - 1, 1)) = 13) Then Exit For ' stop at CR
        '  If text_reached And (Asc(Mid(RTF.Text, L, 1)) = 32) And (Asc(Mid(RTF.Text, L + 1, 1)) <> 32) Then Exit For
        'Next
      '  If L >= 0 Then RTF.SelStart = L
      '  KeyCode = 0
      'Else  ' use windows builtin word shift
        force_shift_up
        resume_control_down
        KeyCode = vbKeyLeft
      'End If

    ElseIf KeyCode = vbKeyL Then ' word right
      If DOSfiletype Then
        goto_next_word
        KeyCode = 0
      Else
        force_shift_up
        resume_control_down
        KeyCode = vbKeyRight
      End If
    End If
  ElseIf control_down Then
   
    For i = 0 To numkeys - 1
      If KeyCode = key(i) Then
        Select Case keyfunc(i)
          Case 0: ' page up
            KeyCode = vbKeyPageUp
            force_control_up
            Exit For
          Case 1: ' page down
            KeyCode = vbKeyPageDown
            force_control_up
            Exit For
          Case 2: ' top
            KeyCode = vbKeyHome
            force_control_up
            Exit For
          Case 3: ' bottom
            KeyCode = vbKeyEnd
            force_control_up
            Exit For
          Case 4: ' insert
            KeyCode = vbKeyInsert
            force_control_up
            Exit For
          Case 5: ' delete
            KeyCode = vbKeyDelete
            force_control_up
            Exit For
          Case 6: ' delete line
            macro_function = 2
            KeyCode = 0
            force_control_up
            Timer1.Interval = 1
            Exit For
          Case 7:  ' insert line
            force_control_up
            Timer1.Interval = 1
            macro_function = 1
            If Asc(Mid(RTF.Text, RTF.SelStart + 1, 1)) <> 13 Then
              KeyCode = vbKeyEnd
            Else
              KeyCode = 0
            End If
            Exit For
          Case 8: ' find
            KeyCode = 0
            mnufind_Click
            Exit For
          Case 9: ' find next
            KeyCode = vbKeyF3
            Exit For
          Case 10: ' cursor up
            KeyCode = vbKeyUp
            force_control_up
            Exit For
          Case 11: ' cursor down
            KeyCode = vbKeyDown
            force_control_up
            Exit For
          Case 12: ' cursor left
            'KeyCode = vbKeyLeft
            'force_control_up
            RTF.SelStart = RTF.SelStart - 1
            KeyCode = 0
            Exit For
          Case 13: ' cursor right
            'KeyCode = vbKeyRight
            'force_control_up
            RTF.SelStart = RTF.SelStart + 1
            KeyCode = 0
            Exit For
          Case 14: ' screen down 1 line
            KeyCode = vbKeyUp
            SendMessageBynum RTF.hwnd, EM_SCROLL, SB_LINEUP, 0
            force_control_up
            Exit For
          Case 15: ' screen up 1 line
            force_control_up
            KeyCode = vbKeyDown
            SendMessageBynum RTF.hwnd, EM_SCROLL, SB_LINEDOWN, 0
            Exit For
          Case 16: ' open file
            KeyCode = 0
            control_down = False
            Call mnuFileOpen_Click
            Exit For
          Case 17: ' save file
            KeyCode = 0
            control_down = False
            Call SaveAFile
            Exit For

          Case 18: ' open URL
            ' release control key status
            control_down = False
            tmp_str = SB.Panels(5)
            Mid(tmp_str, 1, 1) = " "
            SB.Panels(5) = tmp_str

            url_there = False
            more_words_left = True
            While Not url_there And more_words_left
              url_there = test_url
              If Not url_there Then more_words_left = goto_next_word
            Wend
            
          Case 118: ' old open url code
            url_there = test_url
            
            If Not url_there Then
              orig_start = RTF.SelStart
              force_control_up
              Timer1.Interval = 1
              macro_function = 4
              KeyCode = vbKeyEnd
              Exit For
            End If

          Case 19: ' UNUSED! paste text below current line
            KeyCode = 0
            i = 0
            force_control_up
            newstart = RTF.SelStart
            SendKeys "{END}", True
            SendKeys "{ENTER}", True
            SendKeys "v", True
            KeyCode = 0
            resume_control_down
          
            Exit For

          Case 20: ' minimize
            Me.WindowState = 1

          Case Else: If keyfunc(i) <> -1 Then MsgBox "unmatched case - " + Str(keyfunc(i))
        End Select
      End If
    Next

    If i = numkeys Then KeyCode = 0
  End If

  If KeyCode = vbKeyF3 Then
    frmFind.globall.Value = 0
    Call frmFind.cmdFindNext_Click
  End If

  ' hardcode some ALT functions here
  If (Shift = 4) Then
  
    If KeyCode = 186 Then ' down
      SendMessageBynum rtf2.hwnd, WM_VSCROLL, SB_LINEDOWN, 0
    End If
  
    If KeyCode = vbKeyP Then
      SendMessageBynum rtf2.hwnd, WM_VSCROLL, SB_LINEUP, 0
    End If
  
  
    ' Alt-V: Paste under cursor
    If KeyCode = vbKeyV Then
      force_alt_up
      If RTF.SelStart = 0 Then RTF.SelStart = 2
      newstart = InStr(RTF.SelStart, RTF.Text, Chr(13), 0)
'      i = 0
'      newstart = RTF.SelStart
'      While Mid(RTF.Text, newstart, 1) <> vbCr And i < 100
'        i = i + 1
'        newstart = newstart + 1
'      Wend

      RTF.SelStart = newstart
      SendKeys "{ENTER}", True

      SendMessage RTF.hwnd, WM_PASTE, 0, 0
      resume_alt_down
      KeyCode = 0
    End If
    
    ' Alt-N to switch to Normal size
    If KeyCode = vbKeyN Then
      If Me.WindowState <> 0 Then Exit Sub
      mnuViewStatusbar_Click
      mnuMenubar_Click
'      Set_Default_Values
      Me.Left = 2190
      Me.Top = 2145
      Me.Width = 9480
      Me.Height = 6570

      KeyCode = 0
    End If

    ' Alt-D to divide pane
    If KeyCode = vbKeyD Then
       mnuSplit_Click
    End If
    
    ' paste into right pane
    If KeyCode = vbKeyZ Then
      rtf2.Text = ""
      SendMessage rtf2.hwnd, WM_PASTE, 0, 0
      RTF.SetFocus
      
    End If
    
    
    ' Alt-S to split pane
    If KeyCode = vbKeyS Then
        mnuSplit_Click
    End If
        
    ' Alt-R to REDUCE to small size
    If KeyCode = vbKeyF9 Then
      If Me.WindowState <> 0 Then Exit Sub
      mnuViewStatusbar_Click
      If mnuMenubar.Checked Then
        mnuMenubar_Click
      End If
      If Not ontop.Checked Then
        ontop_Click
      End If
      If mnuScrollbars.Checked Then
        mnuScrollbars_Click
      End If
      If mnuViewStatusbar.Checked Then
        mnuViewStatusbar_Click
      End If
      Me.Left = 7605
      Me.Top = 9585
      Me.Width = 4380
      Me.Height = 1215
      KeyCode = 0
    End If
    
    ' Alt-A toggles "always on top"
    If KeyCode = vbKeyA Then
      ontop_Click
      KeyCode = 0
    End If

    ' Alt-C for compact mode
    If KeyCode = vbKeyC Then
      mnuViewTitlebar_Click
    End If

    ' Alt-B for bottom mode
    If KeyCode = vbKeyB Then
      Me.Left = b_left '30
      Me.Top = b_top '8250
      Me.Width = b_width '15330
      Me.Height = b_height '2820
    End If


    ' Alt-M to toggle menubar
    If KeyCode = vbKeyM Then
      mnuMenubar_Click
      KeyCode = 0
    End If
    
    ' Alt-L to toggle horiz Scrollbar
    If KeyCode = vbKeyL Then
      mnuScrollbars_Click
      KeyCode = 0
    End If
    
    ' Alt-t to toggle statusbar
    If KeyCode = vbKeyT Then
      mnuViewStatusbar_Click
      KeyCode = 0
    End If
    
    ' Alt-H to hide/show textbox to enable moving
    If KeyCode = vbKeyH Then
      txtvisible = Not txtvisible
      RTF.Visible = txtvisible
    End If
    
    ' Alt-W toggles wordwrap
    If KeyCode = vbKeyW Then
      mnuFormatWordwrap_Click
      KeyCode = 0
    End If
    
    'Alt-Q: save and exit
    If KeyCode = vbKeyQ Then
      If FileChanged Then SaveAFile
      Unload Me
    End If
    
    ' alt-X to exit without changes
    If KeyCode = vbKeyX Then
      FileChanged = False
      Unload Me
    End If
    
    ' Alt numeric keys will jump to bookmark
    If KeyCode > 47 And KeyCode < 58 Then
      frmFind.findtext = bookmarkstr(KeyCode - 48)
      frmFind.remember.Value = 0
      frmFind.globall.Value = 1
      frmFind.cmdFindNext_Click
    End If
  End If

End Sub

Private Sub Set_Default_Values()
      Me.Left = 2190
      Me.Top = 2145
      Me.Width = 9480
      Me.Height = 6570
End Sub


Private Sub rtf2_KeyDown(KeyCode As Integer, Shift As Integer)
  ' ctrl-A to select all
  If (Shift = 2) And KeyCode = vbKeyA Then
    rtf2.SelStart = 0
    rtf2.SelLength = Len(rtf2.Text)
    Exit Sub
  End If

End Sub

Private Sub Timer1_Timer()
  Timer1.Interval = 0
  Select Case macro_function
    Case 1:
      SendKeys "{ENTER}"
    Case 2:
      If Asc(Mid(RTF.Text, RTF.SelStart, 1)) <> 10 Then SendKeys "{HOME}"
      If Asc(Mid(RTF.Text, RTF.SelStart, 1)) <> 13 Then SendKeys "+{END}"
      SendKeys "{DELETE}"
    Case 3:
      SendKeys "{ENTER}"
      SendKeys "^V"
    Case 4:
      Do While Mid(RTF.Text, RTF.SelStart, 1) = " " And RTF.SelStart > orig_start
        RTF.SelStart = RTF.SelStart - 1
      Loop
      test_url
  End Select
End Sub

Private Sub Form_Load()
  ' still dont know how to dynamically change icon visibility in taskbar
  onlyLoading = True
     
  SB.Panels(5) = "        "
  DOSfiletype = True
  If Len(Dir("c:\jkj")) > 0 Then ' activate debug  mode
    myCommand = "c:\jkj"
  Else
    myCommand = Command()
  End If
 
  If myCommand = "hidetaskbar" Then ' HIDE COMPLETE TASK BAR OF 2nd COMPUTER
      SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
      Me.Left = 0
      Me.Top = 0
      Me.Width = 19290
      Me.Height = 420
      Set Me.Icon = Nothing
  End If
 
  InitCmnDlg Me.hwnd
  cmndlg.flags = 5
  SB.Visible = mnuViewStatusbar.Checked
  RTF.RightMargin = IIf(mnuFormatWordwrap.Checked, 0, 200000)
  RTF.Text = " "
  RTF.SelStart = 0
  RTF.SelLength = 1
  SelFont
  RTF.Text = ""
  Set Undo = New clsUndo
  Undo.RichBox = RTF
  Undo.Reset

  If myCommand <> "" Then
    ChDir myCommand + "\.."
    current_file = FileOnly(myCommand)
    ' filter out quote at end of file
' kwrr commented this 12/15/10
    'If current_file <> "jkj" Then current_file = Mid(current_file, 1, Len(current_file) - 1)
    separator = " - "
  End If
  
  If myCommand = "hidetaskbar" Then
    Me.Caption = ""
  Else
    Me.Caption = current_file + separator + APP_NAME
  End If
  FileChanged = False
  mnuScrollbars.Checked = True
  txtvisible = True
  menuvisible = True
  Call read_ini
  onlyLoading = False ' I hope this doesnt cause problems but necessary asterisk since form_paint isnt called

  If myCommand <> "" Then load_new_file (myCommand)

End Sub

Private Sub Form_Paint()
  Dim keystate As Integer  ' state of the Q key
  
  keystate = GetKeyState(vbKeyControl)
  If keystate And &H8000 Then control_down = True
  
  If onlyLoading Then
    If myCommand <> "" Then
      'We've been shelled
      DoEvents
      NoStatusUpdate = True
      Screen.MousePointer = 11
      SB.Panels(1) = "Loading file...."
      LockWindowUpdate Me.hwnd
      myCommand = strUnQuoteString(myCommand) 'sometimes explorer uses quotes('send to' for example)
      myCommand = GetLongFilename(myCommand) 'looks better than a dos path
      Select Case LCase(ExtOnly(myCommand))
        Case "txt"
            RTF.SelText = OneGulp(myCommand) 'binary read
        Case Else
            RTF.SelText = OneGulp(myCommand) 'otherwise do binary read
      End Select
      Me.Caption = FileOnly(myCommand)
      SB.Panels(1) = myCommand
      SB.Panels(4) = GetFileSize(Len(RTF.Text)) 'show size of file
      RTF.Tag = myCommand
      If FileLen(myCommand) > 100000 Then
        'just using SelFont, RTF selection falls
        'over somewhere around 100k so do this
        'slightly less efficient but more reliable
        'method of font control
        RTF.SelStart = 0
        RTF.SelLength = Len(RTF.Text)
        SelFont
        RTF.SelLength = 0
      End If
      NoStatusUpdate = False
      EditEnable
      RTF.SelStart = 0
      Screen.MousePointer = 0
      LockWindowUpdate 0
      myCommand = ""
      If Not reminded Then Call daily_reminder
    End If
    
    Undo.Reset
    FileChanged = False
    onlyLoading = False
    set_filetype
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim Response As VbMsgBoxResult
  Dim blankfile As Boolean
  blankfile = False

  ' if file is blank (other than spaces, CR/LF's, then don't prompt to save
  For i = 1 To 100 ' only check 1st 100 chars
    If i >= Len(RTF.Text) Then
      blankfile = True
      Exit For
    End If
    
    If (Asc(Mid(RTF.Text, i, 1)) <> 32) And (Asc(Mid(RTF.Text, i, 1)) <> 13) And (Asc(Mid(RTF.Text, i, 1)) <> 10) Then
      blankfile = False
      Exit For
    End If
  Next
  
  If FileChanged And Not blankfile Then

    Response = MsgBox("Save changes ?", vbYesNoCancel)
    Select Case Response
        Case vbCancel
            Cancel = 1 'dont unload, user must want to do something else after all
        Case vbYes
            'SaveAFile returns false if user cancels during save process - dont unload
            If Not SaveAFile Then Cancel = 1
    End Select
  End If
    
  Call write_ini

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    'placing the RTF on a left aligned picturebox makes resizing easier
    PicLeft.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'unload correctly
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

Private Sub mnufind_Click()
  keyfindmode = False
  LockWindowUpdate Me.hwnd
  frmFind.Show , Me
  LockWindowUpdate 0
  Screen.MousePointer = 0
End Sub

Private Sub mnuFileAbout_Click()
    MsgBox APP_NAME + " " + VERSION
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileNew_Click()
  Dim Response As VbMsgBoxResult
  If FileChanged Then 'do we save current doc ?
      Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
      Select Case Response
          Case vbCancel
              Exit Sub
          Case vbYes
              If Not SaveAFile Then Exit Sub
      End Select
  End If
  'clear everything
  RTF.Text = ""
  Me.Caption = "Untitled.txt"
    
'  SB.Panels(1) = "File not saved"
  SB.Panels(4) = GetFileSize(2)
  RTF.Tag = ""
  SelFont 'maintain control of fonts
  FileChanged = False
End Sub

Private Sub mnuFileOpen_Click()
  load_new_file ("")
End Sub

Private Sub load_new_file(file_to_load As String)
  Dim Response As VbMsgBoxResult

  If FileChanged Then
      Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
      Select Case Response
          Case vbCancel
              Exit Sub
          Case vbYes
              If Not SaveAFile Then Exit Sub
      End Select
  End If
  
  If file_to_load = "" Then
    cmndlg.filefilter = "Plain text (*.txt)|*.txt|All files (*.*)|*.*"
    OpenFile
    If Len(cmndlg.filename) = 0 Then Exit Sub
    file_to_load = cmndlg.filename
  End If
'MsgBox "loading " & file_to_load
  changelog = ""
  SB.Panels(1) = "Loading file...."
  NoStatusUpdate = True
  Me.Refresh
  Screen.MousePointer = 11 'hourglass
  LockWindowUpdate Me.hwnd
  RTF.Text = ""
  SelFont
  RTF.SelText = OneGulp(file_to_load) 'binary read

'    If FileLen(cmndlg.filename) > 100000 Then
        'just using SelFont, RTF selection falls
        'over somewhere around 100k so do this
        'slightly less efficient but more reliable
        'method of font control
        RTF.SelStart = 0
        RTF.SelLength = Len(RTF.Text)
        SelFont
        RTF.SelLength = 0
'    End If
  current_file = file_to_load 'FileOnly(file_to_load)  ' kwrr I have to use full path here since I didnt get to making it work with fileonly
  separator = " - "
  Me.Caption = current_file 'cmndlg.filetitle
  SB.Panels(1) = file_to_load 'cmndlg.filename
  SB.Panels(4) = GetFileSize(Len(RTF.Text))
'MsgBox "loading...." & file_to_load
  RTF.Tag = file_to_load 'cmndlg.filename
  FileChanged = False 'reset need to save flag
  RTF.SelStart = 0
  NoStatusUpdate = False
  EditEnable
  LockWindowUpdate 0
  Screen.MousePointer = 0 'hourglass
End Sub

Private Sub mnuFilePageSetup_Click()
    ShowPageSetupDlg
End Sub

Private Sub mnuFilePrint_Click()
' doesnt select printer right YET.. If ShowPrinter Then RTF.SelPrint (Printer.hDC)
  RTF.SelPrint (Printer.hDC)
End Sub

Private Sub mnuFileProperties_Click()
  Dim temp As Variant, z As Long, count As Long, msg As String
  Dim charcnt As Long, linecnt As Long, Mcount As Long
  'otherwise just give document statistics
'  linecnt = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&) 'line count
'  charcnt = Len(RTF.Text) 'character count
'  temp = Split5(RTF.Text, Chr(32)) 'word count
'  For z = 0 To UBound(temp)
'    Select Case Trim(temp(z))
'      Case vbNullString
'      Case vbCrLf
'      Case vbCr
'      Case Else
'          Mcount = Mcount + 1
'    End Select
'  Next z
  
  Call set_filetype
  
  msg = "file type: " + IIf(DOSfiletype, "DOS", "UNIX") + vbCrLf
'  msg = msg + "Words :" + Format(Mcount, "#,###,###,##0") + vbCrLf
'  msg = msg + "Characters :" + Format(charcnt, "#,###,###,##0") + vbCrLf
'  msg = msg + "Lines :" + Format(linecnt, "#,###,###,##0") + vbCrLf + vbCrLf
  msg = msg + "Force Unix Mode?"
  
  i = MsgBox(msg, vbYesNo, APP_NAME)
  If i = 6 Then DOSfiletype = False ' 6=yes  7=no
  
End Sub
Private Sub mnuFileSave_Click()
  On Error Resume Next
  SaveAFile
End Sub
Private Sub mnuFileSaveAs_Click()
  Dim sfile As String
  On Error Resume Next

  With cmndlg
    .initdir = CurDir
    .filefilter = "All files (*.*)|*.*"
    .flags = 5 Or 2
    If Not SaveFile Then Exit Sub
    If Len(.filename) = 0 Then Exit Sub
    sfile = .filename
    Kill .filename
    FileSave RTF.Text, .filename 'plain text
    FileChanged = False 'reset flag
    'current_file = " - " + FileOnly(sfile)
    current_file = FileOnly(sfile)
    separator = " - "
    Me.Caption = current_file + separator + APP_NAME 'FileOnly(sfile)
    SB.Panels(1) = sfile
    SB.Panels(4) = GetFileSize(Len(RTF.Text))
    RTF.Tag = sfile
  End With
End Sub

Private Sub mnuFormatBackcolor_Click()
    'Dim col As Long 'new backcolor
    col = ShowColor
    If col <> -1 Then
        If col < 1 Then col = -col
        RTF.BackColor = col
    End If
End Sub


Private Sub mnuFormatFont_Click()
  On Error GoTo WOOPS
  Dim st As Long, curvl As Long, FontChange As Boolean
  'FileChanged is set to true by RTF change event (in the class)
  'As we are only changing font - not content, we dont want
  'this to alter due to this sub, so remember current state
  'so we can reset it below
  ChangeState = FileChanged
  Undo.IgnoreChange True  'dont add this action to the Undo buffer
  'current position
  curvl = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
  st = RTF.SelStart
  With SelectFont
      ShowFont
      'implement on our new font
      LockWindowUpdate Me.hwnd
      RTF.SelStart = 0
      RTF.SelLength = Len(RTF.Text)
      RTF.SelColor = .mFontColor
      RTF.SelFontName = .mFontName
      RTF.SelFontSize = .mFontsize
      RTF.SelBold = .mBold
      RTF.SelItalic = .mItalic
      RTF.SelStrikeThru = .mStrikethru
      RTF.SelUnderline = .mUnderline
      RTF.SelStart = st
      RTF.SelLength = 0
      SetScrollPos curvl, RTF 'reset to current scroll position
  End With
WOOPS:
  Undo.IgnoreChange False
  FileChanged = ChangeState 'reset to what it was
  Me.Caption = current_file + separator + APP_NAME
  RTF.SetFocus
  LockWindowUpdate 0
End Sub

Private Sub mnuFormatStats_Click()
    Dim temp As Variant, z As Long, count As Long, msg As String
    Dim charcnt As Long, linecnt As Long, Mcount As Long
    linecnt = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&) 'lines
    charcnt = Len(RTF.Text) 'characters
    temp = Split5(RTF.Text, Chr(32)) 'words
    For z = 0 To UBound(temp)
        Select Case Trim(temp(z))
            Case vbNullString
            Case vbCrLf
            Case vbCr
            Case Else
                Mcount = Mcount + 1
        End Select
    Next z
    msg = IIf(RTF.Tag = "", "File not yet saved.", RTF.Tag) + vbCrLf
    msg = msg + "Words :" + Format(Mcount, "#,###,###,##0") + vbCrLf
    msg = msg + "Characters :" + Format(charcnt, "#,###,###,##0") + vbCrLf
    msg = msg + "Lines :" + Format(linecnt, "#,###,###,##0")
    MsgBox msg, vbInformation, APP_NAME
End Sub

Private Sub mnuFormatWordwrap_Click()
    mnuFormatWordwrap.Checked = Not mnuFormatWordwrap.Checked
    SB.Panels(1) = "Wordwrap " & IIf(mnuFormatWordwrap.Checked, "On", "Off")
    RTF.RightMargin = IIf(mnuFormatWordwrap.Checked, 0, 200000)
    tmp_str = SB.Panels(5)
    Mid(tmp_str, 5, 1) = IIf(mnuFormatWordwrap.Checked, "W", " ")
    SB.Panels(5) = tmp_str
End Sub


Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    SB.Visible = mnuViewStatusbar.Checked
End Sub

Private Sub next_menu_Click()
 'frmFind.fromcur.Value = 1
 frmFind.globall.Value = 1
 Call frmFind.cmdFindNext_Click
End Sub

Private Sub PasteTimer_Timer()
    'you could hook the clipboard, but this will do
    'mnuFind.Enabled = Clipboard.GetFormat(vbCFText)
End Sub
Private Sub PicLeft_Resize()
    On Error Resume Next
'Debug.Print "kwrr 4 rtf1 width: " + Str(RTF.Width) + " rtf2: " + Str(RTF.Width)
  If rtf2.Visible Then
    rtf2.Visible = False
  End If
    
    ' kwrr delete commented below
'  If rtf2.Visible Then
'    RTF.Width = (PicLeft.Width - 120) / 2
'    RTF.Height = PicLeft.Height
'    rtf2.Width = (PicLeft.Width - 120) / 2
'    RTF.Height = PicLeft.Height
'  Else
    RTF.Width = PicLeft.Width - 120
    RTF.Height = PicLeft.Height
    rtf2.Width = PicLeft.Width - 120
    RTF.Height = PicLeft.Height
'  End If
'Debug.Print "kwrr 5 rtf1 width: " + Str(RTF.Width) + " rtf2: " + Str(RTF.Width)
End Sub

Private Sub RTF_Change()
  If compact_mode Then
    Exit Sub
  End If
  
  If searching Then
    FileChanged = False
    searching = False
  End If
 
  If Not onlyLoading Then
    FileChanged = True
    If Len(current_file) > 0 Then separator = "* - "
    Me.Caption = current_file + separator + APP_NAME
  End If

  If NoStatusUpdate Then Exit Sub
  SB.Panels(2) = RTF.SelStart
  SB.Panels(3) = RTF.SelLength
End Sub

Private Sub RTF_GotFocus()
  Dim z As Long 'allow tabs WITHIN richtextbox
  ReDim mTStop(0 To Controls.count - 1) As Boolean
  
  On Local Error Resume Next
  For z = 0 To Controls.count - 1
      mTStop(z) = Controls(z).TabStop
      Controls(z).TabStop = False
  Next
  SelFont
End Sub

Private Sub RTF_LostFocus()
  Dim z As Long 'reset tabstops to original state
  
  On Local Error Resume Next
  For z = 0 To Controls.count - 1
      Controls(z).TabStop = mTStop(z)
  Next
End Sub

Private Sub RTF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelFont
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
  'If Button = 2 Then PopupMenu mnuRightClick
'  Button = 1
End Sub

Private Sub RTF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'If Button = 2 Then PopupMenu mnuRightClick
  Button = 1
   
  ' trim trailing spaces from double clicking
  If Len(RTF.SelText) > 0 Then
    ' fixed bug here
    Do While Mid(RTF.SelText, Len(RTF.SelText), 1) = " " And RTF.SelLength > 1
      RTF.SelLength = RTF.SelLength - 1
    Loop
  End If
  
End Sub

Private Sub RTF_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Response As VbMsgBoxResult, temp As String
  If Data.GetFormat(vbCFFiles) Then
    If FileChanged Then 'do we save current doc ?
        Response = MsgBox("The current file has changed. Do you wish to save changes ?", vbYesNoCancel)
        Select Case Response
            Case vbCancel
                Effect = vbDropEffectNone
                Exit Sub
            Case vbYes
                If Not SaveAFile Then
                    Effect = vbDropEffectNone
                    Exit Sub
                End If
        End Select
    End If
    'Data.Files is a collection of the filepaths of files
    'dropped onto a control. Multiple files may be dropped
    'but in this app, we can only open one at a time
    'so we are only interested in Data.Files(1)
    temp = Data.Files(1)
    temp = strUnQuoteString(temp) 'sometimes explorer uses quotes('send to' for example)
    temp = GetLongFilename(temp) 'looks better than a dos path
    RTF.Text = ""
    SelFont

    Select Case LCase(ExtOnly(temp))
        Case "txt"
            RTF.SelText = OneGulp(temp) 'binary read
        Case Else
            RTF.SelText = OneGulp(temp) 'otherwise do binary read
    End Select
    Me.Caption = FileOnly(temp)
    SB.Panels(1) = temp
    SB.Panels(4) = GetFileSize(Len(RTF.Text)) 'show size of file
    RTF.Tag = temp
    Undo.Reset
    FileChanged = False
  Else
    Effect = vbDropEffectNone
  End If
End Sub

Private Sub RTF_SelChange()
    If NoStatusUpdate Then Exit Sub
    EditEnable
End Sub

Public Function SaveAFile() As Boolean
  Dim Response As VbMsgBoxResult, sfile As String
  On Error Resume Next 'DoSaveAs
  
  If current_file = "" Then
'MsgBox "its a new file"
    GoTo DoSaveAs 'must be a new file
  Else
'MsgBox "its an existing file; saving " & current_file & " after killing " & RTF.Tag
    Kill RTF.Tag
'MsgBox "file killed"
    FileSave RTF.Text, current_file
    FileChanged = False
    SaveAFile = True
    current_file = FileOnly(RTF.Tag)
    separator = " - "
    'current_file = " - " + FileOnly(RTF.Tag)
    frmMain.SB.Panels(1) = RTF.Tag + " saved at " + Str(Time)
  End If
  
  If Not compact_mode Then
    Me.Caption = current_file + separator + APP_NAME
  End If
'MsgBox "now to exit function"
  Exit Function
'MsgBox "going to dosaveas"
DoSaveAs:
  With cmndlg
    .filefilter = "Plain text (*.txt)|*.txt|All files (*.*)|*.*"
    .flags = 5 Or 2
    If SaveFile = False Then
        SaveAFile = False
        Exit Function
    End If
    'make sure we have the correct extension
    sfile = .filename

    Select Case .filefilterindex
      Case 1
        If InStr(1, sfile, ".") = 0 Then
            sfile = sfile + ".txt"
        Else
            sfile = ChangeExt(sfile, "txt")
        End If
        FileSave RTF.Text, sfile 'plain text
        current_file = " - " + FileOnly(sfile)
      Case 2
'              If InStr(1, sfile, ".") = 0 Then sfile = sfile + ".txt"
        FileSave RTF.Text, .filename 'plain text
'MsgBox "now file was saved: " & .filename
        'current_file = " - " + FileOnly(.filename)
        current_file = FileOnly(.filename)
        separator = " - "
    End Select
    Me.Caption = .filetitle
'    If mnuViewStatusbar.Checked Then
        SB.Panels(1) = "saved " + .filename + " at " + CStr(Time)
'    Else
'        MsgBox "saved " + .filename + " at " + CStr(Time)
'    End If
    SB.Panels(4) = GetFileSize(Len(RTF.Text))
    RTF.Tag = Mid(current_file, 4)   '.filename
    FileChanged = False 'reset flag
  End With
  SaveAFile = True
  'Me.Caption = APP_NAME + current_file
  Me.Caption = current_file + separator + APP_NAME
End Function

Public Sub SelFont()
End Sub

Public Sub EditEnable()
    'enable menus according to selection length
    SB.Panels(2) = RTF.SelStart
    SB.Panels(3) = RTF.SelLength
End Sub

Private Sub open_url(ByVal url As String)
  Dim filename As String, Dummy As String
  Dim BrowserExec As String * 255
  Dim retval As Long
  Dim FileNumber As Integer

  SB.Panels(1) = "Opening URL"

  ' First, create a known, temporary HTML file
  BrowserExec = Space(255)
  filename = "C:\temphtm.HTM"
  FileNumber = FreeFile                    ' Get unused file number
  Open filename For Output As #FileNumber  ' Create temp HTML file
      Write #FileNumber, "<HTML> <\HTML>"  ' Output text
  Close #FileNumber                        ' Close file
  
  ' Then find the application associated with it
  retval = FindExecutable(filename, Dummy, BrowserExec)
  BrowserExec = Trim(BrowserExec)
  ' If an application is found, launch it!
  If retval <= 32 Or IsEmpty(BrowserExec) Then ' Error
      MsgBox "Could not find associated Browser", vbExclamation, _
        "Browser Not Found"
  Else
      retval = ShellExecute(Me.hwnd, "open", BrowserExec, _
        url, Dummy, SW_SHOWNORMAL)
      If retval <= 32 Then        ' Error
          MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
      End If
  End If
  
  Kill filename                   ' delete temp HTML file
  SB.Panels(1) = ""

End Sub

Public Sub godown(ByVal lines As Integer)
  If lines > 0 Then
    For i = 0 To lines - 1
      SendMessageBynum RTF.hwnd, EM_SCROLL, SB_LINEDOWN, 0
    Next
  Else
      For i = 0 To Abs(lines + 1)
      SendMessageBynum RTF.hwnd, EM_SCROLL, SB_LINEUP, 0
    Next
  End If
End Sub

Sub write_ini()
  If myCommand = "hidetaskbar" Then Exit Sub
  
  Open App.Path + "\" + INI_FILE For Output As #1
  For i = 0 To MAX_SEARCHES
    Print #1, "find=" + frmFind.findhistory.List(i)
  Next
  
  For i = 0 To numkeys - 1
    If keyfunc(i) <> -1 Then Print #1, "key=" + CStr(key(i)) + "," + CStr(keyfunc(i))
  Next
  
  Print #1, "statusbar=" + CStr(Abs(CInt(mnuViewStatusbar.Checked)))
  Print #1, "filter_rtf=" + CStr(Abs(CInt(mnuFilterRTF.Checked)))

  If dont_save_settings Then
    Print #1, settings_string
  Else
    If mnuRememberPos.Checked Then
      Print #1, "wrap=" + CStr(Abs(CInt(mnuFormatWordwrap.Checked)))
      Print #1, "left=" + CStr(Me.Left)
      Print #1, "top=" + CStr(Me.Top)
      Print #1, "width=" + CStr(Me.Width)
      Print #1, "height=" + CStr(Me.Height)
    End If
  End If

  Print #1, "b_left=" + CStr(b_left)
  Print #1, "b_top=" + CStr(b_top)
  Print #1, "b_height=" + CStr(b_height)
  Print #1, "b_width=" + CStr(b_width)
  Print #1, "backcolor=16777215"
  Print #1, "fontname=" + SelectFont.mFontName
  Print #1, "fontsize=" + CStr(SelectFont.mFontsize)
  Print #1, "bold=" + CStr(CInt(SelectFont.mBold))
  Print #1, "italic=" + CStr(CInt(SelectFont.mItalic))
  Print #1, "color=" + CStr(SelectFont.mFontColor)
  Print #1, "permcount=" + frmFind.permcount.Text
  For i = 0 To 9
    If Len(bookmarkstr(i)) > 0 Then Print #1, "bookmark=" & CStr(i) & "|" & bookmarkstr(i)
  Next
  For i = 0 To num_savstrings
    Print #1, savstring(i)
  Next
  Close #1
End Sub


Sub read_ini()
  
  On Error Resume Next 'GoTo WOOPS
  If myCommand = "hidetaskbar" Then GoTo SKIPSTUFF
  perm_idx = MAX_SEARCHES + 1
  
  If Len(Dir$(App.Path + "\" + INI_FILE)) = 0 Then 'Exit Sub
    frmFind.Hide
    
    With SelectFont

      Me.Left = 2190
      Me.Top = 2145
      Me.Width = 9480
      Me.Height = 6570
      RTF.BackColor = 16777215
      .mFontColor = 0
      SB.Visible = 1
        
      .mFontName = "Lucida Console"
      .mFontsize = 10
  
      RTF.SelStart = 0
      RTF.SelLength = Len(RTF.Text)
      RTF.SelColor = .mFontColor
      RTF.SelFontName = .mFontName
      RTF.SelFontSize = .mFontsize
      RTF.SelBold = .mBold
      RTF.SelItalic = .mItalic
      RTF.SelStart = 0
      RTF.SelLength = 0
  
    End With

    RTF.Visible = True
    frmMain.Visible = True
    frmMain.Show
    
    Exit Sub
  Else
        Open App.Path + "\" + INI_FILE For Input As #1 'not curdir
  End If
  

  With SelectFont
  frmFind.Hide

  Do While Not EOF(1)
    Line Input #1, strr
    If Mid(strr, 1, 1) = " " Or Len(strr) = 0 Or Mid(strr, 1, 1) = "#" Then
    Else
      param = Split5(strr, "#") ' filter out end-of-line comments
      param = Split5(param(0), "=")
      If param(0) = "key" Then
        param = Split5(param(1), ",")
        key(numkeys) = Val(param(0))
        keyfunc(Inc(numkeys)) = Val(param(1))
      ElseIf param(0) = "find" Then
        frmFind.findhistory.AddItem param(1)
      ElseIf param(0) = "b_left" Then
        b_left = CInt(param(1))
      ElseIf param(0) = "b_top" Then
        b_top = CInt(param(1))
      ElseIf param(0) = "b_height" Then
        b_height = CInt(param(1))
      ElseIf param(0) = "b_width" Then
        b_width = CInt(param(1))
      ElseIf param(0) = "left" Then
        If (param(1) < 0) Then param(1) = 100
        Me.Left = CInt(param(1))
      ElseIf param(0) = "top" Then
        If (param(1) < 0) Then param(1) = 100
        Me.Top = CInt(param(1))
      ElseIf param(0) = "width" Then
        If (param(1) < 0) Then param(1) = 100
        Me.Width = CInt(param(1))
      ElseIf param(0) = "height" Then
        If (param(1) < 0) Then param(1) = 100
        Me.Height = CInt(param(1))
      ElseIf param(0) = "wrap" Then
        If param(1) = "1" Then mnuFormatWordwrap_Click
      ElseIf param(0) = "filter_rtf" Then
        If param(1) = "1" Then mnuFilterRTF_Click
      ElseIf param(0) = "statusbar" Then
        mnuViewStatusbar.Checked = CBool(param(1))
        SB.Visible = mnuViewStatusbar.Checked
      ElseIf param(0) = "fontname" Then
        .mFontName = param(1)
      ElseIf param(0) = "fontsize" Then
        .mFontsize = CStr(param(1))
      ElseIf param(0) = "bold" Then
        .mBold = param(1)
      ElseIf param(0) = "italic" Then
        .mItalic = param(1)
      ElseIf param(0) = "color" Then
        .mFontColor = param(1)
      ElseIf param(0) = "permcount" Then
        frmFind.permcount.Text = param(1)
      ElseIf param(0) = "backcolor" Then
        col = CDbl(param(1))
        If col < 1 Then col = -col
        RTF.BackColor = col
      ElseIf param(0) = "bookmark" Then
        ' format: bookmark=3|string
        param = Split5(param(1), "|")
        i = CInt(param(0))
        bookmarkstr(i) = param(1)
'      ElseIf param(0) = "icon" Then ' DELETE THIS WHEN DEFAULT IS WORKING
        ' format: icon=z.cxg|c:\z.ico (use FULL path of icon file!)
        ' savstring is array used to save string that will be rewritten to ini in write_ini
 '       savstring(Inc(num_savstrings)) = "icon=" & param(1)
 '       param = Split5(param(1), "|")
 '       If param(0) = current_file Then
 '         Set Me.Icon = LoadPicture(param(1))
 '       End If
      ElseIf param(0) = "default" Then
        ' format: default=z.cxg|wordwrap|x|y|width|height|iconfile (use FULL path of icon file!)
        '            param   0   1       2 3  4     5        6
        ' savstring is array used to save string that will be rewritten to ini in write_ini
        savstring(Inc(num_savstrings)) = "default=" & param(1)
        param = Split5(param(1), "|")
        If param(0) = current_file Then
          settings_string = "wrap=0" & vbCrLf & _
            "left=" + CStr(Me.Left) & vbCrLf & _
            "top=" & CStr(Me.Top) & vbCrLf & _
            "height=" & CStr(Me.Height) & vbCrLf & _
            "width=" & CStr(Me.Width)

          dont_save_settings = True
          Set Me.Icon = LoadPicture(param(6))
          If param(1) = "1" Then mnuFormatWordwrap_Click
          Me.Left = CInt(param(2))
          Me.Top = CInt(param(3))
          Me.Width = CInt(param(4))
          Me.Height = CInt(param(5))
          For i = 7 To UBound(param)
            If i = 7 Then mnuFiles(0).Visible = True
            Load mnuFiles(i - 6)
            mnuFiles(i - 6).Caption = param(i)
          Next
        End If
      End If
    End If
  Loop
  
  RTF.SelStart = 0
  RTF.SelLength = Len(RTF.Text)
  RTF.SelColor = .mFontColor
  RTF.SelFontName = .mFontName
  RTF.SelFontSize = .mFontsize
  RTF.SelBold = .mBold
  RTF.SelItalic = .mItalic
  RTF.SelStart = 0
  RTF.SelLength = 0
  
  End With

SKIPSTUFF:
  RTF.Visible = True
  frmMain.Visible = True
  
NOSEARCHFILE:
  Close #1
  Exit Sub
  
WOOPS:
 MsgBox "invalid icon file: " & param(1)
End Sub



