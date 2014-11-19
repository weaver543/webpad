VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CheckBox stay 
      Caption         =   "Stay"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton addtoperm 
      Caption         =   "Add"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox permcount 
      Height          =   285
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "4"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox globall 
      Caption         =   "&Global"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox ChWord 
      Caption         =   "&Whole word"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox ChCase 
      Caption         =   "Match &Case"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox remember 
      Caption         =   "&Remember"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ListBox findhistory 
      DragIcon        =   "frmFind.frx":0000
      Height          =   2985
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox findtext 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sticky"
      Height          =   975
      Left            =   1680
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
      Begin VB.Label Label1 
         Caption         =   "Count"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim updown_search As Boolean
Dim keyfindmode As Boolean, cancelkey As Boolean
Dim st As Long
Dim mmatchCase As Integer
Dim mWholeword As Integer
Dim Found As Long
Dim vStrPos As Long
Dim StopNow As Boolean
Dim searching As Boolean
Dim savtext As String
Dim control_spaced As Boolean

Private Sub cboFind_Change()
  If Len(Trim(cboFind.Text)) = 0 Then
      cmdFindNext.Enabled = False
  Else
      cmdFindNext.Enabled = True
  End If
  st = 0
End Sub

Private Sub cboFind_Click()
 ' MsgBox cboFind.ListIndex
  'MsgBox cboFind.Text
  globall.Value = 1
'  fromcur.Value = 0
  Call cmdFindNext_Click
  frmFind.Hide
End Sub

Private Sub addtoperm_Click()
  If findhistory.ListIndex = -1 Then Exit Sub
  tmp_str = findhistory.List(findhistory.ListIndex)
  For i = findhistory.ListIndex + 1 To MAX_SEARCHES
    findhistory.List(i - 1) = findhistory.List(i)
  Next
  findhistory.List(MAX_SEARCHES) = tmp_str
End Sub

Private Sub cmdCancel_Click()

    If searching Then
        StopNow = True
        searching = False
    Else
        'unload Me
        Me.Hide
        frmMain.EditEnable
    End If
End Sub

Public Sub cmdFindNext_Click()
  Dim CurrentLine As Long, topidx As Integer
  If Not FileChanged Then frmMain.searching = True
  
  frmMain.SB.Panels(1) = ""

  If newsearch Or (Not newsearch And globall.Value = 1) Then
    st = 0
  Else
    st = frmMain.RTF.SelStart + 1
  End If

  If ChWord.Value = 1 Then
      mWholeword = 2
  Else
      mWholeword = 0
  End If
  
  If ChCase.Value = 1 Then
      mmatchCase = 4
  Else
      mmatchCase = 0
  End If
  
  vStrPos = SendMessageByString&(findtext.hwnd, CB_FINDSTRINGEXACT, 0, findtext.Text)
  
  Found = frmMain.RTF.Find(findtext.Text, st, , mWholeword Or mmatchCase)

  If Found <> -1 Then
    half_screen = (frmMain.RTF.Height \ 186) \ 2

    st = Found + Len(findtext.Text)
    frmMain.RTF.SetFocus
    
     ' scroll found string to center
    half_screen = (frmMain.RTF.Height \ 186) \ 2
    frmMain.onlyLoading = True
    '(highlighting cause problems) HighLightSelection frmMain, frmMain.RTF, vbYellow 'GetHLColor(cboHLColor)
    frmMain.onlyLoading = False
    
    first_visible = SendMessage(frmMain.RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)

    move_amt = frmMain.RTF.GetLineFromChar(frmMain.RTF.SelStart) - first_visible - half_screen
   
    frmMain.godown move_amt
    
    If remember.Value = 1 And frmFind.globall.Value = 1 And Not control_spaced Then
      If findhistory.ListCount > MAX_SEARCHES - Val(permcount) Then
        topidx = MAX_SEARCHES - Val(permcount)
      Else
        topidx = findhistory.ListCount
      End If
      
      ' move history list down to make room to add to top
      For i = topidx To 1 Step -1
        findhistory.List(i) = findhistory.List(i - 1)
      Next
      
      findhistory.List(0) = findtext.Text
    End If
    
    If stay.Value <> 1 Then frmFind.Hide
    'frmMain.SB.Panels(4) = "'" + findtext.Text + "' found."
    frmMain.SB.Panels(4) = Len(findtext.Text)
  Else
    
    st = 0
    frmMain.SB.Panels(1) = "'" + findtext.Text + "' not found."
   findtext.SelStart = 0
    findtext.SelLength = Len(findtext.Text)
    
  End If
  
  'LockWindowUpdate 0
  
  If newsearch Then newsearch = False
End Sub


Private Sub cmdSelection_Click(Index As Integer)
    'get selection from frmMain.RTF
    If Index = 0 Then
        cboFind.Text = frmMain.RTF.SelText
    Else
        cboReplace.Text = frmMain.RTF.SelText
    End If
End Sub


Private Sub findhistory_Click()
  If stay.Value = 1 Then
    findhistory_DblClick
  Else
    findtext.Text = findhistory.Text
  End If
End Sub

Private Sub findhistory_DblClick()
  findtext.Text = findhistory.Text
  globall.Value = 1
  remember.Value = 0
  Call cmdFindNext_Click
End Sub

Private Sub findtext_Click()
  findtext.SelStart = 0
  findtext.SelLength = Len(findtext.Text)
End Sub

Private Sub findtext_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Or KeyAscii = 6 Or KeyAscii = 18 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Deactivate()
  If control_spaced Then findtext.Text = savtext
  control_spaced = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift <> 2 Then ' if control not down
    keyfindmode = False
  Else
    ' all other than 3 keys here will cancel keyfind mode
    If KeyCode <> vbKeyD And KeyCode <> vbKeyS And KeyCode <> vbKeyF And KeyCode <> vbKeyR And KeyCode <> vbKeySpace And KeyCode <> vbKeyReturn And KeyCode <> vbKeySpace Then keyfindmode = False
  End If

  If Shift = 2 Then ' if control down
    If KeyCode = vbKeyReturn Then ' prepend [ char
      findtext.SelStart = 0
      findtext.SelLength = 0
      findtext.SelText = "["
      cmdFindNext_Click
      cancelkey = True
      control_down = True
    Else
      If KeyCode = vbKeyF Then
        keyfindmode = True
        If findhistory.ListIndex = MAX_SEARCHES Then
          findhistory.ListIndex = 0
        Else
          findhistory.ListIndex = findhistory.ListIndex + 1
        End If
      ElseIf KeyCode = vbKeyR Or KeyCode = vbKeyS Or KeyCode = vbKeyD Then
        If findhistory.ListIndex = -1 Then findhistory.ListIndex = 0
        keyfindmode = True
        findhistory.ListIndex = findhistory.ListIndex - 1
        If findhistory.ListIndex = -1 Then findhistory.ListIndex = MAX_SEARCHES
      ElseIf KeyCode = vbKeySpace Then
        KeyCode = vbKeyI
        savtext = findtext.Text
        control_spaced = True ' to prevent space from inserting into text
        findtext.Text = findhistory.Text
        cmdFindNext_Click
      End If
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If cancelkey Then
    KeyAscii = 0
    keyfindmode = False
    cancelkey = False
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 
  ' letting up on control when keyfindmode will activate search
  If keyfindmode And KeyCode = vbKeyControl Then
    remember.Value = 0
    cmdFindNext_Click
    remember.Value = 1
  End If
  
  If KeyCode = vbKeyControl Then control_down = False
  If KeyCode = vbKeyShift Then shift_down = False
  If KeyCode = vbKeyEscape Then Me.Hide
  If KeyCode = vbKeyReturn Then cmdFindNext_Click
  
End Sub

Private Sub Form_Load()
  Dim v As Integer, temp As String
  newsearch = True

'  frmFind.Show
'  findtext.SetFocus
End Sub

Private Sub Form_Paint()

  If frmMain.ontop.Checked Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

  keyfindmode = True
   
  remember.Value = 1
  findtext.Text = ""
  findtext.SelStart = 1
  findtext.SetFocus
  globall.Value = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
  Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  frmMain.EditEnable
End Sub

Private Sub permcount_Change()
  If Not permcount.Text Like "#" Then
    MsgBox "Number from 0-9 required"
    permcount.Text = "0"
  End If
End Sub

Private Sub permcount_Click()
  permcount.SelStart = 0
  permcount.SelLength = 1
End Sub

