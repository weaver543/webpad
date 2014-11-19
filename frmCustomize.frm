VERSION 5.00
Begin VB.Form frmCustomize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customize Shortcuts"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
   Icon            =   "frmCustomize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cancelbtn 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton dummybox 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton okbtn 
      Caption         =   "OK"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton clearbtn 
      Caption         =   "Delete"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "pressing ANY key will assign that key to selected function"
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "frmCustomize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private assign_key As Boolean, save_ini As Boolean

Private Sub cancelbtn_Click()
  save_ini = False
  Unload Me
End Sub

Private Sub clearbtn_Click()
  Text1.Text = ""
  num_found = 0
  previously_assigned = -2
  For i = numkeys - 1 To 0 Step -1
    If keyfunc(i) = List1.ListIndex Then
      If previously_assigned = -2 Then
        previously_assigned = -1
        keyfunc(i) = -1  ' delete it\
      ElseIf previously_assigned = -1 Then
        previously_assigned = key(i)
      End If
    End If
  Next

  For i = 0 To numkeys - 1
    If keyfunc(i) = List1.ListIndex Then num_found = num_found + 1
  Next
      
  If num_found = 0 Then clearbtn.Enabled = False
  Label1.Caption = CStr(num_found) + " keys assigned"
  show_keycode (previously_assigned)
End Sub

Private Sub okbtn_Click()
  save_ini = True
  Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 17 Or KeyCode = 18 Then Exit Sub

  If List1.ListIndex = -1 Then
    MsgBox "click function on left then press key that will be assigned to it when control is pressed"
    Exit Sub
  End If
  
  num_found = 0
  
  If numkeys > MAX_KEYS - 1 Then
    MsgBox "Max number of keys already assigned"
    Exit Sub
  End If

  ' clear other functions that have same key assigned
  For i = 0 To numkeys - 1
    If key(i) = KeyCode Then
      keyfunc(i) = -1
    End If
  Next
    
  key(numkeys) = KeyCode
  keyfunc(numkeys) = List1.ListIndex
  numkeys = numkeys + 1
  assign_key = False
  
  For i = 0 To numkeys - 1
    If keyfunc(i) = List1.ListIndex Then num_found = num_found + 1
  Next
      
  clearbtn.Enabled = True
  Label1.Caption = CStr(num_found) + " keys assigned"
  
  show_keycode (KeyCode)
End Sub

Private Sub show_keycode(KeyCode As Integer)
  
  If KeyCode < 1 Then Exit Sub
  
  Select Case KeyCode
  Case 20
    tmp_str = "CAPS"
  Case 32
    tmp_str = "SPACE"
  Case 192
    tmp_str = "`"
  Case 112
    tmp_str = "F1"
  Case 113
    tmp_str = "F2"
  Case 114
    tmp_str = "F3"
  Case 115
    tmp_str = "F4"
  Case 116
    tmp_str = "F5"
  Case 117
    tmp_str = "F6"
  Case 118
    tmp_str = "F7"
  Case 119
    tmp_str = "F8"
  Case 120
    tmp_str = "F9"
  Case 121
    tmp_str = "F10"
  Case 122
    tmp_str = "F11"
  Case 123
    tmp_str = "F12"
  Case 220
    tmp_str = "\"
  Case 219
    tmp_str = "["
  Case 221
    tmp_str = "]"
  Case 187
    tmp_str = "="
  Case 188
    tmp_str = ","
  Case 190
    tmp_str = "."
  Case 189
    tmp_str = "-"
  Case 191
    tmp_str = "/"
  Case 220
    tmp_str = "\"
  Case Else
    tmp_str = Chr(KeyCode)
  End Select
  Text1.Text = tmp_str

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  assign_key = False
End Sub

Private Sub Form_Load()
  
functionlist = Array("0=Page up", "1=Page down", "2=Top", "3=Bottom", "4=Insert", "5=Delete", _
  "6=Delete line", "7=Insert Line", "8=Find", "9=Find Next", "10=Cursor up", "11=Cursor down", _
  "12=Cursor left", "13=Cursor right", "14=Screen down", "15=Screen up", _
  "16=Open File", "17=Save", "18=Open URL under cursor", "19=Paste under cursor", _
  "20=Minimize")
  
  For i = 0 To UBound(functionlist)
    param = Split5(functionlist(i), "=")
   
    List1.AddItem param(1)
  Next

End Sub

Private Sub Form_Paint()
  dummybox.SetFocus
End Sub

Private Sub List1_Click()
  
  num_found = 0
  dummybox.SetFocus
  Text1.Text = ""
  
  For i = 0 To numkeys - 1
    If keyfunc(i) = List1.ListIndex Then
      show_keycode (key(i))
      num_found = num_found + 1
    End If
  Next
  
  If num_found = 0 Then
    clearbtn.Enabled = False
  Else
    clearbtn.Enabled = True
  End If
  
  Label1.Caption = CStr(num_found) + " keys assigned"
  
End Sub
