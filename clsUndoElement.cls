VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUndoElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

  
'Variables to hold Undo/Redo data
Public SelStart As Long
Public SelUndoLength As Long
Public SelRedoLength As Long
Public SelUndoText As String
Public SelRedoText As String
Public UndoTabPosition As Long
Public RedoTabPosition As Long
Public Action As Boolean

