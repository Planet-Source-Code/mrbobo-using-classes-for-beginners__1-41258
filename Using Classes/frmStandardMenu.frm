VERSION 5.00
Begin VB.Form frmStandardMenu 
   Caption         =   "Notes"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6675
   Icon            =   "frmStandardMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.TextBox Text2 
         Height          =   1695
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   1335
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsWordWrap 
         Caption         =   "WordWrap"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolsDate 
         Caption         =   "Time/Date"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuToolsSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsFind 
         Caption         =   "&Find and Replace"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuToolsFont 
         Caption         =   "Font"
      End
   End
End
Attribute VB_Name = "frmStandardMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - December 2002
'If you got it elsewhere - they stole it from PSC.

'Please visit our website at www.psst.com.au

'By using a class you can dramatically reduce the amount of code
'required in your forms. This makes coding more complex apps much
'easier and clearer.

'Whilst this project is intended to show the use of class modules,
'it is a close replica of Notepad and shows some text/file handling
'methods which may come in handy for beginners. It is not intended
'however to be a full blown text editor.

'To the more advanced coders : I have tried to make the
'class as simple and easy to understand as possible. The task of
'a text editor is so simple that it barely requires a class and all
'the code could just as easliy be incorporated on this form. This
'is intended to show beginners how to incorporate classes into
'thier apps and a text editor provides a simple platform for
'that demonstration.

Option Explicit
Public ClTxt As ClPlainText
Private Sub Form_Load()
    Set ClTxt = New ClPlainText
    With ClTxt
        .DisplayFileNameInTitleBar(Me) = True
        .PlainTextBox = Text1
        .PopUpEditMenu = mnuEdit
        .NewFile
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ClTxt.CheckSave Then Cancel = 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PicLeft.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

Private Sub mnuEdit_Click()
    'Enable menu items appropriately
    mnuEditCopy.Enabled = ClTxt.PlainTextBox.SelLength > 0
    mnuEditCut.Enabled = mnuEditCopy.Enabled
    mnuEditDelete.Enabled = mnuEditCopy.Enabled
    mnuEditUndo.Enabled = ClTxt.CanUndo
    mnuEditPaste.Enabled = Clipboard.GetFormat(vbCFText)
End Sub

Private Sub mnuEditCopy_Click()
    ClTxt.Copy
End Sub

Private Sub mnuEditCut_Click()
    ClTxt.Cut
End Sub

Private Sub mnuEditDelete_Click()
    ClTxt.Delete
End Sub

Private Sub mnuEditPaste_Click()
    ClTxt.Paste
End Sub

Private Sub mnuEditSelectAll_Click()
    ClTxt.SelectAll
End Sub

Private Sub mnuEditUndo_Click()
    ClTxt.Undo
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    ClTxt.NewFile
End Sub

Private Sub mnuFileOpen_Click()
    ClTxt.LoadFile
End Sub

Private Sub mnuFileSave_Click()
    ClTxt.SaveFile
End Sub

Private Sub mnuFileSaveAs_Click()
    ClTxt.SaveAsFile
End Sub

Private Sub mnuTools_Click()
    mnuToolsFind.Enabled = Len(ClTxt.PlainTextBox.Text) > 0

End Sub

Private Sub mnuToolsDate_Click()
    ClTxt.PlainTextBox.SelLength = 0
    ClTxt.PlainTextBox.SelText = Now
End Sub

Private Sub mnuToolsFind_Click()
    frmFindReplace.Show , Me
End Sub

Private Sub mnuToolsFont_Click()
    ClTxt.ShowFont
End Sub

Private Sub mnuToolsWordWrap_Click()
    Dim temp As String, ChangeState As Boolean, CurStart As Long
    mnuToolsWordWrap.Checked = Not mnuToolsWordWrap.Checked
    'Remember old conditions
    ChangeState = ClTxt.FileChanged
    CurStart = ClTxt.PlainTextBox.SelStart
    temp = ClTxt.PlainTextBox.Text
    'Make sure all textboxes have the same Font/Forecolor
    Set Text1.Font = ClTxt.PlainTextBox.Font
    Set Text2.Font = ClTxt.PlainTextBox.Font
    Text1.ForeColor = ClTxt.PlainTextBox.ForeColor
    Text2.ForeColor = ClTxt.PlainTextBox.ForeColor
    'Swap textboxes
    ClTxt.PlainTextBox = IIf(mnuToolsWordWrap.Checked, Text1, Text2)
    'Update new textbox contents/conditions
    ClTxt.PlainTextBox.Text = temp
    ClTxt.PlainTextBox.SelStart = CurStart
    ClTxt.FileChanged = ChangeState
    'Show the correct textbox
    Text1.Visible = mnuToolsWordWrap.Checked
    Text2.Visible = Not Text1.Visible
    ClTxt.PlainTextBox.SetFocus
End Sub

Private Sub PicLeft_Resize()
    On Error Resume Next
    Text1.Move 0, 0, PicLeft.ScaleWidth, PicLeft.ScaleHeight
    Text2.Move 0, 0, PicLeft.ScaleWidth, PicLeft.ScaleHeight
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Ignore shortcut key strokes - the class will deal with them
    Select Case Shift
        Case 2
            Select Case KeyCode
                Case vbKeyZ, vbKeyDelete, vbKeyX, vbKeyC, vbKeyV, vbKeyA
                    KeyCode = 0
            End Select
        Case 3
            Select Case KeyCode
                Case 45: KeyCode = 0
            End Select
    End Select
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    'Ignore shortcut key strokes - the class will deal with them
    Select Case Shift
        Case 2
            Select Case KeyCode
                Case vbKeyZ, vbKeyDelete, vbKeyX, vbKeyC, vbKeyV, vbKeyA
                    KeyCode = 0
            End Select
        Case 3
            Select Case KeyCode
                Case 45: KeyCode = 0
            End Select
    End Select

End Sub
