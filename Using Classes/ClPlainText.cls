VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClPlainText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - December 2002
'If you got it elsewhere - they stole it from PSC.

'Please visit our website at www.psst.com.au

'Correctly written classes makes your code...
'- re-usable
'- faster
'- reliable
'- memory efficient
'They also...
'- speed devolopment time
'- make debugging easier
'- can be placed in DLL's making them accessable from
'  multiple applications simultaneously
'
'As time goes by they become even better
'because each time you use the same class
'in different apps, any bugs that do show
'up get removed for future apps as well

'By adding a class to your "Templates" directory
'in Visual Studio, they are easily added to a new
'project. You might notice the name of the main
'form in this demo is "frmStandardMenu". This is
'one of my templates that is an empty form
'with the "File" menu already built. Many projects
'will require a "File" menu so why rebuild it each time?

Option Explicit
'Font dialog declarations
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 31
End Type
Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_ROMAN = 16
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const REGULAR_FONTTYPE = &H400
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

'Open/Save dialog declarations
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'Edit functions
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLEAR = &H303
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_PASTE = &H302
Private Const EM_SETSEL = &HB1
Private Const WM_UNDO = &H304
Private Const EM_CANUNDO = &HC6
'Variables used
Private WithEvents m_PlainTextBox As TextBox
Attribute m_PlainTextBox.VB_VarHelpID = -1
Private m_PopUpEditMenu As Menu
Private m_DisplayFileNameInTitleBar As Boolean
Private m_TitleBarForm As Form
Private m_CurrentFile As String
Public FileChanged As Boolean

'Properties - trying to keep the number of Properties to
'a minimum reduces memory usage

'Notice how each property has a corresponding vaiable of the
'same name prefixed by "m_". When referring to a property from
'within the class use the variable, whilst from a form use the
'property name itself. If the property refers to an object
'(in this demo a Form, Textbox or Menu) you need to use
'"Set" when applying the property to the variable and vice versa

'Properties are the heart of classes - some subs/functions
'could in fact be in separate modules and sometimes this
'is preferable

'It is useful to note that Usercontrols behave in very similar ways

Public Property Get DisplayFileNameInTitleBar(TitleBarForm As Form) As Boolean
    DisplayFileNameInTitleBar = m_DisplayFileNameInTitleBar
End Property
Public Property Let DisplayFileNameInTitleBar(TitleBarForm As Form, ByVal vNewValue As Boolean)
    'Two properties in one!
    Set m_TitleBarForm = TitleBarForm
    m_DisplayFileNameInTitleBar = vNewValue
End Property
Public Property Get CurrentFile() As String
    CurrentFile = m_CurrentFile
End Property

Public Property Let CurrentFile(ByVal vNewValue As String)
    m_CurrentFile = vNewValue
End Property
Public Property Get PopUpEditMenu() As Menu
    Set PopUpEditMenu = m_PopUpEditMenu
End Property

Public Property Let PopUpEditMenu(ByVal vNewValue As Menu)
    Set m_PopUpEditMenu = vNewValue
End Property
Public Property Get PlainTextBox() As TextBox
    Set PlainTextBox = m_PlainTextBox
End Property
Public Property Let PlainTextBox(ByVal vNewValue As TextBox)
    Set m_PlainTextBox = vNewValue
End Property
'Edit subs - using the API is very reliable and fast
Public Sub Cut()
    SendMessage m_PlainTextBox.hWnd, WM_CUT, 0, 0
    m_PlainTextBox.SetFocus
End Sub

Public Sub Copy()
    SendMessage m_PlainTextBox.hWnd, WM_COPY, 0, 0
    m_PlainTextBox.SetFocus
End Sub

Public Sub Paste()
    SendMessage m_PlainTextBox.hWnd, WM_PASTE, 0, 0
    m_PlainTextBox.SetFocus
End Sub

Public Sub Delete()
    SendMessage m_PlainTextBox.hWnd, WM_CLEAR, 0, 0
    m_PlainTextBox.SetFocus
End Sub

Public Sub Undo()
    SendMessage m_PlainTextBox.hWnd, WM_UNDO, 0, 0
    m_PlainTextBox.SetFocus
End Sub

Public Sub SelectAll()
    SendMessage m_PlainTextBox.hWnd, EM_SETSEL, 0, Len(m_PlainTextBox.Text)
    m_PlainTextBox.SetFocus
End Sub
Public Function CanUndo() As Boolean
    CanUndo = SendMessage(m_PlainTextBox.hWnd, EM_CANUNDO, 0, 0)
End Function

'Textbox events

'Notice that in the declaration section we declared...
'"Private WithEvents m_PlainTextBox As TextBox"
'Look in the left hand combobox in the IDE and you will see
'"m_PlainTextBox". Click on it. Now in the right hand combobox
'you will see all the events associated with a standard textbox
'just as if this class was a form. Once you have set the property
'"PlainTextBox" to a textbox control on a form, all the events
'for that textbox will be fired here in the class as well.
'This way the class can easily respond to those events without
'the need for any extra code


Private Sub m_PlainTextBox_Change()
    FileChanged = True
End Sub

Private Sub m_PlainTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    'We need to intercept these keys otherwise
    'the textbox may do an action twice
    If Shift = 2 And KeyCode = vbKeyZ Then
        Undo
        KeyCode = 0
    ElseIf KeyCode = vbKeyDelete Then
        Delete
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyX Then
        Cut
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyC Then
        Copy
        KeyCode = 0
    ElseIf (Shift = 2 And KeyCode = vbKeyV) Or (Shift = 3 And KeyCode = 45) Then
        Paste
        KeyCode = 0
    ElseIf Shift = 2 And KeyCode = vbKeyA Then
        SelectAll
        KeyCode = 0
    End If
End Sub

Private Sub m_PlainTextBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if we are using a custom Popup menu
    If Button = 2 Then
        If Not m_PopUpEditMenu Is Nothing Then
            m_PlainTextBox.Enabled = False
            m_PopUpEditMenu.Parent.PopupMenu m_PopUpEditMenu
            m_PlainTextBox.Enabled = True
        End If
    End If
End Sub


'File functions
Public Sub LoadFile()
    'Commondialog
    Dim OFName As OPENFILENAME
    Dim f As Integer, temp As String, mfile As String
    If Not CheckSave Then Exit Sub 'Save changes?
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = m_PlainTextBox.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Plain text (*.txt)" & Chr(0) & "*.txt" & Chr(0) & "All files (*.*)" & Chr(0) & "*.*"
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.Flags = 5
    If GetOpenFileName(OFName) Then
        'The user selected a file - open it!
        mfile = StripTerminator(Trim$(OFName.lpstrFile))
        f = FreeFile
        DoEvents
        Open mfile For Binary As #f
            temp = String(LOF(f), Chr$(0))
            Get #f, , temp
        Close #f
        m_PlainTextBox.Text = temp
        m_CurrentFile = mfile
        If m_DisplayFileNameInTitleBar Then m_TitleBarForm.Caption = App.Title & " - " & Mid$(m_CurrentFile, InStrRev(m_CurrentFile, "\") + 1)
        FileChanged = False
    End If

End Sub

Public Function SaveAsFile() As Boolean
    'Commondialog
    Dim OFName As OPENFILENAME, fTitle As String
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = m_PlainTextBox.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Plain text (*.txt)" & Chr(0) & "*.txt" & Chr(0) & "All files (*.*)" & Chr(0) & "*.*"
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.Flags = 5
    If GetSaveFileName(OFName) Then
        fTitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
        m_CurrentFile = StripTerminator(Trim$(OFName.lpstrFile))
        If OFName.nFilterIndex <> 1 Then 'any file
            If InStr(1, fTitle, ".") = 0 Then 'did user add an extension?
                fTitle = fTitle & ".txt" 'if not add txt for them
                m_CurrentFile = m_CurrentFile & ".txt"
            End If
        Else
            'make sure the extension is txt
            fTitle = ChangeExt(fTitle, "txt")
            m_CurrentFile = ChangeExt(m_CurrentFile, "txt")
        End If
        If m_DisplayFileNameInTitleBar Then m_TitleBarForm.Caption = App.Title & " - " & fTitle
        SaveAsFile = SaveFile
    Else
        SaveAsFile = False
    End If
End Function
Public Function SaveFile(Optional filepath As String) As Boolean
    Dim f As Integer
    On Error Resume Next
    If filepath = "" Then
        filepath = m_CurrentFile
        If m_CurrentFile = "" Then
            SaveFile = SaveAsFile
            Exit Function
        End If
    End If
    If Dir(filepath) <> "" Then Kill filepath
    'Write to disk
    f = FreeFile
    Open filepath For Binary As #f
        Put #f, , m_PlainTextBox.Text
        Close #f
    SaveFile = True
    FileChanged = False
End Function
Public Sub NewFile()
    If Not CheckSave Then Exit Sub 'Save changes?
    m_CurrentFile = ""
    m_PlainTextBox.Text = ""
    If m_DisplayFileNameInTitleBar Then m_TitleBarForm.Caption = App.Title & " - Untitled.txt"
    FileChanged = False 'Reset flag
End Sub
Public Function CheckSave() As Boolean
    Dim Response As VbMsgBoxResult
    'Used to confirm saving changes when opening another file
    'or creating a new one
    If Not FileChanged Then
        CheckSave = True
        Exit Function
    End If
    Response = MsgBox("This file has changed. Do you want to save changes?", vbYesNoCancel + vbQuestion)
    Select Case Response
        Case vbYes
            CheckSave = SaveFile
        Case vbNo
            CheckSave = True
        Case vbCancel
            CheckSave = False
    End Select
End Function
Private Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    'Simple string parse
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = Mid$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Private Function StripTerminator(ByVal strString As String) As String
    'API nearly always returns a string with trailing null characters
    'so we need to remove them
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'Find and Replace
Public Function FindString(Find As String, StartPosition As Long, Optional MatchCase As Boolean = False) As Boolean
    Dim z As Long, temp As String
    temp = m_PlainTextBox.Text
    z = InStr(StartPosition + 1, temp, Find, IIf(MatchCase, vbBinaryCompare, vbTextCompare))
    If z <> 0 Then
        m_PlainTextBox.SelStart = z - 1
        m_PlainTextBox.SelLength = Len(Find)
        FindString = True
    Else
        FindString = False
    End If
    m_PlainTextBox.SetFocus
End Function

Public Function ReplaceString(Find As String, vReplace As String, Start As Long, Optional ReplaceAll As Boolean = False, Optional MatchCase As Boolean = False) As Boolean
    Dim z As Long, temp As String
    'We're using the inbuilt vb function "Replace" which is
    'actually a "Regular Expression" operation
    temp = m_PlainTextBox.Text
    z = InStr(Start + 1, temp, Find, IIf(MatchCase, vbBinaryCompare, vbTextCompare))
    If z <> 0 Then
        If ReplaceAll Then
            m_PlainTextBox.Text = Replace(temp, Find, vReplace, 1, , IIf(MatchCase, vbBinaryCompare, vbTextCompare))
        Else
            'Notice we need the "Left" function to grab the text BEFORE the start position
            m_PlainTextBox.Text = Left(temp, Start) & Replace(temp, Find, vReplace, Start + 1, 1, IIf(MatchCase, vbBinaryCompare, vbTextCompare))
        End If
        ReplaceString = True
    Else
        ReplaceString = False
    End If
    m_PlainTextBox.SetFocus
End Function
'Font dialog
Public Function ShowFont() As Boolean
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim retval As Long
    'Firstly load into the LOGFONT structure our current font
    lfont.lfHeight = 0
    lfont.lfItalic = m_PlainTextBox.Font.Italic
    lfont.lfUnderline = m_PlainTextBox.Font.Underline
    lfont.lfStrikeOut = m_PlainTextBox.Font.Strikethrough
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfHeight = m_PlainTextBox.Font.Size * 1.33
    If m_PlainTextBox.Font.Bold Then
        lfont.lfWidth = 700
    Else
        lfont.lfWidth = 0
    End If
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = m_PlainTextBox.Font.Name & vbNullChar
    'Now move the LOGFONT structure into the Font Dialog
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)
    cf.lStructSize = Len(cf)
    cf.hwndOwner = m_PlainTextBox.hWnd
    cf.lpLogFont = pMem
    cf.iPointSize = m_PlainTextBox.Font.Size * 10
    cf.rgbColors = m_PlainTextBox.ForeColor
    cf.nSizeMin = 10
    cf.nSizeMax = 72
    'Set the dialog options we want
    cf.Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT 'Or CF_USESTYLE
    'Show the dialog
    retval = CHOOSEFONT(cf)
    If retval <> 0 Then
        ShowFont = True
        CopyMemory lfont, ByVal pMem, Len(lfont)
        'We need to set some properties to false first
        m_PlainTextBox.Font.Bold = False
        m_PlainTextBox.Font.Italic = False
        m_PlainTextBox.Font.Underline = False
        m_PlainTextBox.Font.Strikethrough = False
        'Apply the returned settings straight into the current textbox
        m_PlainTextBox.Font.Name = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        m_PlainTextBox.Font.Size = cf.iPointSize / 10
        If lfont.lfWeight = 700 Then m_PlainTextBox.Font.Bold = True
        m_PlainTextBox.Font.Italic = lfont.lfItalic
        m_PlainTextBox.Font.Underline = lfont.lfUnderline
        m_PlainTextBox.Font.Strikethrough = lfont.lfStrikeOut
        m_PlainTextBox.ForeColor = cf.rgbColors
    Else
        'Cancel/close pressed
        ShowFont = False
    End If
    'Always free memory
    retval = GlobalUnlock(hMem)
    retval = GlobalFree(hMem)
End Function

