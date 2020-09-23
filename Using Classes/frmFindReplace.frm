VERSION 5.00
Begin VB.Form frmFindReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find and Replace"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   300
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   300
      Left            =   3840
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Enabled         =   0   'False
      Height          =   300
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox ChMatchCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - December 2002
'If you got it elsewhere - they stole it from PSC.

'Please visit our website at www.psst.com.au

Option Explicit
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const vbMsgBoxOnTop As Long = &H40000
Private CurrentStart As Long

Private Sub cmdFind_Click()
    'The actual search is done in the class and returns "True" if anything was found
    If frmStandardMenu.ClTxt.FindString(txtFind.Text, CurrentStart, CBool(ChMatchCase.Value)) Then
        'The class will have selected the text if it found it so move
        'the variable accordingly
        CurrentStart = frmStandardMenu.ClTxt.PlainTextBox.SelStart
    Else
        MsgBox "Search text is not found", vbExclamation + vbMsgBoxOnTop
        CurrentStart = 0
    End If
    cmdFind.Caption = "Find next"
End Sub

Private Sub cmdReplace_Click()
    If frmStandardMenu.ClTxt.ReplaceString(txtFind.Text, txtReplace.Text, CurrentStart, , CBool(ChMatchCase.Value)) Then
        CurrentStart = frmStandardMenu.ClTxt.PlainTextBox.SelStart
        'Do a search for the next item ready for another replace
        If frmStandardMenu.ClTxt.FindString(txtFind.Text, CurrentStart, CBool(ChMatchCase.Value)) Then
            CurrentStart = frmStandardMenu.ClTxt.PlainTextBox.SelStart
        End If
    Else
        MsgBox "Search text is not found", vbExclamation + vbMsgBoxOnTop
        CurrentStart = 0
    End If
End Sub

Private Sub cmdReplaceAll_Click()
    If frmStandardMenu.ClTxt.ReplaceString(txtFind.Text, txtReplace.Text, CurrentStart, True, CBool(ChMatchCase.Value)) Then
        CurrentStart = frmStandardMenu.ClTxt.PlainTextBox.SelStart
    Else
        MsgBox "Search text is not found", vbExclamation + vbMsgBoxOnTop
        CurrentStart = 0
    End If
End Sub

Private Sub Form_Load()
    'Form on top, use main form's icon
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 1 Or 2
    Me.Icon = frmStandardMenu.Icon
End Sub
'Enable buttons according to textbox contents
Private Sub txtFind_Change()
    cmdFind.Enabled = Len(txtFind.Text) > 0
    CurrentStart = 0
End Sub

Private Sub txtReplace_Change()
    cmdReplace.Enabled = Len(txtFind.Text) > 0
    cmdReplaceAll.Enabled = cmdReplace.Enabled
    CurrentStart = 0
End Sub
