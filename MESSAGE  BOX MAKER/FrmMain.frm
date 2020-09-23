VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Box Maker"
   ClientHeight    =   4770
   ClientLeft      =   4605
   ClientTop       =   3465
   ClientWidth     =   6480
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MESSAGE MAIN"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6480
   Begin VB.Frame FraVBCODE 
      Caption         =   "Generated Code"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   6255
      Begin VB.TextBox txtVBCODE 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame FraInput 
      Caption         =   "Input"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton CMDABOUT 
         Caption         =   "&About"
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox lstBUTTONTYPE 
         Height          =   840
         ItemData        =   "FrmMain.frx":0442
         Left            =   4440
         List            =   "FrmMain.frx":0458
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton CMDPREVIEW 
         Caption         =   "&Preview"
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton CMDCODE 
         Caption         =   "&Generate Code"
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox LSTMMODALTYPE 
         Height          =   840
         ItemData        =   "FrmMain.frx":04B1
         Left            =   2640
         List            =   "FrmMain.frx":04BB
         TabIndex        =   3
         Top             =   1920
         Width           =   1575
      End
      Begin VB.PictureBox PICCRITICAL 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Picture         =   "FrmMain.frx":04E0
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox PICExclaimation 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Picture         =   "FrmMain.frx":0922
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox PICINFO 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Picture         =   "FrmMain.frx":0D64
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox PICQuestion 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Picture         =   "FrmMain.frx":11A6
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox lstIconType 
         Height          =   840
         ItemData        =   "FrmMain.frx":15E8
         Left            =   720
         List            =   "FrmMain.frx":15FB
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtPrompt 
         Height          =   975
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "FrmMain.frx":1648
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtTITLE 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Text            =   "Message Box Maker"
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label LbLPrompt 
         Caption         =   "Prompt:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label LbLTITLE 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Menu mnucode 
      Caption         =   "&Code"
      Begin VB.Menu mnucodecopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnucodeline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucodeDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuGenerate 
      Caption         =   "Generate"
      Visible         =   0   'False
      Begin VB.Menu mnugenVB 
         Caption         =   "&Visual Basic"
      End
   End
   Begin VB.Menu mnuhidden 
      Caption         =   "<HIDDEN MENU>"
      Visible         =   0   'False
      Begin VB.Menu mnuhiddenshow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuhiddenEXIT 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203    'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private nid As NOTIFYICONDATA

'user defined type required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
' ICONS
Private Const MB_ICONASTERISK = &H40&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONMASK = &HF0&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONHAND = &H10&
Private Const MB_ICONSTOP = MB_ICONHAND
' MODALS
Private Const MB_APPLMODAL = &H0&
Private Const MB_SYSTEMMODAL = &H1000&
'BUTTONS
Private Const MB_OK = &H0&
Private Const MB_OKCANCEL = &H1&
Private Const MB_RETRYCANCEL = &H5&
Private Const MB_YESNO = &H4&
Private Const MB_YESNOCANCEL = &H3&
Private Const MB_ABORTRETRYIGNORE = &H2&

' DIMS
Private MODALTYPE As Variant
Private ICONTYPE As Variant
Private BUTTONTYPE As Variant
Private TXTPROMPT1, TXTTITLE1 As Variant
Private VBQUOTE As String
Option Explicit

Private Sub CMDABOUT_Click()
MessageBox Me.hwnd, "Hey," & Chr(10) & "This program was made by Jaime Muscatelli. This program was made in Visual Basic 6 Professional. This program is freeware, and I used WIN32 API to write the MSGBOX(MessageBox) CODE.  " & Chr(10) & "      I want to thank Impulse, Xeo, and Jenn." & Chr(10) & Chr(10) & "                                - JAIME ", "VB Message Box Maker", &H40& + MB_SYSTEMMODAL + MB_OK
End Sub

Private Sub CMDCODE_Click()
PopupMenu mnuGenerate
End Sub

Private Sub CMDPREVIEW_Click()
TXTPROMPT1 = txtPrompt.Text
TXTTITLE1 = txtTITLE.Text
Beep
MessageBox Me.hwnd, txtPrompt, TXTTITLE1, ICONTYPE + MODALTYPE + BUTTONTYPE
End Sub

Private Sub Form_Load()
Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        ''''''
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ''''''The callback should be the mousemove event
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        ''''''Heres the tooltip in the taskbar'''''
        .szTip = "VB Message Box Maker" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim RESULT As Long
    Dim msg As Long

    'set up the msg handler
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg
        'respond to the right mouse button
        Case WM_RBUTTONUP
            'bring our window to the foreground
            RESULT = SetForegroundWindow(Me.hwnd)
            'mnuPopupMenu is in the menu editor for this form
            'it is set to be invisible
            Me.PopupMenu mnuhidden
        'respond to the left double click
        Case WM_LBUTTONDBLCLK
            Me.Show
            Me.WindowState = 0
            RESULT = SetForegroundWindow(Me.hwnd)
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Dim msgRESULT As Integer
msgRESULT = MessageBox(Me.hwnd, "Do you wish to exit VB Message Box Maker?", Me.Caption, MB_YESNO + MB_ICONQUESTION + MB_SYSTEMMODAL)
If msgRESULT = vbYes Then
Shell_NotifyIcon NIM_DELETE, nid
End
End If
End Sub

Private Sub lstBUTTONTYPE_Click()
If lstBUTTONTYPE.ListIndex = "0" Then
BUTTONTYPE = &H0&
ElseIf lstBUTTONTYPE.ListIndex = "1" Then
BUTTONTYPE = &H1&
ElseIf lstBUTTONTYPE.ListIndex = "2" Then
BUTTONTYPE = &H4&
ElseIf lstBUTTONTYPE.ListIndex = "3" Then
BUTTONTYPE = &H3&
ElseIf lstBUTTONTYPE.ListIndex = "4" Then
BUTTONTYPE = &H5&
ElseIf lstBUTTONTYPE.ListIndex = "5" Then
BUTTONTYPE = &H2&
End If
End Sub

Private Sub lstIconType_Click()
If lstIconType.ListIndex = "0" Then
'PIC VIEW
MODALTYPE = &O0&
PICINFO.Visible = False
PICExclaimation.Visible = False
PICQuestion.Visible = False
PICCRITICAL.Visible = False
'END PIC VIEW
ElseIf lstIconType.ListIndex = "1" Then
MODALTYPE = &H10&
'PIC VIEW
PICINFO.Visible = False
PICExclaimation.Visible = False
PICQuestion.Visible = False
PICCRITICAL.Visible = True
'END PIC VIEW
ElseIf lstIconType.ListIndex = "2" Then
MODALTYPE = &H20&
'PIC VIEW
PICINFO.Visible = False
PICExclaimation.Visible = False
PICQuestion.Visible = True
PICCRITICAL.Visible = False
'END PIC VIEW
ElseIf lstIconType.ListIndex = "3" Then
MODALTYPE = &H40&
'PIC VIEW
PICINFO.Visible = True
PICExclaimation.Visible = False
PICQuestion.Visible = False
PICCRITICAL.Visible = False
'END PIC VIEW
ElseIf lstIconType.ListIndex = "4" Then
MODALTYPE = &H30&
'PIC VIEW
PICINFO.Visible = False
PICExclaimation.Visible = True
PICQuestion.Visible = False
PICCRITICAL.Visible = False
'END PIC VIEW
End If
End Sub

Private Sub LSTMMODALTYPE_Click()
If LSTMMODALTYPE.ListIndex = "0" Then
ICONTYPE = &H1000&
ElseIf LSTMMODALTYPE.ListIndex = "1" Then
ICONTYPE = &H0&
End If
End Sub



Private Sub mnuabout_Click()
MessageBox Me.hwnd, "Hey," & Chr(10) & "This program was made by Jaime Muscatelli. This program was made in Visual Basic 6 Professional. This program is freeware, and I used WIN32 API to write the MSGBOX(MessageBox) CODE.  " & Chr(10) & "      I want to thank Impulse, Xeo, and Jenn." & Chr(10) & Chr(10) & "                                - JAIME ", "VB Message Box Maker", &H40& + MB_SYSTEMMODAL + MB_OK
End Sub

Private Sub mnucodecopy_Click()
Dim TEXTSELECT As Integer
txtVBCODE.SetFocus
TEXTSELECT = Len(txtVBCODE)
txtVBCODE.SelStart = 0
txtVBCODE.SelLength = TEXTSELECT
Clipboard.Clear
Clipboard.SetText txtVBCODE.Text
Clipboard.SetText txtVBCODE.Text
End Sub

Private Sub mnucodeDelete_Click()
txtVBCODE.SetFocus
Clipboard.Clear
txtVBCODE.SelText = ""
txtVBCODE.Text = vbNullString
End Sub
Private Sub mnugenVB_Click()
Dim NEWMODALTYPE As String
Dim NEWICONTYPE As String
Dim NEWBUTTONTYPE As String
VBQUOTE = Chr(34)
' ICONS
If MODALTYPE = &H10& Then
NEWMODALTYPE = "VBCRITICAL"
ElseIf MODALTYPE = &H20& Then
NEWMODALTYPE = "VBQUESTION"
ElseIf MODALTYPE = &H40& Then
NEWMODALTYPE = "VBINFORMATION"
ElseIf MODALTYPE = &H30& Then
NEWMODALTYPE = "VBEXCLAIMATION"
End If
' MODALS
If ICONTYPE = &H1000& Then
NEWICONTYPE = "VBSYSTEMMODAL"
ElseIf ICONTYPE = &H0& Then
NEWICONTYPE = "VBAPPLICATIONMODAL"
End If
' BUTTONS
If BUTTONTYPE = &H0& Then
NEWBUTTONTYPE = "VBOKONLY"
ElseIf BUTTONTYPE = &H1& Then
NEWBUTTONTYPE = "VBOKCANCEL"
ElseIf BUTTONTYPE = &H4& Then
NEWBUTTONTYPE = "VBYESNO"
ElseIf BUTTONTYPE = &H3& Then
NEWBUTTONTYPE = "VBYESNOCANCEL"
ElseIf BUTTONTYPE = &H5& Then
NEWBUTTONTYPE = "VBRETRYCANCEL"
ElseIf BUTTONTYPE = &H2& Then
NEWBUTTONTYPE = "VBABORTRETRYIGNORE"
End If
   If NEWMODALTYPE = vbNullString Then
   txtVBCODE.Text = "Msgbox " & VBQUOTE & txtPrompt & VBQUOTE & "," & NEWICONTYPE & " + " & NEWBUTTONTYPE & "," & VBQUOTE & txtTITLE & VBQUOTE
   Exit Sub
   End If
txtVBCODE.Text = "Msgbox " & VBQUOTE & txtPrompt & VBQUOTE & "," & NEWICONTYPE & " + " & NEWMODALTYPE & " + " & NEWBUTTONTYPE & "," & VBQUOTE & txtTITLE & VBQUOTE
End Sub

Private Sub mnuhiddenEXIT_Click()
Dim Cancel As Integer
Cancel = True
Dim msgRESULT As Integer
msgRESULT = MessageBox(Me.hwnd, "Do you wish to exit VB Message Box Maker?", Me.Caption, MB_YESNO + MB_ICONQUESTION + MB_SYSTEMMODAL)
If msgRESULT = vbYes Then
Shell_NotifyIcon NIM_DELETE, nid
End
End If
End Sub

Private Sub mnuhiddenshow_Click()
            Me.Show
            Me.WindowState = 0
End Sub

Private Sub txtPrompt_GotFocus()
Dim TEXTSELECT As Integer
TEXTSELECT = Len(txtPrompt)
txtPrompt.SelStart = 0
txtPrompt.SelLength = TEXTSELECT
End Sub

Private Sub txtTITLE_GotFocus()
Dim TEXTSELECT As Integer
TEXTSELECT = Len(txtTITLE)
txtTITLE.SelStart = 0
txtTITLE.SelLength = TEXTSELECT
End Sub


Private Sub txtVBCODE_GotFocus()
Dim TEXTSELECT As Integer
TEXTSELECT = Len(txtVBCODE)
txtVBCODE.SelStart = 0
txtVBCODE.SelLength = TEXTSELECT
End Sub
