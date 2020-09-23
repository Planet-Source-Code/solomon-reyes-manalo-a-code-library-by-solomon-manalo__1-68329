VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl HTMLControl 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   ScaleHeight     =   4155
   ScaleWidth      =   5250
   Begin VB.PictureBox picLineNumbers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   1215
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   3255
      Left            =   1260
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"HTMLControl.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopupCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPopupPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPopupSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPopupSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupWordWrap 
         Caption         =   "Word Wrap"
      End
   End
End
Attribute VB_Name = "HTMLControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************
' Subclass Declarations
'*****************************************************************************
'Implements ISubclass
'Private m_emr As EMsgResponse
'*****************************************************************************
' API Declarations
'*****************************************************************************
Public bChange As Boolean

Private Const EM_SETMARGINS = &HD3
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETSEL = &HB1
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINESCROLL = &HB6
Private Const EM_HIDESELECTION = &H43F
Private Const EC_LEFTMARGIN = 1

Private Const WM_COMMAND = &H111
'Private Const WM_VSCROLL = &H115
'Private Const WM_HSCROLL = &H114 ' 276
'Private Const WM_MOUSEWHEEL = &H20A
'Private Const WM_KEYDOWN = &H100
'Private Const WM_KEYUP = &H101
'Private Const WM_CHAR = &H102
'Private Const WM_PAINT = 15
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageVal Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'*****************************************************************************
' Const Declarations
'*****************************************************************************
Private Const def_booWordWrap = False
Private Const def_booAutoIndent = True
Private Const def_booReadOnly = False
Private Const VScrollBarHeight = 315

'*****************************************************************************
' Variable Declarations
'*****************************************************************************
Private pWordWrap As Boolean
Private pReadOnly As Boolean
Private pAutoIndent As Boolean
Private pDisableSelChange As Boolean
Private TabRegex As RegExp
Private SaveLoc As clsRTBSelPosSaver
Private cMatches As MatchCollection
Private Matches As Match
Private WithEvents RTFUndo As clsUndo
Attribute RTFUndo.VB_VarHelpID = -1

'Subclassing Variables
'mWndProcOrg holds the original address of the Window Procedure for this window. This is used to
'route messages to the original procedure after you process them.

Private mWndProcOrg      As Long
Private mHWndSubClassed  As Long                         'Handle (hWnd) of the subclassed window.



Public Property Get SelText() As String
    SelText = rtbText.SelText
End Property
Public Property Let SelText(ByVal strNewValue As String)
    Screen.MousePointer = vbHourglass
    rtbText.Text = strNewValue
    Screen.MousePointer = vbNormal
End Property
Public Property Let SelStart(ByVal newSelStart As Long)
    rtbText.SelStart = newSelStart
End Property
Public Property Let SelLength(ByVal newSelLength As Long)
    rtbText.SelLength = newSelLength
End Property
Public Property Get SelStart() As Long
    SelStart = rtbText.SelStart
End Property
Public Property Get SelLength() As Long
    SelLength = rtbText.SelLength
End Property
Public Property Let SetWidth(newWidth As Long)
    UserControl.Width = newWidth
End Property
Public Property Let SetHeight(newHeight As Long)
    UserControl.Height = newHeight
End Property
Public Property Get Text() As String
    Text = rtbText.Text
End Property
Public Property Get SelRTF() As String
    SelRTF = rtbText.SelRTF
End Property
Public Property Get TextRTF() As String
    TextRTF = rtbText.TextRTF
End Property
Public Property Let TextRTF(newTxtRTF As String)
    rtbText.TextRTF = newTxtRTF
End Property
Public Property Let AutoIndent(ByVal Value As Boolean)
    pAutoIndent = Value
End Property
Public Property Get AutoIndent() As Boolean
    AutoIndent = pAutoIndent
End Property
Public Property Get ReadOnly() As Boolean
   ReadOnly = pReadOnly
End Property
Public Property Let ReadOnly(ByVal Value As Boolean)
   pReadOnly = Value
   rtbText.BackColor = IIf(pReadOnly, RGB(230, 230, 230), vbWhite)
   SendMessage rtbText.hWnd, EM_SETREADONLY, ByVal CLng(pReadOnly), 0&
End Property
Public Property Get WordWrap() As Boolean
    WordWrap = pWordWrap
End Property
Public Property Let WordWrap(ByVal Value As Boolean)
    pWordWrap = Value
    mnuPopupWordWrap.Checked = Value
    If pWordWrap Then
        ' true
        rtbText.RightMargin = rtbText.Width - (picLineNumbers.TextWidth("12345678") + Screen.TwipsPerPixelX)
    Else
        rtbText.RightMargin = 9999999
    End If
    
    WriteLineNumbers
End Property
Public Property Get canUndo() As Boolean
   canUndo = RTFUndo.canUndo And Not ReadOnly
End Property
Public Property Get canRedo() As Boolean
   canRedo = RTFUndo.canRedo And Not ReadOnly
End Property
Public Property Get FirstVisibleLine() As Long
   FirstVisibleLine = SendMessage(rtbText.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&) + 1
End Property
Public Property Let FirstVisibleLine(ByVal line As Long)
   LineScroll line - FirstVisibleLine
End Property
Public Property Get getLineIndex(ByVal line As Long) As Long
   getLineIndex = SendMessage(rtbText.hWnd, EM_LINEINDEX, ByVal (line - 1), 0)
End Property
Public Property Get getLineNumber(ByVal charPos As Long) As Long
   getLineNumber = SendMessage(rtbText.hWnd, EM_LINEFROMCHAR, charPos, 0&) + 1
End Property

Private Sub mnuPopupCopy_Click()
    Copy
End Sub

Private Sub mnuPopupCut_Click()
    Cut
End Sub

Private Sub mnuPopupPaste_Click()
    Paste
End Sub

Private Sub mnuPopupSelectAll_Click()
    SelectALL
End Sub

Private Sub mnuPopupWordWrap_Click()
    WordWrap = Not mnuPopupWordWrap.Checked
End Sub

Private Sub picLineNumbers_Click()
WriteLineNumbers
End Sub

Private Sub picLineNumbers_GotFocus()
WriteLineNumbers
End Sub

Private Sub picLineNumbers_LostFocus()
WriteLineNumbers
End Sub

Private Sub picLineNumbers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
WriteLineNumbers
End Sub

Private Sub picLineNumbers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
WriteLineNumbers
End Sub

Private Sub picLineNumbers_Paint()
WriteLineNumbers
End Sub

Private Sub picLineNumbers_Resize()
WriteLineNumbers
End Sub

Private Sub rtbText_Change()
bChange = True
WriteLineNumbers
End Sub

Private Sub rtbText_Click()
WriteLineNumbers
End Sub

Private Sub rtbText_GotFocus()
    On Error Resume Next
    Dim Control As Control
    For Each Control In Controls
        Control.TabStop = False
    Next Control
End Sub

Private Sub rtbText_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim offSet As Integer
    LockWindowUpdate True
    If KeyCode = vbKeyReturn And Shift = 0 And pAutoIndent Then RunAutoIndent: KeyCode = 0
    If Shift = vbCtrlMask Then
        Select Case KeyCode
        Case vbKeyA
            ' Select all
            SelectALL
        Case vbKeyC
            'copy
            Copy
        Case vbKeyO
            MsgBox CountLineReturns
            MsgBox "FVL: " & FirstVisibleLine
        Case vbKeyR
            Redo
        Case vbKeyV
            ' paste
            Paste
        Case vbKeyX
            Cut
        Case vbKeyZ
            Undo
        End Select
        rtbText.SetFocus
        KeyCode = 0
        Shift = 0
    Else
        Select Case KeyCode
        Case vbKeyTab
            'RichTxtBox.SelRTF = vbTab
            If Shift Then
                If rtbText.SelLength = 0 Then
                    Set SaveLoc = New clsRTBSelPosSaver
                    SaveLoc.Save rtbText
                    LockWindowUpdate True
                    rtbText.SelStart = getLineIndex(getLineNumber(rtbText.SelStart))
                    rtbText.SelLength = 1
                    If rtbText.SelText = Chr$(vbKeyTab) Then
                        rtbText.SelText = ""
                        offSet = 1
                    Else
                        offSet = 0
                    End If
                    LockWindowUpdate False
                    SaveLoc.Restore rtbText, offSet
                    Set SaveLoc = Nothing
                Else
                    BlockUnIndent
                End If
            Else
                If rtbText.SelLength = 0 Then
                    rtbText.SelText = vbTab
                Else
                    BlockIndent
                End If
            End If
            KeyCode = 0
        End Select
    End If
    LockWindowUpdate False
End Sub

Private Sub rtbText_LostFocus()
WriteLineNumbers
End Sub

Private Sub rtbText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub



Private Sub rtbText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
WriteLineNumbers
End Sub

Private Sub rtbText_SelChange()
    Dim Ln As Long
    Ln = rtbText.SelLength
    With UserControl
        ' Determine which options are available
        .mnuPopupCut.Enabled = Ln
        .mnuPopupCopy.Enabled = Ln
        .mnuPopupPaste.Enabled = Len(Clipboard.GetText(1))
        .mnuPopupSelectAll.Enabled = CBool(Len(rtbText.Text))
    End With
    If Not pDisableSelChange Then
        WriteLineNumbers
    End If
End Sub

Private Sub rtbText_Validate(Cancel As Boolean)
WriteLineNumbers
End Sub



Private Sub UserControl_Initialize()
    'SubClass
    Set TabRegex = New RegExp
    TabRegex.Pattern = "(\t)"
    TabRegex.Global = True
    TabRegex.IgnoreCase = False
    
    Set RTFUndo = New clsUndo
    RTFUndo.AssignToRichTextBox UserControl.Controls, rtbText
    
    SendMessageLong rtbText.hWnd, EM_SETMARGINS, EC_LEFTMARGIN, (picLineNumbers.TextWidth("12345") / Screen.TwipsPerPixelX + 5)
    WordWrap = False
    rtbText_SelChange
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    WordWrap = PropBag.ReadProperty("HLWordWrap", def_booWordWrap)
    AutoIndent = PropBag.ReadProperty("HLAutoIndent", def_booAutoIndent)
    ReadOnly = PropBag.ReadProperty("HLReadOnly", def_booReadOnly)
End Sub

Private Sub UserControl_Resize()
    picLineNumbers.Move 0, 0, picLineNumbers.ScaleWidth, ScaleHeight
    rtbText.Move 0, 0, ScaleWidth, ScaleHeight
    WriteLineNumbers
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("HLWordWrap", pWordWrap, def_booWordWrap)
    Call PropBag.WriteProperty("HLAutoIndent", pAutoIndent, def_booAutoIndent)
    Call PropBag.WriteProperty("HLReadOnly", pReadOnly, def_booReadOnly)
End Sub

Private Sub UserControl_Terminate()
    'UnSubClass
End Sub

'*****************************************************************************
' General Functions
'*****************************************************************************
Public Sub Undo()
   If ReadOnly Then Exit Sub
   RTFUndo.Undo
End Sub
Public Sub Redo()
   If ReadOnly Then Exit Sub
   RTFUndo.Redo
End Sub
Public Sub Copy()
    If ReadOnly Then Exit Sub
    EditFunction WM_COPY
End Sub
Public Sub Cut()
    If ReadOnly Then Exit Sub
    EditFunction WM_CUT
End Sub
Public Sub Paste()
    If ReadOnly Then Exit Sub
    EditFunction WM_PASTE
End Sub
Public Sub SelectALL()
    If ReadOnly Then Exit Sub
    rtbText.SelStart = 0
    rtbText.SelLength = Len(rtbText.Text)
    rtbText.SetFocus
End Sub

Private Sub EditFunction(Action As Integer)
  Call SendMessage(rtbText.hWnd, Action, 0, 0&)
  If Action <> WM_COPY Then rtbText.SelText = ""
End Sub

Private Sub WriteLineNumbers()
    Dim y           As Long
    Dim x           As Long
    Dim lStart      As Long
    Dim FontHeight  As Long
    Dim lFinish     As Long
    Dim lCurrent    As Long, _
    twiX As Long, _
    twiY As Long
       
    twiX = Screen.TwipsPerPixelX
    twiY = Screen.TwipsPerPixelY

    lStart = FirstVisibleLine
    lCurrent = CurrentLine
    With picLineNumbers
        .Move 2 * twiX, 2 * twiY, picLineNumbers.TextWidth("12345") + twiX, ZeroIfNegative(ScaleHeight - VScrollBarHeight)
        .Cls
        .Font = rtbText.Font
        .FontSize = rtbText.Font.Size
        FontHeight = .TextHeight("12345")
        .CurrentY = Screen.TwipsPerPixelY * 2
        '.CurrentY = .CurrentY + 15
        lFinish = (rtbText.Height / FontHeight) + lStart
        If lFinish > LineCount Then lFinish = LineCount
        y = CountLineReturns
        ' loop from the first visible line in the rtb to the end of the page
        For x = lStart To lFinish
            If x = lCurrent Then
                .FontBold = True
                .ForeColor = vbBlue
                .FontUnderline = True
            Else
                .ForeColor = &H808080
                .FontBold = False
                .FontUnderline = False
            End If
            ' check for wordwrap
            If WordWrap Then
                If GetLine(x) Or x = 1 Then
                    picLineNumbers.Print Right$("     " & y, 5)
                    y = y + 1
                Else
                    picLineNumbers.Print Right$("     -", 5)
                End If
                
            Else
                picLineNumbers.Print Right$("     " & x, 5)
            End If
            picLineNumbers.Refresh
        Next x
    End With
End Sub
Private Function CountLineReturns() As Integer
    Dim FVL As Long
    Dim x, y As Integer
    y = 0
    FVL = FirstVisibleLine
    For x = 0 To FVL - 1
        If GetLineNoMinus(x) Then
            y = y + 1
        End If
    Next x
    CountLineReturns = y
End Function
Private Function GetLineNoMinus(ByVal lineNum As Integer) As Boolean
    Dim sBuffer As String
    Dim retVal
    Dim retVal1
    sBuffer = String(255, Chr(32))
    retVal = SendMessageStr(rtbText.hWnd, EM_GETLINE, lineNum, ByVal sBuffer)
    retVal1 = Left(sBuffer, retVal)
    If Right(retVal1, 1) = vbLf Or Right(retVal1, 1) = vbCrLf Or Right(retVal1, 1) = vbCr Then
        GetLineNoMinus = True
    Else
        GetLineNoMinus = False
    End If
End Function
Public Function LineCount() As Long
    ' return the total line count of the code window
    LineCount = SendMessage(rtbText.hWnd, EM_GETLINECOUNT, 0, 0)
End Function
Public Sub LineScroll(ByVal Lines As Long)
   SendMessageVal rtbText.hWnd, EM_LINESCROLL, 0&, Lines
End Sub
Public Function CurrentLine() As Long
    CurrentLine = SendMessage(rtbText.hWnd, EM_LINEFROMCHAR, rtbText.SelStart, 0) + 1
End Function
Private Sub RunAutoIndent()
    Dim s As String, l As Long, I As Integer
    s = GetLineText
    Set cMatches = TabRegex.Execute(s)
    l = cMatches.count
    rtbText.SelText = vbCrLf
    For I = 1 To l
        rtbText.SelText = vbTab
    Next I
End Sub

Private Function GetLine(ByVal lineNum As Integer) As Boolean
    Dim sBuffer As String
    Dim retVal
    Dim retVal1
    sBuffer = String(255, Chr(32))
    lineNum = ZeroIfNegative(lineNum - 2)
    'MsgBox lineNum
    retVal = SendMessageStr(rtbText.hWnd, EM_GETLINE, lineNum, ByVal sBuffer)
    retVal1 = Left(sBuffer, retVal)
    'Debug.Print retVal1
    If Right(retVal1, 1) = vbLf Or Right(retVal1, 1) = vbCrLf Or Right(retVal1, 1) = vbCr Then
        GetLine = True
    Else
        GetLine = False
    End If
End Function
Public Function GetLineText() As String
    On Error Resume Next
    Dim line As Long, lngStart As Long
    Dim start As Long
    line = CurrentLine
    lngStart = SendMessage(rtbText.hWnd, EM_LINEINDEX, line - 1, 0&)
    start = lngStart
    line = line + 1
    lngStart = SendMessage(rtbText.hWnd, EM_LINEINDEX, line - 1, 0&)
    If lngStart = -1 Then lngStart = Len(rtbText.Text) + 2
    GetLineText = Mid$(rtbText.Text, start + 1, lngStart - start - 2)
End Function
Public Function GetColumn() As Integer
   Dim lLine As Long
   Dim cCol As Long, lChar As Long, I As Long

   lChar = rtbText.SelStart + 1
   cCol = SendMessageLong(rtbText.hWnd, EM_LINELENGTH, lChar - 1, 0&)
   lLine = 1 + SendMessageLong(rtbText.hWnd, EM_LINEFROMCHAR, rtbText.SelStart, 0&)
   I = SendMessageLong(rtbText.hWnd, EM_LINEINDEX, lLine - 1, 0&)
   GetColumn = lChar - I - 1

End Function

'*****************************************************************************
' Block Indent and UnIndent
'*****************************************************************************
Public Function BlockIndent()
   Dim firstLine As Long, _
    lastLine As Long, _
    FVL As Long, _
    I As Long, _
    oldWordWrap As Boolean
   
   pDisableSelChange = True
   FVL = FirstVisibleLine
   
   oldWordWrap = WordWrap
   If oldWordWrap Then WordWrap = False
   
   ' get the range to intend
   firstLine = getLineNumber(rtbText.SelStart)
   lastLine = getLineNumber(rtbText.SelStart + rtbText.SelLength - 1)

   ' do the intend
   For I = firstLine To lastLine
      rtbText.SelStart = getLineIndex(I)
      rtbText.SelLength = 0
      rtbText.SelText = Chr$(vbKeyTab)
   Next
   
   ' set the text selection
   rtbText.SelStart = getLineIndex(firstLine)
   If lastLine = LineCount Then
      rtbText.SelLength = Len(rtbText.Text)
   Else
      rtbText.SelLength = getLineIndex(lastLine + 1) - 1 - getLineIndex(firstLine)
   End If
   
   If oldWordWrap Then WordWrap = True
   FirstVisibleLine = FVL
   pDisableSelChange = False
End Function
Public Function BlockUnIndent()
   Dim firstLine As Long, _
       lastLine As Long, _
       FVL As Long, _
       I As Long, _
       oldWordWrap As Boolean
   
   pDisableSelChange = True
   FVL = FirstVisibleLine
   
   oldWordWrap = WordWrap
   If oldWordWrap Then WordWrap = False
   
   ' get the range to intend
   firstLine = getLineNumber(rtbText.SelStart)
   lastLine = getLineNumber(rtbText.SelStart + rtbText.SelLength - 1)

   ' do the intend
   For I = firstLine To lastLine
      rtbText.SelStart = getLineIndex(I)
      rtbText.SelLength = 1
      If rtbText.SelText = Chr$(vbKeyTab) Then rtbText.SelText = ""
   Next
   
   ' set the text selection
   rtbText.SelStart = getLineIndex(firstLine)
   If lastLine = LineCount Then
      rtbText.SelLength = Len(rtbText.Text)
   Else
      rtbText.SelLength = getLineIndex(lastLine + 1) - 1 - getLineIndex(firstLine)
   End If

   FirstVisibleLine = FVL
   pDisableSelChange = False
End Function
'=================================== SUBCLASSING CODE FOR USERCONTROL (Support MouseWheel on 98 and NT4)=========
'see also mMouseWheel.bas

Private Sub SubClass()
  '-------------------------------------------------------------
  'Initiates the subclassing of this UserControl's window (hwnd).
  'Records the original WinProc of the window in mWndProcOrg.
  'Places a pointer to the object in the window's UserData area.
  '-------------------------------------------------------------

  'Exit if the window is already subclassed.
  If mWndProcOrg <> 0 Then Exit Sub

  'Redirect the window's messages from this control's default Window Procedure to the SubWndProc function
  'in your .BAS module and record the address of the previous Window Procedure for this window in mWndProcOrg.
  mWndProcOrg = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)
  'mWndProcOrg = SetWindowLong(rtbText.hWnd, GWL_WNDPROC, AddressOf SubWndProc)

  'Record your window handle in case SetWindowLong gave you a new one.
  'You will need this handle so that you can unsubclass.
  mHWndSubClassed = hWnd

  'Store a pointer to this object in the UserData section of this window that will be used later to get
  'the pointer to the control based on the handle (hwnd) of the window getting the message.
  Call SetWindowLong(hWnd, GWL_USERDATA, ObjPtr(Me))
    
  'Get the Size of a Wheel Scroll in lines
  gucWheelScrollLines = SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, pulScrollLines, 0)
End Sub

Private Sub UnSubClass()
  '-----------------------------------------------------------------------------------------------
  'Unsubclasses this UserControl's window (hwnd), setting the address of the Windows Procedure
  'back to the address it was at before it was subclassed.
  '------------------------------------------------------------------------------------------------
  
  If mWndProcOrg = 0 Then Exit Sub  'Ensures that you don't try to unsubclass the window when it is not subclassed.
  SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg     'Reset the window's function back to the original address.
  mWndProcOrg = 0                   '0 Indicates that you are no longer subclassed.
End Sub

Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '--------------------------------------------------------------
  'Process the window's messages that are sent to your UserControl. The WindowProc function is declared as
  'a "Friend" function so that the .BAS module can call the function but the function cannot be seen from
  'outside the UserControl project.
  '--------------------------------------------------------------

  'We are only intersetsed in picking up Mousewheel messages. We handle them as if the correct scroll key
  'had been repeatedly pressed by the user
  
  'Dim ScrollAmt As Long
  Select Case uMsg
    Case WM_COMMAND:
      If (wParam And (MK_SHIFT Or MK_CONTROL)) = 0 Then   ' Don't handle zoom and datazoom.
        'MsgBox "hello"
        'gcWheelDelta = gcWheelDelta - (wParam And &HFFFF0000) / 65536
        'If Abs(gcWheelDelta) >= WHEEL_DELTA Then
        '
        '  ScrollAmt = gcWheelDelta / WHEEL_DELTA
        '
        '  Do While gcWheelDelta < -WHEEL_DELTA
        '    gcWheelDelta = gcWheelDelta + WHEEL_DELTA
        '  Loop
        '  Do While gcWheelDelta > WHEEL_DELTA
        '    gcWheelDelta = gcWheelDelta - WHEEL_DELTA
        '  Loop
            
        '  Call PretendMouseKey(ScrollAmt)
        'End If
        WriteLineNumbers
      End If
  End Select
  
  'Forwards the window's messages that came in to the original Window Procedure that handles the messages
  'and returns the result back to the SubWndProc function.
  WindowProc = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, ByVal lParam)
End Function

