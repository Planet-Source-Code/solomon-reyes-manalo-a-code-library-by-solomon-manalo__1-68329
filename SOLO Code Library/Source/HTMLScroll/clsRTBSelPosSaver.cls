VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRTBSelPosSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************************************************
' API Declarations
'*****************************************************************************
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Const EM_LINESCROLL = &HB6
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const WM_SETREDRAW = &HB
Private Const WM_USER = &H400
Private Const EM_GETSCROLLPOS = WM_USER + 221
Private Const EM_SETSCROLLPOS = WM_USER + 222

Private Type POINTL
    x As Long
    y As Long
End Type

'*****************************************************************************
' Public Vars
'*****************************************************************************
Public SelStart As Long
Public SelLength As Long
Public FirstVisibleLine As Long
Private lLine As POINTL

'*****************************************************************************
' Friend Methods
'*****************************************************************************
Friend Sub Save(ByRef rtb As RichTextBox)
   If Not rtb Is Nothing Then
      SelStart = rtb.SelStart
      SelLength = rtb.SelLength
      SendMessage rtb.hWnd, EM_GETSCROLLPOS, 0, lLine
      FirstVisibleLine = SendMessageLong(rtb.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&) + 1
   End If
End Sub

Friend Sub Restore(ByRef rtb As RichTextBox, Optional offSet As Integer = 0)
   If Not rtb Is Nothing Then
      rtb.SelStart = ZeroIfNegative(SelStart - offSet)
      rtb.SelLength = SelLength
      SendMessageLong rtb.hWnd, EM_LINESCROLL, 0&, FirstVisibleLine - (SendMessageLong(rtb.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&) + 1)
      SendMessage rtb.hWnd, EM_SETSCROLLPOS, 0, lLine
   End If
End Sub

