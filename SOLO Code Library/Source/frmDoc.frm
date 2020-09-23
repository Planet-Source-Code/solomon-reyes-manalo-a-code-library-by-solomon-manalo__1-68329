VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDoc 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   8670
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   3795
      Left            =   60
      ScaleHeight     =   3735
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   540
      Width           =   4995
      Begin VB.PictureBox picLines 
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
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3495
         ScaleWidth      =   840
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   840
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3495
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   6165
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmDoc.frx":0000
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
   End
   Begin CodeLibrary.McToolBar mctoolDoc 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Button_Count    =   8
      ButtonsPerRow   =   11
      HoverColor      =   16744576
      TooTipStyle     =   0
      BackGradient    =   3
      ButtonsMode     =   4
      ButtonsBackColor=   14807794
      ButtonsPerRow_Chev=   11
      ButtonToolTipIcon1=   1
      Button_Type1    =   1
      ButtonIcon2     =   "frmDoc.frx":0089
      ButtonToolTipIcon2=   1
      ButtonCaption3  =   ""
      ButtonIcon3     =   "frmDoc.frx":0803
      ButtonToolTipIcon3=   1
      ButtonCaption4  =   ""
      ButtonIcon4     =   "frmDoc.frx":0F7D
      ButtonToolTipIcon4=   1
      ButtonToolTipIcon5=   1
      Button_Type5    =   1
      ButtonCaption6  =   ""
      ButtonIcon6     =   "frmDoc.frx":16F7
      ButtonToolTipIcon6=   1
      ButtonToolTipIcon7=   1
      ButtonToolTipIcon8=   1
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SendMessageByRef Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE

Public Language, pathf As String



Private Sub Form_Resize()
On Error Resume Next
With Me
  If .Height <= 4155 Or .Width <= 5835 Then
        .Height = 4155
        .Width = 5835
  End If
     .Picture1.Left = 45
     .Picture1.Top = 510
     .Picture1.Width = .Width - 200
     .Picture1.Height = .Height - 940
     
     With .Picture1
          Me.RichTextBox1.Left = .Left + 50
          Me.RichTextBox1.Width = .Width - 180
          Me.RichTextBox1.Height = .Height - 320
     End With
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' This kills the subclass so we don't screw up someone's machine
    Call SetWindowLong(Me.RichTextBox1.hWnd, GWL_WNDPROC, lPrevWndProc)
End Sub

Private Sub Form_Load()
On Error Resume Next
    Picture1.Width = Me.ScaleWidth
    Picture1.Height = Me.ScaleHeight
    RichTextBox1.Width = Picture1.ScaleWidth - 480
    RichTextBox1.Height = Picture1.ScaleHeight
    picLines.Height = Picture1.ScaleHeight
    
    lPrevWndProc = SetWindowLong(RichTextBox1.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub mctoolDoc_Click(ByVal ButtonIndex As Long)
Select Case ButtonIndex
       Case 2:
             frmSaveCode.Show 1
End Select
End Sub

Private Sub RichTextBox1_Change()
DrawLines picLines, Me.RichTextBox1
End Sub

Private Sub RichTextBox1_GotFocus()
mdiMain.tmrActiveDoc.Enabled = True
End Sub

Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)
DrawLines picLines, Me.RichTextBox1
End Sub

Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DrawLines picLines, Me.RichTextBox1
End Sub

Private Sub RichTextBox1_SelChange()
DrawLines picLines, Me.RichTextBox1
End Sub


Public Function LineCount(txt As RichTextBox) As Long
    LineCount = SendMessageByRef(txt.hWnd, EM_GETLINECOUNT, 0&, 0&)
End Function

Public Function LineForCharacterIndex(lIndex As Long, txt As RichTextBox) As Long
   LineForCharacterIndex = SendMessageByLong(txt.hWnd, EM_LINEFROMCHAR, lIndex, 0)
End Function

Public Function FirstVisibleLine(txt As RichTextBox) As Long
   FirstVisibleLine = SendMessageByLong(txt.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Function

' This actually draws the line numbers created by the guys at vbaccelerator. Visit them at
' http://www.vbaccelerator.com. If you use this make sure to give them credit
Public Sub DrawLines(picTo As PictureBox, txt As RichTextBox)
Dim lLine As Long
Dim lCount As Long
Dim lCurrent As Long
Dim hBr As Long
Dim lEnd As Long
Dim lhDC As Long
Dim bComplete As Boolean
Dim tR As RECT, tTR As RECT
Dim oCol As OLE_COLOR
Dim lStart As Long
Dim lEndLine As Long
Dim tPO As POINTAPI
Dim lLineHeight As Long
Dim hPen As Long
Dim hPenOld As Long

   'Debug.Print "DrawLines"
   lhDC = picTo.hdc
   DrawText lhDC, "Hy", 2, tTR, DT_CALCRECT
   lLineHeight = tTR.Bottom - tTR.Top
   
   lCount = LineCount(Me.RichTextBox1)
   lCurrent = SendMessageLong(txt.hWnd, EM_LINEFROMCHAR, txt.SelStart, 0&)
   lStart = txt.SelStart
   lEnd = txt.SelStart + txt.SelLength - 1
   If (lEnd > lStart) Then
      lEndLine = LineForCharacterIndex(lEnd, Me.RichTextBox1)
   Else
      lEndLine = lCurrent
   End If
   lLine = FirstVisibleLine(Me.RichTextBox1)
   GetClientRect picTo.hWnd, tR
   lEnd = tR.Bottom - tR.Top
      
   hBr = CreateSolidBrush(TranslateColor(picTo.BackColor))
   FillRect lhDC, tR, hBr
   DeleteObject hBr
   tR.Left = 2
   tR.Right = tR.Right - 2
   tR.Top = 0
   tR.Bottom = tR.Top + lLineHeight
   
   SetTextColor lhDC, TranslateColor(vbButtonShadow)
   
   Do
      ' Ensure correct colour:
      If (lLine = lCurrent) Then
         SetTextColor lhDC, TranslateColor(vbWindowText)
      ElseIf (lLine = lEndLine + 1) Then
         SetTextColor lhDC, TranslateColor(vbButtonShadow)
      End If
      ' Draw the line number:
      DrawText lhDC, CStr(lLine + 1), -1, tR, DT_RIGHT
      
      ' Increment the line:
      lLine = lLine + 1
      ' Increment the position:
      OffsetRect tR, 0, lLineHeight
      If (tR.Bottom > lEnd) Or (lLine + 1 > lCount) Then
         bComplete = True
      End If
   Loop While Not bComplete
   
   ' Draw a line...
   MoveToEx lhDC, tR.Right + 1, 0, tPO
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbButtonShadow))
   hPenOld = SelectObject(lhDC, hPen)
   LineTo lhDC, tR.Right + 1, lEnd
   SelectObject lhDC, hPenOld
   DeleteObject hPen
   If picTo.AutoRedraw Then
      picTo.Refresh
   End If
   
End Sub
