VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
' Multiple Undo Class for TextBoxes and RichTextBoxes v2.03
' Copyright ©2002 by Sebastian Thomschke, All Rights Reserved.
' http://www.sebthom.de
'*********************************************************************
' If you like this code, please vote for it at Planet-Source-Code.com:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34335&lngWId=1
' Thank you
'*********************************************************************
' You are free to use this code within your own applications, but you
' are expressly forbidden from selling or otherwise distributing this
' source code without prior written consent.
'*********************************************************************
' Thanks to MrBobo for his suggestions for improvement
'*********************************************************************
Option Explicit

Public Enum undoActionType
   added
   Modified
   deleted
   none
End Enum


'*********************************************************************
' Private API Declarations
'*********************************************************************
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function InvalidateClientRect Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Private Const WM_SETREDRAW = &HB
Private Const EM_LINESCROLL = &HB6
Private Const EM_GETFIRSTVISIBLELINE = &HCE


'*********************************************************************
' Private Vars
'*********************************************************************
Private WithEvents MyRTB As RichTextBox
Attribute MyRTB.VB_VarHelpID = -1
Private WithEvents MyTB As TextBox
Attribute MyTB.VB_VarHelpID = -1
Private WithEvents MyTimer As VB.Timer
Attribute MyTimer.VB_VarHelpID = -1

Private MyTextControl As Control

Private TrackUndo As Boolean
Private UndoStack As New Collection
Private RedoStack As New Collection

Private SelStart_BeforeChange As Long
Private SelLength_BeforeChange As Long
Private SelStart_AfterChange As Long
Private SelLength_AfterChange As Long

Private Text_BeforeChange As String



'*********************************************************************
' Friend Events
'*********************************************************************
Public Event StateChanged()


'*********************************************************************
' Friend Properties
'*********************************************************************
Friend Property Get canUndo() As Boolean
   canUndo = UndoStack.Count > 1 And TrackUndo
End Property

Friend Property Get canRedo() As Boolean
   canRedo = RedoStack.Count > 0 And TrackUndo
End Property

Friend Property Let IgnoreChange(ByVal IIgnoreChange As Boolean)
   ' turns the tracking of changes on or off
   
   ' when tracking is enabled again, we create an undo item to not get out of sync
   ' TakeUndoSnapShot is intelligent enough not to create an undoitem if no changes took part since the turn off
   If Not IIgnoreChange And Not TrackUndo Then TakeUndoSnapShot
   
   TrackUndo = Not IIgnoreChange
   
   RaiseEvent StateChanged
End Property

Friend Property Get IgnoreChange() As Boolean
   IgnoreChange = Not TrackUndo
End Property

Friend Property Get getUndoCount() As Long
   getUndoCount = UndoStack.Count - 1
End Property

Friend Property Get getRedoCount() As Long
   getRedoCount = RedoStack.Count
End Property

Friend Property Get getUndoActionType(undoStep As Long) As undoActionType
   If UndoStack.Count > 1 And undoStep > 0 And undoStep < UndoStack.Count Then
      If Len(UndoStack.Item(undoStep + 1).TextBlock_BeforeChange) > 0 And Len(UndoStack.Item(undoStep + 1).TextBlock_AfterChange) = 0 Then
         getUndoActionType = deleted
      ElseIf Len(UndoStack.Item(undoStep + 1).TextBlock_BeforeChange) = 0 And (UndoStack.Item(undoStep + 1).TextBlock_AfterChange > 0) Then
         getUndoActionType = added
      Else
         getUndoActionType = Modified
      End If
   Else
      getUndoActionType = none
   End If
End Property

Friend Property Get getUndoActionTime(undoStep As Long) As Date
   If UndoStack.Count > 1 And undoStep > 0 And undoStep < UndoStack.Count Then
      getUndoActionTime = UndoStack.Item(undoStep + 1).Time
   Else
      getUndoActionTime = Null
   End If
End Property

Friend Property Get getRedoActionType(redoStep As Long) As undoActionType
   If RedoStack.Count > 0 And redoStep > 0 And redoStep <= RedoStack.Count Then
      If Len(RedoStack.Item(redoStep).TextBlock_BeforeChange) > 0 And Len(RedoStack.Item(redoStep).TextBlock_AfterChange) = 0 Then
         getRedoActionType = deleted
      ElseIf Len(RedoStack.Item(redoStep).TextBlock_BeforeChange) = 0 And Len(RedoStack.Item(redoStep).TextBlock_AfterChange) > 0 Then
         getRedoActionType = added
      Else
         getRedoActionType = Modified
      End If
   Else
      getRedoActionType = none
   End If
End Property

Friend Property Get getRedoActionTime(redoStep As Long) As Date
   If RedoStack.Count > 0 And redoStep > 0 And redoStep <= RedoStack.Count Then
      getRedoActionTime = RedoStack.Item(redoStep).Time
   Else
      getRedoActionTime = Null
   End If
End Property


'*********************************************************************
' Friend Methods
'
' all public methods return True on success and False on failure
'*********************************************************************
Friend Function AssignToRichTextBox(ByRef Controls As Object, ByRef RichTextBox As RichTextBox, Optional Delay As Long = 250) As Boolean
   AssignToRichTextBox = False
   ' this method assigns the tracking capability to the specified richtextbox and resets the undo-/redostack
   ' the controls object is needed to create a timer object on runtime
   Set MyTB = Nothing
   Set MyRTB = RichTextBox
   Set MyTextControl = RichTextBox
   
   AssignToRichTextBox = CreateTimer(Controls, Delay) And Reset
End Function

Friend Function AssignToTextBox(ByRef Controls As Object, ByRef TextBox As TextBox, Optional Delay As Long = 250) As Boolean
   AssignToTextBox = False
   ' this method assigns the tracking capability to the specified textbox and resets the undo-/redostack
   ' the controls object is needed to create a timer object on runtime
   Set MyTB = TextBox
   Set MyRTB = Nothing
   Set MyTextControl = TextBox
   
   AssignToTextBox = CreateTimer(Controls, Delay) And Reset
End Function

Friend Function Reset() As Boolean
   Reset = False
   
   If MyTextControl Is Nothing Then
      Debug.Print "clUndo.Reset() : Can't proceed. No textbox control assigned"
      Exit Function
   End If
   
   Set UndoStack = New Collection
   Set RedoStack = New Collection
   
   If MyTB Is Nothing Then
      Text_BeforeChange = MyTextControl.TextRTF
   Else
      Text_BeforeChange = MyTextControl.Text
   End If
   
   OnSelChange
      
   TrackUndo = True
   
   Reset = TakeUndoSnapShot
   
   RaiseEvent StateChanged
End Function

Friend Function Redo(Optional ByVal RedoSteps As Long = 1) As Boolean
   Redo = False
   
   If MyTextControl Is Nothing Then
      Debug.Print "clUndo.Redo() : Can't proceed. No textbox control assigned"
      Exit Function
   End If
   
   If RedoStack.Count < 1 Or RedoSteps < 1 Then Exit Function
   
   Dim RedoItem As clsUndoElement, i As Long
   
   TrackUndo = False
   
   If RedoSteps > RedoStack.Count Then RedoSteps = RedoStack.Count

   SetRefresh MyTextControl.hWnd, False
   
   For i = 1 To RedoSteps
   
      Set RedoItem = RedoStack.Item(RedoStack.Count)

      ' redo the text changes
      If MyTB Is Nothing Then
         MyTextControl.TextRTF = Left$(MyTextControl.TextRTF, RedoItem.MatchingCharsFromLeft) & RedoItem.TextBlock_AfterChange & Right$(MyTextControl.TextRTF, RedoItem.MatchingCharsFromRight)
         Text_BeforeChange = MyTextControl.TextRTF
      Else
         MyTextControl.Text = Left$(MyTB.Text, RedoItem.MatchingCharsFromLeft) & RedoItem.TextBlock_AfterChange & Right$(MyTB.Text, RedoItem.MatchingCharsFromRight)
         Text_BeforeChange = MyTextControl.Text
      End If
      
      ' Moving the undo item to the undostack
      UndoStack.Add Item:=RedoItem
      RedoStack.Remove RedoStack.Count
      
   Next
   
   SendMessage MyTextControl.hWnd, EM_LINESCROLL, 0&, ByVal (RedoItem.FirstVisibleLine)
   
   MyTextControl.SelStart = RedoItem.SelStart_AfterChange
   MyTextControl.SelLength = RedoItem.SelLength_AfterChange
   
   SetRefresh MyTextControl.hWnd, True
   
   TrackUndo = True
   
   Redo = True
   
   RaiseEvent StateChanged
End Function

Friend Function Undo(Optional ByVal UndoSteps As Long = 1) As Boolean
   Undo = False

   If MyTextControl Is Nothing Then
      Debug.Print "clUndo.Undo() : Can't proceed. No textbox control assigned"
      Exit Function
   End If
   
   If UndoStack.Count < 2 Or UndoSteps < 1 Then Exit Function
   
   TrackUndo = False
   
   If UndoSteps >= UndoStack.Count Then UndoSteps = UndoStack.Count - 1
   
   Dim UndoItem As clsUndoElement, i As Long
   
   SetRefresh MyTextControl.hWnd, False
   
   For i = 1 To UndoSteps

      Set UndoItem = UndoStack.Item(UndoStack.Count)

      ' undo the text changes
      If MyTB Is Nothing Then
         MyTextControl.TextRTF = Left$(MyTextControl.TextRTF, UndoItem.MatchingCharsFromLeft) & UndoItem.TextBlock_BeforeChange & Right$(MyTextControl.TextRTF, UndoItem.MatchingCharsFromRight)
         Text_BeforeChange = MyTextControl.TextRTF
      Else
         MyTextControl.Text = Left$(MyTextControl.Text, UndoItem.MatchingCharsFromLeft) & UndoItem.TextBlock_BeforeChange & Right$(MyTextControl.Text, UndoItem.MatchingCharsFromRight)
         Text_BeforeChange = MyTextControl.Text
      End If
      
      ' Moving the undo item to the redostack
      RedoStack.Add Item:=UndoItem
      UndoStack.Remove UndoStack.Count
      
   Next
   
   SendMessage MyTextControl.hWnd, EM_LINESCROLL, 0&, ByVal (UndoItem.FirstVisibleLine)
   MyTextControl.SelStart = UndoItem.SelStart_BeforeChange
   MyTextControl.SelLength = UndoItem.SelLength_BeforeChange
   
   SetRefresh MyTextControl.hWnd, True
   
   TrackUndo = True
   
   Undo = True
   
   RaiseEvent StateChanged
End Function

Friend Function TakeUndoSnapShot() As Boolean
   TakeUndoSnapShot = False
   
   If MyTextControl Is Nothing Then
      Debug.Print "clUndo.TakeUndoSnapShot() : Can't proceed. No textbox control assigned"
      Exit Function
   End If
   
   ' add the new undo element
   Dim UndoItem As New clsUndoElement, _
       Text_AfterChange As String, _
       Len_Text_AfterChange As Long, _
       Len_Text_BeforeChange As Long
   
   If MyTB Is Nothing Then
      Text_AfterChange = MyTextControl.TextRTF
   Else
      Text_AfterChange = MyTextControl.Text
   End If
   
   Len_Text_AfterChange = Len(Text_AfterChange)
   Len_Text_BeforeChange = Len(Text_BeforeChange)
   
   UndoItem.MatchingCharsFromLeft = getMatchingCharCount( _
      Text_BeforeChange, _
      Text_AfterChange _
   )
   
   UndoItem.MatchingCharsFromRight = getMatchingCharCount( _
      Right$(Text_BeforeChange, Len_Text_BeforeChange - UndoItem.MatchingCharsFromLeft), _
      Right$(Text_AfterChange, Len_Text_AfterChange - UndoItem.MatchingCharsFromLeft), _
      True _
   )
   
   UndoItem.TextBlock_BeforeChange = Mid$( _
      Text_BeforeChange, _
      UndoItem.MatchingCharsFromLeft + 1, _
      Len_Text_BeforeChange - UndoItem.MatchingCharsFromLeft - UndoItem.MatchingCharsFromRight _
   )
   
   UndoItem.TextBlock_AfterChange = Mid$( _
      Text_AfterChange, _
      UndoItem.MatchingCharsFromLeft + 1, _
      Len_Text_AfterChange - UndoItem.MatchingCharsFromLeft - UndoItem.MatchingCharsFromRight _
   )
      
   UndoItem.SelStart_BeforeChange = SelStart_BeforeChange
   UndoItem.SelLength_BeforeChange = SelLength_BeforeChange
   UndoItem.SelStart_AfterChange = MyTextControl.SelStart
   UndoItem.SelLength_AfterChange = MyTextControl.SelLength
   UndoItem.Time = Now
   
   'Exit if no change took place
   If UndoStack.Count > 0 And Len_Text_BeforeChange = Len_Text_AfterChange And UndoItem.MatchingCharsFromLeft = Len_Text_BeforeChange Then Exit Function
   
   UndoItem.FirstVisibleLine = SendMessage(MyTextControl.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
   
   ' add the undoitem to the undostack
   UndoStack.Add Item:=UndoItem
            
   ' remove redo elements
   Set RedoStack = New Collection
   
   ' preparing for next undo snapshot
   Text_BeforeChange = Text_AfterChange
   SelStart_BeforeChange = MyTextControl.SelStart
   SelLength_BeforeChange = MyTextControl.SelLength
   SelStart_AfterChange = MyTextControl.SelStart
   SelLength_AfterChange = MyTextControl.SelLength
      
   TakeUndoSnapShot = True
      
   RaiseEvent StateChanged
End Function



'*********************************************************************
' Private Common Textbox Events
'*********************************************************************
Private Sub OnChange()
   If TrackUndo Then
      ' restarting the timer by switching it off and on
      MyTimer.enabled = False
      MyTimer.enabled = True
   Else
      MyTimer.enabled = False
   End If
End Sub

Private Sub OnMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' a mouse click will break the delay and force an undosnapshot if text has changed
   If MyTimer.enabled = True Then Call MyTimer_Timer
End Sub

Private Sub OnKeyDown(KeyCode As Integer, Shift As Integer)
   Dim CtrlOnly As Boolean

   CtrlOnly = (Shift = vbCtrlMask)
   
   'CTRL+Z = Undo
   If CtrlOnly And (KeyCode = vbKeyZ) Then
      
      Undo
      KeyCode = 0
      
   'CTRL+Y = Redo
   ElseIf CtrlOnly And (KeyCode = vbKeyY) Then
      
      Redo
      KeyCode = 0
      
   ' navigating the text caret and pressing the spacebar will break the delay and force an undosnapshot, but only if the text has changed
   ElseIf _
      (KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12) Or _
      (KeyCode = vbKeyReturn) Or _
      (KeyCode >= vbKeySpace And KeyCode <= vbKeySnapshot) Then
      
      ' MyTimer.Enabled = True indicates that the text has been changed
      If MyTimer.enabled = True Then Call MyTimer_Timer
      
   End If
End Sub

Private Sub OnSelChange()
   If Not MyTimer.enabled Then
      SelStart_BeforeChange = SelStart_AfterChange
      SelLength_BeforeChange = SelLength_AfterChange
      SelStart_AfterChange = MyTextControl.SelStart
      SelLength_AfterChange = MyTextControl.SelLength
   End If
End Sub


'*********************************************************************
' MyTB Event Wrapper
'*********************************************************************
Private Sub MyTB_Change()
   Call OnChange
End Sub

Private Sub MyTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Call OnSelChange
   Call OnMouseDown(Button, Shift, x, y)
End Sub

Private Sub MyTB_KeyDown(KeyCode As Integer, Shift As Integer)
   Call OnSelChange
   Call OnKeyDown(KeyCode, Shift)
End Sub



'*********************************************************************
' MyRTB Event Wrapper
'*********************************************************************
Private Sub MyRTB_Change()
   Call OnChange
End Sub

Private Sub MyRTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Call OnMouseDown(Button, Shift, x, y)
End Sub

Private Sub MyRTB_KeyDown(KeyCode As Integer, Shift As Integer)
   Call OnKeyDown(KeyCode, Shift)
End Sub

Private Sub MyRTB_SelChange()
   Call OnSelChange
End Sub



'*********************************************************************
' MyTimer Events
'*********************************************************************
Private Sub MyTimer_Timer()
   MyTimer.enabled = False

   If TrackUndo Then TakeUndoSnapShot
End Sub



'*********************************************************************
' Private Methods
'*********************************************************************
Private Function getMatchingCharCount(ByVal StringA As String, ByVal StringB As String, Optional FromRightToLeft As Boolean = False) As Long
   ' compares two strings from the left to the right and returns the count of characters
   ' upto which both strings are identical
   ' for performance reasons a binary search like comparison algorithm is used
   ' (c) 2002 Sebastian Thomschke
   
   Dim Low As Long, High As Long, Middle As Long, isMatching As Boolean
   Dim MatchingLen As Long, LenStringA As Long, LenStringB As Long
   
   Low = 1
   High = IIf(Len(StringA) < Len(StringB), Len(StringA), Len(StringB))
   If High = 0 Then GoTo getMatchingCharCount_FoundIt
   
   LenStringA = Len(StringA)
   LenStringB = Len(StringB)
   
   getMatchingCharCount = 0
   Do
      Middle = Low + ((High - Low) \ 2)
      If FromRightToLeft Then
         isMatching = (Mid$(StringA, LenStringA - Middle + 1, Middle - MatchingLen) = Mid$(StringB, LenStringB - Middle + 1, Middle - MatchingLen))
      Else
         isMatching = (Mid$(StringA, MatchingLen + 1, Middle - MatchingLen) = Mid$(StringB, MatchingLen + 1, Middle - MatchingLen))
      End If
      If isMatching Then
         ' match is in upper half
         MatchingLen = Middle
         Low = Middle + 1
         If Low > High Then GoTo getMatchingCharCount_FoundIt
      Else
         ' match is in lower half
         High = Middle - 1
         If High = 0 Then   ' No match found
            GoTo getMatchingCharCount_FoundIt
         ElseIf Low > High Then
            GoTo getMatchingCharCount_FoundIt
         End If
      End If
   Loop
   Exit Function
   
getMatchingCharCount_FoundIt:
   getMatchingCharCount = High
   Exit Function
End Function

Private Sub SetRefresh(ByVal hWnd As Long, ByVal EnableRefresh As Boolean)
   ' enables or disables the automatic refresh for a control
   If EnableRefresh Then
      SendMessage hWnd, WM_SETREDRAW, 1, 0
      InvalidateClientRect hWnd, 0, 1
   Else
      SendMessage hWnd, WM_SETREDRAW, 0, 0
   End If
End Sub

Private Function CreateTimer(ByRef Controls As Object, Optional Delay As Long = 250) As Boolean
   On Error GoTo CreateTimer_Error

   CreateTimer = False
   
   ' the timer is used to avoid manual undoing of every single character
   ' instead inputs made within a defined time are grouped to one undo action
   If MyTimer Is Nothing Then Set MyTimer = Controls.Add("VB.Timer", "UndoTimer_" & MyTextControl.Name)
   MyTimer.enabled = False
   MyTimer.Interval = Delay
   
   CreateTimer = True
   Exit Function
   
CreateTimer_Error:
   MsgBox _
      Err.Description, _
      vbOKOnly + vbExclamation + vbMsgBoxHelpButton, _
      Err.Source & " - Unexpected Error No. " & Err.Number, _
      Err.HelpFile, _
      Err.HelpContext
   Exit Function
End Function
