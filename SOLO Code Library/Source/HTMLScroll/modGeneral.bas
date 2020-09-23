Attribute VB_Name = "modGeneral"
'*****************************************************************************
'* ZeroIfNegative
'*****************************************************************************
Public Function ZeroIfNegative(ByVal value As Long) As Long
   ' Returns Zero if the value is negative mainly used by usercontrol_resize events
   If value > 0 Then
      ZeroIfNegative = value
   Else
      ZeroIfNegative = 0
   End If
   'alternatively you could write: ZeroIfNegative = IIf(Value > 0, Value, 0)
   'but the "old" if statement is much faster
End Function
