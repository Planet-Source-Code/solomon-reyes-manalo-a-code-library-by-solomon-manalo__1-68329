Attribute VB_Name = "modFiles"
Option Explicit
Public Const fExt As String = ".code"


'Original Coding by: Solomon Manalo

Public Function SaveFile(filepath As String, content As String) As Boolean
Dim file_Content As TextStream     'Content Of the file
Dim file_Text As String            'Content Of the file to be spoken!

On Error GoTo Hell:
    file_Text = content

    Set file_Content = fso.CreateTextFile(filepath, True, True)
    file_Content.Write (file_Text)
    file_Content.Close
    SaveFile = True
    Exit Function
Hell:
    SaveFile = False
    MsgBox Err.Description
End Function

Public Function OpenFile(filepath As String) As String
On Error GoTo Hell:
Dim file_Content As TextStream     'Content Of the file
Dim file_Text As String            'Content Of the file to be spoken!
    
    file_Text = ""
    Set file_Content = fso.OpenTextFile(filepath, ForReading, False, TristateMixed)
    While Not file_Content.AtEndOfStream
          file_Text = file_Text & file_Content.ReadAll
    Wend
    OpenFile = file_Text
    file_Content.Close
Hell:
End Function

Public Function FindSecurityFile() As Boolean
Dim p As String
p = fso.GetSpecialFolder(SystemFolder) & "\"
FindSecurityFile = fso.FileExists(p & "Users.SecurityFile")
End Function

Public Function SaveSecurityFile(usr, pss As String) As Boolean
Dim p As String
p = fso.GetSpecialFolder(SystemFolder) & "\"
SaveSecurityFile = SaveFile(p & "Users.SecurityFile", usr & "//" & pss)
End Function

Public Function ReadSecurityFile() As String
Dim p As String
p = fso.GetSpecialFolder(SystemFolder) & "\"
ReadSecurityFile = OpenFile(p & "Users.SecurityFile")
End Function
