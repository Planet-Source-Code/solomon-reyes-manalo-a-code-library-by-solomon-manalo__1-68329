VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderSysObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Original Coding by: Solomon Manalo

Enum FolderAttributes
      isAlias = vbAlias
      isNormal = vbNormal
      isArchive = vbArchive
      isHidden = vbHidden
      isReadOnly = vbReadOnly
      isSystem = vbSystem
      isDirectory = vbDirectory
End Enum

Enum FolderAttrChangeTo
      [Archive] = vbArchive
      [Archive+ReadOnly] = vbArchive + vbReadOnly
      [Archive+Hidden] = vbArchive + vbHidden
      [Archive+ReadOnly+Hidden] = vbArchive + vbReadOnly + vbHidden
      
      [ReadOnly] = vbReadOnly
      [ReadOnly+Hidden] = vbReadOnly + vbHidden
      [ReadOnly+Archive] = vbReadOnly + vbArchive
      [ReadOnly+Archive+Hidden] = vbReadOnly + vbArchive + vbHidden
      
      [Hidden] = vbHidden
      [Hidden+ReadOnly] = vbHidden + vbReadOnly
      [Hidden+Archive] = vbHidden + vbArchive
      [Hidden+Archive+ReadOnly] = vbHidden + vbArchive + vbReadOnly
End Enum

'============================================================
Public Function GetFolderName(path As String) As String
Dim I As Integer
Dim rev As String

rev = StrReverse(path)
I = InStr(rev, "\")
If I = 0 Then
   GetFolderName = path
Else
   GetFolderName = StrReverse(Left(rev, I - 1))
End If
End Function

Function GetDirectories(path As String, Optional Attributes As _
    VbFileAttribute, Optional IncludePath As Boolean) As String()
    
    Dim result() As String
    Dim dirname As String, count As Long, path2 As String
    Const ALLOC_CHUNK = 50
    ReDim result(ALLOC_CHUNK) As String
    ' Build the path name + backslash.
    path2 = path
    
    If Right$(path2, 1) <> "\" Then path2 = path2 & "\"
    dirname = Dir$(path2 & "*.*", vbDirectory Or Attributes)
    
    Do While Len(dirname)
        If dirname = "." Or dirname = ".." Then
            ' Exclude the "." and ".." entries.
        ElseIf (GetAttr(path2 & dirname) And vbDirectory) = 0 Then
            ' This is a regular file.
        Else
            ' This is a directory.
            count = count + 1
            If count > UBound(result) Then
                ' Resize the result array if necessary.
                ReDim Preserve result(count + ALLOC_CHUNK) As String
            End If
            ' Include the path if requested.
            If IncludePath Then dirname = path2 & dirname
            result(count) = dirname
        End If
        dirname = Dir$
    Loop
    ' Trim the result array.
    ReDim Preserve result(count) As String
    GetDirectories = result
End Function

Public Function RenameFolder(path As String, NewFileName As String)
Dim Loc, fname, fExt As String
On Error GoTo Erx:
fname = GetFolderName(path)
Loc = Replace(path, fname, "")
Name path As Loc & NewFileName
RenameFolder = True
Exit Function
Erx:
   RenameFolder = False
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
 
Public Function MarkFolderAttributes(FolderPath As String, _
                                   ChangeAttrTO As FolderAttrChangeTo) As Boolean
On Error GoTo Erx:
SetAttr FolderPath, ChangeAttrTO
MarkFolderAttributes = True
Exit Function
Erx:
   MarkFolderAttributes = False
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function FolderExists(path As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir$(path & "\nul") <> "")
End Function

