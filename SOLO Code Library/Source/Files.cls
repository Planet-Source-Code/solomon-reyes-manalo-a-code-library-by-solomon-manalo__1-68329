VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSysObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Original Coding by: Solomon Manalo

Private Declare Function StrFormatByteSize Lib _
   "shlwapi" Alias "StrFormatByteSizeA" (ByVal _
   dw As Long, ByVal pszBuf As String, ByRef _
   cchBuf As Long) As String
   
Enum FilesAttributes
      isAlias = vbAlias
      isNormal = vbNormal
      isArchive = vbArchive
      isHidden = vbHidden
      isReadOnly = vbReadOnly
      isSystem = vbSystem
      isDirectory = vbDirectory
End Enum

Enum FileAttrChangeTo
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

'==========================================================
Public Function GetFileName(filepath As String) As String
Dim I As Integer
Dim rev As String

rev = StrReverse(filepath)
I = InStr(rev, "\")
If I = 0 Then
   GetFileName = filepath
Else
   GetFileName = StrReverse(Left(rev, I - 1))
End If
End Function

Public Function GetFiles(FromFolder As String, Optional Attributes As _
    VbFileAttribute) As String()
    
    Dim result() As String
    Dim FileName As String, count As Long, path2 As String
    Const ALLOC_CHUNK = 50
    ReDim result(0 To ALLOC_CHUNK) As String
    
    FileName = Dir$(FromFolder, Attributes)
    
    Do While Len(FileName)
        count = count + 1
        If count > UBound(result) Then
            ' Resize the result array if necessary.
            ReDim Preserve result(0 To count + ALLOC_CHUNK) As String
        End If
        result(count) = FileName
        ' Get ready for the next iteration.
        FileName = Dir$
    Loop
    ' Trim the result array.
    ReDim Preserve result(0 To count) As String
    GetFiles = result
End Function

Public Function GetFileAttribute(filepath As String) As String
    Dim result As String, Attr As Long
    Attr = GetAttr(filepath)
    ' GetAttr also works with directories.
    If Attr And vbDirectory Then result = result & " Directory"
    If Attr And vbReadOnly Then result = result & " ReadOnly"
    If Attr And vbHidden Then result = result & " Hidden"
    If Attr And vbSystem Then result = result & " System"
    If Attr And vbArchive Then result = result & " Archive"
    ' Discard the first (extra) space.
    GetFileAttribute = Mid$(result, 2)
End Function

Public Function GetFileSize(filepath As String) As String
Dim Buffer As String
Dim result As String
Dim File_LEN As Long

File_LEN = FileLen(filepath)

   Buffer = Space$(255)
   result = StrFormatByteSize(File_LEN, Buffer, Len(Buffer))
    
   If InStr(result, vbNullChar) > 1 Then
      GetFileSize = Left$(result, InStr(result, vbNullChar) - 1)
   End If
End Function

Public Function GetFileExt(filepath As String) As String
Dim rev, ext As String
Dim stopAt As Integer

rev = StrReverse(GetFileName(filepath))
stopAt = InStr(rev, ".")
ext = Left(rev, stopAt)
GetFileExt = StrReverse(ext)
End Function

Public Function MarkFileAttributes(filepath As String, _
                                   ChangeAttrTO As FileAttrChangeTo) As Boolean
On Error GoTo Erx:
SetAttr filepath, ChangeAttrTO
MarkFileAttributes = True
Exit Function

Erx:
   MarkFileAttributes = False
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RenameFile(filepath As String, NewFileName As String)
Dim Loc, fname, fExt As String

On Error GoTo Erx:
fname = GetFileName(filepath)
Loc = Replace(filepath, fname, "")
fExt = GetFileExt(filepath)

Name filepath As Loc & NewFileName & fExt

RenameFile = True
Exit Function

Erx:
   RenameFile = False
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
