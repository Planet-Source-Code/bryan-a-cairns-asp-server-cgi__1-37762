Attribute VB_Name = "mod_FileAccess"

Public Sub RemoveFile(sFile)
On Error Resume Next
Kill sFile
End Sub

Public Function OpenTextFile(sFile As String) As String
'Reads an entire file into a string
On Error GoTo EH
Dim TMPTXT As String
Dim FinTxt As String
Dim iFile As Integer
iFile = FreeFile
Open sFile For Binary Access Read As #iFile
TMPTXT = Space$(LOF(iFile))
Get #iFile, , TMPTXT
Close #iFile
OpenTextFile = TMPTXT
Exit Function
EH:
OpenTextFile = ""
Exit Function
End Function

Public Function WriteTextFile(sFile As String, sData As String) As Boolean
On Error GoTo EH
Dim iFile As Integer
If CheckFile(sFile) = True Then
Kill sFile
End If
iFile = FreeFile

Open sFile For Binary Access Write As #iFile
Put #iFile, 1, sData
Close #iFile

WriteTextFile = True
Exit Function
EH:
WriteTextFile = False
Exit Function
End Function

Public Function ParsePath(ByVal TempPath As String, ReturnType As Integer)
'Parses a filename path
'Returns:
'Drive
'Directory
'Filename
'Extention

    Dim DriveLetter As String
    Dim DirPath As String
    Dim fname As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean

    If ReturnType <> 0 And ReturnType <> 1 And ReturnType <> 2 And ReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If

        DriveLetter = ""
        DirPath = ""
        fname = ""
        Extension = ""

        If Mid(TempPath, 2, 1) = ":" Then ' Find the drive letter.
            DriveLetter = Left(TempPath, 2)
            TempPath = Mid(TempPath, 3)
        End If

            PathLength = Len(TempPath)

            For Offset = PathLength To 1 Step -1 ' Find the next delimiter.
                Select Case Mid(TempPath, Offset, 1)
                 Case ".": ' This indicates either an extension or a . or a ..
                 ThisLength = Len(TempPath) - Offset

                 If ThisLength >= 1 Then ' Extension
                     Extension = Mid(TempPath, Offset, ThisLength + 1)
                 End If

                     TempPath = Left(TempPath, Offset - 1)
                     Case "\": ' This indicates a path delimiter.
                     ThisLength = Len(TempPath) - Offset

                     If ThisLength >= 1 Then ' Filename
                         fname = Mid(TempPath, Offset + 1, ThisLength)
                         TempPath = Left(TempPath, Offset)
                         FileNameFound = True
                         Exit For
                     End If

                         Case Else
                    End Select

                    Next Offset


                        If FileNameFound = False Then
                            fname = TempPath
                        Else
                            DirPath = TempPath
                        End If


                            If ReturnType = 0 Then
                                ParsePath = DriveLetter
                            ElseIf ReturnType = 1 Then
                                ParsePath = DirPath
                            ElseIf ReturnType = 2 Then
                                ParsePath = fname
                            ElseIf ReturnType = 3 Then
                                ParsePath = Extension
                            End If

End Function
Public Sub CheckTMPDir(sDir As String, dKill As Boolean)
'Check the tmp dir - creates as needed
On Error Resume Next
Dim Iret
Iret = Dir(sDir, vbDirectory)
If Iret > "" And dKill = True Then
RmTree sDir
MkDir sDir
Else
If Iret = "" Then
MkDir sDir
End If
End If

End Sub

Public Function CheckFile(sFile As String) As Boolean
'Does a file exist TRUE / FALSE
On Error Resume Next
If sFile = "" Then
CheckFile = False
Exit Function
End If
Dim Iret
Iret = Dir(sFile)
If Iret > "" Then
CheckFile = True
Else
If Iret = "" Then
CheckFile = False
End If
End If

End Function
Public Sub RmTree(ByVal vDir As Variant)
'Removes a Directory structor
On Error Resume Next
Dim vFile As Variant
    ' Check if "\" was placed at end
    ' If So, Remove it
If Right(vDir, 1) = "\" Then
        vDir = Left(vDir, Len(vDir) - 1)
    End If
' Check if Directory is Valid
    ' If Not, Exit Sub
    vFile = Dir(vDir, vbDirectory)
If vFile = "" Then
        Exit Sub
    End If
' Search For First File
    vFile = Dir(vDir & "\", vbDirectory)
    ' Loop Until All Files and Directories
    ' Have been Deleted
Do Until vFile = ""


        If vFile = "." Or vFile = ".." Then
            vFile = Dir
        ElseIf (GetAttr(vDir & "\" & vFile) And _
            vbDirectory) = vbDirectory Then
            RmTree vDir & "\" & vFile
            vFile = Dir(vDir & "\", vbDirectory)
        Else
            Kill vDir & "\" & vFile
            vFile = Dir
        End If


    Loop


    ' Remove Top Most Directory
    RmDir vDir
End Sub

Public Function WinPathtoDOS(sPath As String) As String
'converts windows path to DOS path
Dim iDir() As String
Dim sDir As String
Dim iCount As Integer

sDir = ""
iDir = Split(sPath, "\", , vbBinaryCompare)

For iCount = 1 To UBound(iDir)
    If Len(iDir(iCount)) > 8 Then
        If iCount <> UBound(iDir) Then
            iDir(iCount) = Mid(iDir(iCount), 1, 6) & "~1"
        Else
            If Len(iDir(iCount)) > 12 Then
                iDir(iCount) = Mid(iDir(iCount), 1, 6) & "~1" & Mid(iDir(iCount), Len(iDir(iCount)) - 3)
            End If
        End If
    End If
Next iCount

For iCount = 0 To UBound(iDir)
    If iCount <> UBound(iDir) Then
        sDir = sDir & iDir(iCount) & "\"
    Else
        sDir = sDir & iDir(iCount)
    End If
Next iCount

WinPathtoDOS = sDir

End Function



