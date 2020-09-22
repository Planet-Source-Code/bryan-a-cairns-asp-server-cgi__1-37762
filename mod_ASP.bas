Attribute VB_Name = "mod_ASP"
Public fso As New FileSystemObject
Public LocalRoot As String
Private Const StartTag = "<%"
Private Const EndTag = "%>"

Public Function ServASPPage(sCommand As String, sParams As String, sHeaders As String, sIP As String, sLocalPath As String, sWebPath As String)
On Error GoTo EH

Dim Response As cASPResponse
Dim Server As cASPServer
Dim Request As cASPRequest
Dim cScript As cScripter
Dim cScritCont As cScriptControl
Set cScritCont = New cScriptControl
Set Response = New cASPResponse
Set Server = New cASPServer
Set Request = New cASPRequest
Set cScript = New cScripter
Set cScritCont.cScript = Form1.ScriptControl1

Response.StartResponse

'Parse the pages
ParseASPPage sLocalPath, Response, Request, Server, cScript
cScript.ScriptCode = FixScriptCode(cScript.ScriptCode)

'dump to a file for testing
WriteTextFile App.Path & "\asp\data.txt", cScript.ScriptCode

'After all Parsing, shove the buffer into the script control and try to run it
cScritCont.StartScript Response, Request, Server, cScript.ScriptCode
ServASPPage = Response.ReturnText

Response.StartResponse
cScritCont.cScript.Reset
Exit Function
EH:
    ServASPPage = "<BR>500 ASP Server Error<BR>" & Err.Number & " - " & Err.Description
Exit Function
End Function

Public Function FixScriptCode(sCode As String) As String
'add all variables to the beginning
Dim Ipos As Single
Dim EPos As Single

Dim sFinal As String
Dim sTMP As String
Ipos = InStr(1, sCode, "Dim")
If Ipos = 0 Then
FixScriptCode = sCode
Exit Function
End If
EPos = 1
Do
Ipos = InStr(EPos, sCode, vbCrLf)
If Ipos = 0 Then Exit Do

'get the text line
sTMP = Mid(sCode, EPos, Ipos - EPos)
EPos = Ipos + Len(vbCrLf)

If InStr(1, sTMP, "Dim") <> 0 Then
    sFinal = sTMP & vbCrLf & sFinal
Else
    sFinal = sFinal & sTMP & vbCrLf
End If
Loop
FixScriptCode = sFinal
End Function

Private Function GetAllIncludeFiles(sTXT As String) As String
On Error GoTo EH
Dim Ipos As Long
Dim EPos As Long
Dim sTMP As String
Dim sFile As String
Dim Icounter As Integer
'Looking for... <!-- #include file = "filename.asp" -->'
Ipos = 0
EPos = 0

Do
Icounter = Icounter + 1
If Icounter = 100 Then Exit Do
If Ipos > Len(sTXT) Then Exit Do

Ipos = InStr(Ipos + 1, sTXT, "<!-- #")
If Ipos <> 0 Then
Ipos = Ipos + 1
    EPos = InStr(Ipos, sTXT, "-->")
    If EPos <> 0 Then
    'we found one
    sFile = GetQoutedText(Mid(sTXT, Ipos, EPos - Ipos))
    If sFile <> "" Then
        sFile = Replace(sFile, "/", "\")
        sTMP = sTMP & sFile & vbCrLf
    End If
    End If
Else
 Exit Do
End If
Loop

GetAllIncludeFiles = sTMP
Exit Function
EH:
GetAllIncludeFiles = ""
Exit Function
End Function

Private Function GetQoutedText(sTXT As String) As String
On Error GoTo EH
Dim Ipos As Long
Dim EPos As Long
Dim sTMP As String
Ipos = InStr(1, sTXT, Chr(34))
If Ipos <> 0 Then
Ipos = Ipos + 1
    EPos = InStr(Ipos, sTXT, Chr(34))
    If EPos <> 0 Then
        sTMP = Mid(sTXT, Ipos, EPos - Ipos)
    End If
End If
GetQoutedText = sTMP
Exit Function
EH:
GetQoutedText = ""
Exit Function
End Function

Private Sub ParseASPPage(sStartFile As String, Response As cASPResponse, Request As cASPRequest, Server As cASPServer, cScript As cScripter)
Dim sFileData As String
Dim sTMP() As String
Dim sIncludes() As String
Dim sCurrentFile As String
Dim I As Integer
Dim H As Integer
Dim J As Integer
Dim bFound As Boolean

On Error GoTo EH
'load the inital include files
sFileData = OpenTextFile(sStartFile)
sIncludes = Split(GetAllIncludeFiles(sFileData))
sTMP = sIncludes

'remove all references to include files
sFileData = ReturnPageData(sFileData)

'add the code to the script buffer
'cScript.AddCode sFileData

'load all include files
If UBound(sTMP) > -1 Then
    Do
    sCurrentFile = Replace(LocalRoot & sTMP(I), vbCrLf, "")
    If CheckFile(sCurrentFile) = False Then
        Response.sWrite "<BR>500 ASP Server Error<BR>Include file not found: " & sTMP(I)
        Exit Sub
    End If
        sFileData = OpenTextFile(sCurrentFile)
        sIncludes = Split(GetAllIncludeFiles(sFileData))
        'remove all references to include files
        sFileData = ReturnPageData(sFileData)
        'add the code to the script buffer
        cScript.AddCode sFileData
        
        
            If UBound(sIncludes) > -1 Then
            'add them to sTMP if not already there
            bFound = False
            For H = LBound(sIncludes) To UBound(sIncludes)
                For J = LBound(sTMP) To UBound(sTMP)
                    If LCase(sTMP(J)) = LCase(sIncludes(H)) Then
                        bFound = True
                        Exit For
                    End If
                Next J
                If bFound = False Then
                    ReDim Preserve sTMP(UBound(sTMP) + 1)
                    sTMP(UBound(sTMP)) = sIncludes(H)
                End If
            Next H
            End If
        If I = UBound(sTMP) Then Exit Do
        I = I + 1
    Loop
End If

'load the start page last
sFileData = OpenTextFile(sStartFile)
sFileData = ReturnPageData(sFileData)
cScript.AddCode sFileData

'cscript.ScriptCode is now the raw code of all files needed

FixPagesCode sStartFile, Response, Request, Server, cScript
cScript.ScriptCode = Replace(cScript.ScriptCode, "response.swrite(" & Chr(34) & Chr(34) & ")", "")
Exit Sub
EH:
Response.sWrite "<BR>500 ASP Server Error<BR>" & Err.Number & " - " & Err.Description
Exit Sub
End Sub

Private Sub FixPagesCode(sStartFile As String, Response As cASPResponse, Request As cASPRequest, Server As cASPServer, cScript As cScripter)
'now we have to parse the cScript.ScriptCode so it will go into
'The MS Script Control without errors
'Basically convert all text not inside of <% %> to Response.Write ""
'Then we have to replace certain variables so they work in our system
'For instance we can not use Response.Write, so we use Response.sWrite
On Error GoTo EH
Dim Ipos As Single
Dim EPos As Single
Dim sTMP As String
Dim CurrentPos As Single
Dim sFinished As String
Dim bNoFormat As Boolean
EPos = 1
CurrentPos = 1
bNoFormat = False
'parse and fix the code, shove it back into the cScript.ScriptCode
Do
'find the next start tag
Ipos = InStr(EPos, cScript.ScriptCode, StartTag)
If Ipos = 0 Then
'shove the entire thing into the reponse buffer
bNoFormat = True
Exit Do
End If

'find the next end tag
EPos = InStr(Ipos, cScript.ScriptCode, EndTag)
If EPos = 0 Then Exit Do

'we have a start and stop tag position...
'get the code outside the tag
sTMP = Mid(cScript.ScriptCode, CurrentPos, Ipos - CurrentPos)
'code outside the tags just gets shoved into the response.swrite buffer

sTMP = Replace(sTMP, Chr(34), Chr(34) & " & chr(34) & " & Chr(34))
sTMP = RemoveVBCRLF(sTMP)
sFinished = sFinished & "Response.Write(" & Chr(34) & sTMP & Chr(34) & ")" & vbCrLf


'get the code in the tag
Ipos = Ipos + Len(StartTag)
sTMP = Mid(cScript.ScriptCode, Ipos, EPos - Ipos)
sFinished = sFinished & sTMP & vbCrLf

'Reset the positions
EPos = EPos + Len(EndTag)
CurrentPos = EPos
Loop

'do any extra formatting we need...
'If bNoFormat = False Then
'change Reponse.Write to our class format: Reponse.sWrite
sFinished = Replace(sFinished, "response.write", "response.swrite", , , vbTextCompare)
'change Server.CreateObject to our class format: Server.sCreateObject
sFinished = Replace(sFinished, "server.createobject", "server.screateobject", , , vbTextCompare)
'remove any reference to include files <!-- #include file = "file.asp" -->
'End If

cScript.ScriptCode = sFinished
Exit Sub
EH:
Response.sWrite "<BR>500 ASP Server Error<BR>" & Err.Number & " - " & Err.Description
Exit Sub
End Sub

Private Function RemoveVBCRLF(sTXT As String) As String
'Removes the beginning vbcrlf and ending vbcrlf
Dim sTMP As String
sTMP = sTXT

If Left(sTMP, 2) = vbCrLf Then
    sTMP = Mid(sTMP, 3, Len(sTMP))
End If

If Right(sTMP, 2) = vbCrLf Then
    sTMP = Mid(sTMP, 1, Len(sTMP) - 2)
End If

sTMP = Replace(sTMP, vbCrLf, Chr(34) & " & vbcrlf & " & Chr(34))
RemoveVBCRLF = sTMP
End Function

Private Function ReturnPageData(sTXT As String) As String
'Find and remove references to include files
'<!-- #include file = "setup.asp" -->
On Error GoTo EH:
Dim Ipos As Single
Dim EPos As Single
Dim LastPos As Single

EPos = 1
LastPos = 0
Do
Ipos = InStr(EPos, sTXT, "#include")
If Ipos = 0 Then Exit Do
EPos = InStr(Ipos, sTXT, "-->")
If EPos = 0 Then Exit Do
LastPos = EPos + Len("-->")
Loop

If LastPos = 0 Then
    ReturnPageData = sTXT
Else
    ReturnPageData = Mid(sTXT, LastPos, Len(sTXT))
End If

Exit Function
EH:
ReturnPageData = "<BR>500 ASP Server Error<BR>" & Err.Number & " - " & Err.Description
Exit Function
End Function
