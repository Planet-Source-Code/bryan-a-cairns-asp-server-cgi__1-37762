VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cASPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'To Do
'Add Server Variables

Public Function sCreateObject(sTXT As String) As Object
On Error GoTo EH
Set sCreateObject = CreateObject(sTXT)
Exit Function
EH:
Set sCreateObject = Nothing
'sReponse.GetLastError = "Could not Create Object: " & Err.Number & Err.Description
Exit Function
End Function

Public Function MapPath(sPath As String) As String
On Error GoTo EH
Dim sTMP As String
sTMP = Replace(sPath, "/", "\")
    MapPath = fso.BuildPath(fso.GetParentFolderName(LocalRoot), sTMP)
Exit Function
EH:
MapPath = ""
sReponse.GetLastError = "Could not Map Path: " & Err.Number & Err.Description
Exit Function
End Function
