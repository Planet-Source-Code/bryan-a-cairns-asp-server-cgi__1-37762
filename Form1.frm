VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TestValue 
      Caption         =   "Test Parser"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4680
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Output File:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Root Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim cPlug As psplugin
Set cPlug = New psplugin
LocalRoot = Text1.Text
WriteTextFile Text2.Text, cPlug.Command_Get("GET", "", "", "127.0.0.1", "F:\Server\PluginTemplates\ASP\Examples\Test\ASP\default.asp", "/default.asp")
End Sub

Private Sub Form_Load()
Text1.Text = App.Path & "\ASP\"
Text2.Text = App.Path & "\ASP\test.html"
'Command1_Click
'DoEvents
'Unload Me
End Sub

Private Sub TestValue_Click()
Dim cRequest As cASPRequest
Dim sTMP As String
Dim I As Integer

Set cRequest = New cASPRequest

sTMP = "a=1" & vbCrLf
sTMP = sTMP & "b=2" & vbCrLf
sTMP = sTMP & "c=3" & vbCrLf
sTMP = sTMP & "d=4" & vbCrLf
cRequest.ClearAll
cRequest.AddValues sTMP, cRequest.Querystring, vbCrLf, "="

For I = 0 To cRequest.Querystring.Count - 1
MsgBox cRequest.Querystring.Keys(I) & " = " & cRequest.Querystring.Items(I)
Next I

MsgBox cRequest.Querystring("a")
End Sub
