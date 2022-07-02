VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vFileType As FileType
Private vCode, vFilePath, vFileName As String

Public Property Get FileType() As FileType
    FileType = vFileType
End Property

Public Property Let FileType(inFileType As FileType)
    vFileType = inFileType
End Property

Public Property Get FilePath() As String
    FilePath = vFilePath
End Property

Public Property Let FilePath(inFilePath As String)
    vFilePath = inFilePath
End Property

Public Property Get FileName() As String
    FileName = vFileName
End Property

Public Property Let FileName(inFileName As String)
    vFileName = inFileName
End Property

Public Property Get Code() As String
    Code = vCode
End Property

Public Sub LoadCode()
    Open vFilePath + vFileName For Input As #2
    
    Dim LinesFromFile, NextLine As String

    Do Until EOF(2)
        Line Input #2, NextLine
        NextLine = Replace(NextLine, ".", "")
        LinesFromFile = LinesFromFile + NextLine + Chr(13) + Chr(10)
    Loop
    
    vCode = LinesFromFile
    
    Close #2
End Sub
