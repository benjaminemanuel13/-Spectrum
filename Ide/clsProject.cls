VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim vName, vPath, vFileName As String
Dim fileNone As String

Dim files() As clsFile
Dim fileAdded As Boolean

Public Property Get Name()
    Name = vName
End Property

Public Property Let Name(inName)
    vName = inName
End Property

Public Property Get path()
    path = vPath
End Property

Public Property Let path(inPath)
    vPath = inPath
End Property

Public Property Get FileName()
    FileName = vFileName
End Property

Public Property Let FileName(inFileName)
    vFileName = inFileName
End Property

Public Property Get File(inIndex) As clsFile
    Set File = files(inIndex)
End Property

Public Property Get FileCount() As Integer
    FileCount = UBound(files) + 1
End Property

Public Sub AddFile(inFile As clsFile)
    Dim count As Integer
    
    If fileAdded Then
            count = UBound(files)
        Else
            count = 0
            fileAdded = True
    End If
    
    ReDim Preserve files(count + 1)
    
    Set files(count) = inFile
End Sub

Private Sub Class_Initialize()
    fileAdded = False
    fileNone = "temp"
End Sub

Public Sub Compile()
    CreateAsmFile
    RunRasm
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Do While fs.FileExists(vPath + "temp\temp.tzx") = False
    Loop
    
    Call FileCopy(vPath + "temp\temp.tzx", vPath + "bin\" + vName + ".tzx")
    Kill (vPath + "temp\temp.asm")
    Kill (vPath + "temp\temp.bin")
    Kill (vPath + "temp\temp.tap")
    Kill (vPath + "temp\temp.tzx")
    Kill (vPath + "bin\main.z80asm")
    
    MsgBox ("Program Compiled")
End Sub

Private Sub CreateAsmFile()
    Dim max, count As Integer
    max = UBound(files) - 1
    
    Open vPath + "\bin\main.z80asm" For Output As #1
    
    For count = 0 To max
        Dim thisFile As clsFile
        Set thisFile = files(count)
        
        Print #1, thisFile.Code
    Next
    
    Close #1
    
    Call FileCopy(vPath + "bin\main.z80asm", vPath + "temp\temp.asm")
End Sub

Private Sub RunRasm()
    Call ChDir(vPath + "temp")

    Dim asmFilename As String
    asmFilename = fileNone + ".asm"
    
    Dim binFilename As String
    binFilename = fileNone + ".bin"
    
    Dim tapFilename As String
    tapFilename = fileNone + ".tap"
    
    Dim tzxFilename As String
    tzxFilename = fileNone + ".tzx"
    
    Dim runit As String
    runit = Chr$(34) + rasmPath + "compile.bat" + Chr$(34) + " " + vPath + "temp" + " " + asmFilename + " " + binFilename + " " + tapFilename + " " + tzxFilename
    
    Call Shell(runit, vbMaximizedFocus)
End Sub
