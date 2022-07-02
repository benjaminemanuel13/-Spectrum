Attribute VB_Name = "modGlobal"
Global projectsForm As New frmProjects
Global mainForm As frmMain

Public Const rasmPath = "D:\Inetpub\ftproot\Spectrum\AsmCompile\"
Global basePath As String

Global currentProject As clsProject

Enum FileType
    MainFile
    Include
    Library
    Image
End Enum

Sub Main()
    basePath = "C:\$Spectrum\"
    projectsForm.Show
End Sub
