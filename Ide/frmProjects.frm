VERSION 5.00
Begin VB.Form frmProjects 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Available Projects"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewProject 
      Caption         =   "New..."
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteProject 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenProject 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox lstProjects 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Projects"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim projects() As clsProject

Private Sub cmdClose_Click()
    Unload projectsForm
End Sub

Private Sub cmdOpenProject_Click()
    Dim projectIndex As Integer
    projectIndex = lstProjects.ListIndex
    
    Set currentProject = projects(projectIndex)
    
    Set mainForm = New frmMain
    Call LoadFiles(currentProject)
    
    mainForm.Caption = "Project: " + currentProject.Name
    
    PopulateForm
    
    mainForm.Show
    
    Unload projectsForm
End Sub
 
Private Sub PopulateForm()
    Dim count, place As Integer
    count = currentProject.FileCount - 2
    
    For place = 0 To count
        Dim File As clsFile
        Set File = currentProject.File(place)
        
        If File.FileType = MainFile Then
            mainForm.lstFiles.AddItem (File.FileName)
        ElseIf File.FileType = Include Then
            mainForm.lstIncludes.AddItem (File.FileName)
        ElseIf File.FileType = Library Then
            mainForm.lstLibraries.AddItem (File.FileName)
        End If
    Next
End Sub
 
Private Sub LoadFiles(inProject As clsProject)
    Open inProject.path + inProject.FileName For Input As #1
    
    Dim File, FilePath As String
    Dim place As Integer
    place = 0
    
    Do Until EOF(1)
        Line Input #1, File
        Dim firstchar As String
        firstchar = Left(File, 1)
        
        If firstchar = "#" Then
                place = place + 1
            Else
                Dim newFile As clsFile
                Set newFile = New clsFile
                
                If place = 1 Then
                    FilePath = inProject.path
                    newFile.FileType = MainFile
                ElseIf place = 2 Then
                    FilePath = inProject.path + "Include\"
                    newFile.FileType = Include
                ElseIf place = 3 Then
                    FilePath = basePath + "Library\"
                    newFile.FileType = Library
                ElseIf place = 4 Then
                    FilePath = inProject.path + "Images\"
                    newFile.FileType = Image
                End If
                
                newFile.FilePath = FilePath
                newFile.FileName = File
                
                newFile.LoadCode
                
                Call inProject.AddFile(newFile)
                Call mainForm.AddFile(newFile.FileType, newFile)
        End If
    Loop
    
    Close #1
End Sub

Private Sub Form_Load()
    Open basePath + "Projects.txt" For Input As #1
    
    Dim ProjectLine, ProjectName As String
    Dim numberProjects As Integer
    
    Do Until EOF(1)
        Line Input #1, ProjectLine
        
        numberProjects = numberProjects + 1
        ReDim Preserve projects(numberProjects)
        
        Dim newProject As New clsProject
        
        Dim splitted() As String
        splitted = Split(ProjectLine, ",")
        
        Let newProject.Name = splitted(0)
        Let newProject.path = splitted(1)
        Let newProject.FileName = splitted(2)
        
        Set projects(numberProjects - 1) = newProject
        
        lstProjects.AddItem (newProject.Name)
    Loop
    
    Close #1
End Sub

Private Sub lstProjects_Click()
    cmdOpenProject.Enabled = True
    cmdDeleteProject.Enabled = True
End Sub
