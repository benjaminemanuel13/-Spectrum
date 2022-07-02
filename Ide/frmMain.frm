VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   10560
      TabIndex        =   12
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   10560
      TabIndex        =   11
      Top             =   6720
      Width           =   1935
   End
   Begin VB.ListBox lstImages 
      Height          =   2205
      Left            =   10560
      TabIndex        =   10
      Top             =   720
      Width           =   1935
   End
   Begin VB.ListBox lstLibraries 
      Height          =   2400
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ListBox lstIncludes 
      Height          =   2205
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ListBox lstFiles 
      Height          =   1035
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   7935
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   495
      Left            =   10560
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Images"
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Library Files"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Included Files"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Code"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Files"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mains() As clsFile
Private startedMains As Boolean

Private includes() As clsFile
Private startedIncludes As Boolean

Private libraries() As clsFile
Private startedLibraries As Boolean

Private Sub cmdClose_Click()
    Unload mainForm
End Sub

Private Sub cmdCompile_Click()
    currentProject.Compile
End Sub

Public Sub AddFile(inType As FileType, inFile As clsFile)
    Dim count As Integer
    
    If inType = MainFile Then
        If startedMains Then
            count = UBound(mains) + 1
        Else
            count = 1
            startedMains = True
        End If
        ReDim Preserve mains(count)
        Set mains(count - 1) = inFile
    ElseIf inType = Include Then
        If startedIncludes Then
            count = UBound(includes) + 1
        Else
            count = 1
            startedIncludes = True
        End If
        ReDim Preserve includes(count)
        Set includes(count - 1) = inFile
    ElseIf inType = Library Then
        If startedLibraries Then
            count = UBound(libraries) + 1
        Else
            count = 1
            startedLibraries = True
        End If
        ReDim Preserve libraries(count)
        Set libraries(count - 1) = inFile
    ElseIf inType = Image Then
    
    End If
    
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub lstFiles_Click()
    Dim File As clsFile
    Set File = mains(lstFiles.ListIndex)
    
    txtCode.Text = File.Code
End Sub

Private Sub lstIncludes_Click()
    Dim File As clsFile
    Set File = includes(lstIncludes.ListIndex)
    
    txtCode.Text = File.Code
End Sub

Private Sub lstLibraries_Click()
    Dim File As clsFile
    Set File = libraries(lstLibraries.ListIndex)
    
    txtCode.Text = File.Code
End Sub
