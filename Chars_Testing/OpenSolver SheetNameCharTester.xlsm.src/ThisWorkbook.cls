VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    If MsgBox("Export VBA Code?", vbYesNo, "Before you go") = vbYes Then
        ExportVBACode
    End If

End Sub

Public Sub ExportVBACode()

    With New CodeExporter
        Set .SourceWorkbook = ThisWorkbook
        .ExportTo exportPath
    End With

End Sub

Private Function exportPath() As String

    exportPath = ThisWorkbook.FullName & ".src"
    If Not FileOrDirExists(exportPath) Then
        ' Modification JWD: We need late binding on FSO here to avoid a dependence on Microsoft Scripting Runtime
        Dim fso As Object ' FileSystemObject
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.GetFolder(ThisWorkbook.Path).SubFolders.Add (ThisWorkbook.Name & ".src")
    End If
        

End Function
'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================
Private Sub Workbook_Open()
    FixReferences
End Sub

Private Sub FixReferences()
    Dim VBProj As VBIDE.VBProject, chkRef As VBIDE.Reference
    Set VBProj = ThisWorkbook.VBProject
    
    ' Remove any OpenSolver reference
    For Each chkRef In VBProj.References
        ' Debug.Print chkRef.Name
        If InStr(chkRef.Name, "OpenSolver") Then
            VBProj.References.Remove chkRef
        End If
    Next chkRef
    
    ' Add the reference back
    ' Assumes OpenSolver and OpenSolverTester folders are in the same root directory
    Dim OpenSolverPath As String
    OpenSolverPath = ThisWorkbook.Path & Application.PathSeparator & _
                     ".." & Application.PathSeparator & _
                     "OpenSolver" & Application.PathSeparator & _
                     "OpenSolver.xlam"
    On Error GoTo NotFound
    VBProj.References.AddFromFile OpenSolverPath
    Exit Sub
    
NotFound:
    MsgBox "Unable to add reference to OpenSolver.xlam at:" & vbNewLine & _
           OpenSolverPath & vbNewLine & vbNewLine & _
           "Make sure the OpenSolver and OpenSolverTester folders are in the same root directory."
End Sub
