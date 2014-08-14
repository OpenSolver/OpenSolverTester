VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
#If Win32 Then
    Sheet53.ListBoxLinear.Height = 60
    Sheet53.ListBoxLinear.Width = Sheet53.Range("A5").Width - 20
    Sheet53.ListBoxLinear.Left = 10
    Sheet53.ListBoxLinear.Top = Sheet53.Range("A5").Top
    
    Sheet53.ListBoxNonLinear.Height = 80
    Sheet53.ListBoxNonLinear.Width = Sheet53.Range("A11").Width - 20
    Sheet53.ListBoxNonLinear.Left = 10
    Sheet53.ListBoxNonLinear.Top = Sheet53.Range("A11").Top
        
    Sheet53.ListBoxTest.Height = 250
    Sheet53.ListBoxTest.Width = Sheet53.Range("A17").Width - 20
    Sheet53.ListBoxTest.Left = 10
    Sheet53.ListBoxTest.Top = Sheet53.Range("A17").Top
    
    RefreshTestListBox
#End If
End Sub

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
