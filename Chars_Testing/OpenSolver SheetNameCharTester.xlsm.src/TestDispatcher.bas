Attribute VB_Name = "TestDispatcher"
Option Explicit

Function ApiTest(SheetName As String, Solver As String) As TestResult
    #If Mac Then
        If Solver = "NOMAD" Then
            ApiTest = NA
            Exit Function
        End If
    #End If
    
    Select Case SheetName
    Case "SimpleLP":                    ApiTest = SimpleLP.Test(Solver)
    Case "SimpleIP":                    ApiTest = SimpleIP.Test(Solver)
    Case "BadName!":                    ApiTest = BadName.Test(Solver)
    Case "!BadName":                    ApiTest = BadName1.Test(Solver)
    Case "Bad!Name":                    ApiTest = BadName2.Test(Solver)
    Case "@BadName":                    ApiTest = BadName3.Test(Solver)
    Case "Bad@Name":                    ApiTest = BadName4.Test(Solver)
    Case "BadName@":                    ApiTest = BadName5.Test(Solver)
    Case "EscapeSheetName(1)+2-1":      ApiTest = EscapeSheetName.Test(Solver)
    Case Else:                          ApiTest = NA
    End Select
End Function


