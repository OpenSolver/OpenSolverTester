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
    Case "#BadName":                    ApiTest = BadName6.Test(Solver)
    Case "Bad#Name":                    ApiTest = BadName7.Test(Solver)
    Case "BadName#":                    ApiTest = BadName8.Test(Solver)
    Case "$BadName":                    ApiTest = BadName9.Test(Solver)
    Case "Bad$Name":                    ApiTest = BadName10.Test(Solver)
    Case "BadName$":                    ApiTest = BadName11.Test(Solver)
    Case "%BadName":                    ApiTest = BadName12.Test(Solver)
    Case "Bad%Name":                    ApiTest = BadName13.Test(Solver)
    Case "BadName%":                    ApiTest = BadName14.Test(Solver)
    Case "^BadName":                    ApiTest = BadName15.Test(Solver)
    Case "Bad^Name":                    ApiTest = BadName16.Test(Solver)
    Case "BadName^":                    ApiTest = BadName17.Test(Solver)
    Case "&BadName":                    ApiTest = BadName18.Test(Solver)
    Case "Bad&Name":                    ApiTest = BadName19.Test(Solver)
    Case "BadName&":                    ApiTest = BadName20.Test(Solver)
    Case "|BadName":                    ApiTest = BadName21.Test(Solver)
    Case "Bad|Name":                    ApiTest = BadName22.Test(Solver)
    Case "BadName|":                    ApiTest = BadName23.Test(Solver)
    Case "-BadName":                    ApiTest = BadName24.Test(Solver)
    Case "Bad-Name":                    ApiTest = BadName25.Test(Solver)
    Case "BadName-":                    ApiTest = BadName26.Test(Solver)
    Case "=BadName":                    ApiTest = BadName27.Test(Solver)
    Case "Bad=Name":                    ApiTest = BadName28.Test(Solver)
    Case "BadName=":                    ApiTest = BadName29.Test(Solver)
    Case "EscapeSheetName(1)+2-1":      ApiTest = EscapeSheetName.Test(Solver)
    Case Else:                          ApiTest = NA
    End Select
End Function


