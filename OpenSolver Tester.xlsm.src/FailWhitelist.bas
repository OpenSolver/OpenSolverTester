Attribute VB_Name = "FailWhitelist"
' A whitelist for tests that are expected to fail
Function TestShouldFail(SheetName As String, Solver As String) As Boolean
    Select Case SheetName & "_" & Solver
    
    ' ==================================
    ' COUENNE
    ' ==================================
    
    Case "Test3_Couenne"        ' Reports infeasible, bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test4_Couenne"        ' Reports infeasible, bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test9A_Couenne"       ' Reports optimal when problem is infeasible, bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test15c_Couenne"      ' Incorrect optimal solution, bound on binary var is removed. Bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test15d_Couenne"      ' Sometimes gets incorrect optimal solution, bound on binary var is removed. Bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test22_Couenne"       ' Ignores lower bound, bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test23_Couenne"       ' Ignores lower bound, bug with AMPL Couenne 0.4.7, older/NEOS versions work fine
        TestShouldFail = True
        
    Case "Test28_CBCOptions_Couenne"    ' Doesn't report unbounded solution
        TestShouldFail = True
        
    Case "Test36_Couenne"       ' Couenne doesn't yet support MAX (0.4.7)
        TestShouldFail = True
        
    ' ==================================
    ' BONMIN
    ' ==================================
        
    Case "Test28_CBCOptions_Bonmin"     ' Reports infeasible rather than unbounded
        TestShouldFail = True
        
    Case "Test41_Bonmin"        ' Bonmin can't solve this problem, reports unbounded. Same with NEOS Bonmin
        TestShouldFail = True
        
    ' ==================================
    ' NOMAD
    ' ==================================
        
    Case "Test28_CBCOptions_NOMAD"      ' Reports large optimal solution rather than unbounded
        TestShouldFail = True
        
    ' ==================================
    ' NEOS COUENNE
    ' ==================================
        
    Case "Test28_CBCOptions_NeosCou"     ' Reports large optimal solution rather than unbounded
        TestShouldFail = True
        
    Case "Test36_NeosCou"       ' Couenne doesn't yet support MAX (0.4.7)
        TestShouldFail = True
        
    ' ==================================
    ' NEOS BONMIN
    ' ==================================
        
    Case "Test40_NeosBon"       ' Crashes on "invalid number". Seems to be an AMPL > .nl conversion error, as our .nl version works fine
        TestShouldFail = True
        
    Case "Test41_NeosBon"       ' Bonmin can't solve this problem, reports unbounded. Same with local Bonmin
        TestShouldFail = True
        
        
    Case Else
        TestShouldFail = False

    End Select
End Function
