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
    Case "IP_SUMIF":                    ApiTest = IP_SUMIF.Test(Solver)
    Case "OverlappingVars":             ApiTest = OverlappingVars.Test(Solver)
    Case "BlankRHS":                    ApiTest = BlankRHS.Test(Solver)
    Case "BlankLHS":                    ApiTest = BlankLHS.Test(Solver)
    Case "TextRHS":                     ApiTest = TextRHS.Test(Solver)
    Case "ErrorRHS":                    ApiTest = ErrorRHS.Test(Solver)
    Case "ErrorLHS":                    ApiTest = ErrorLHS.Test(Solver)
    Case "ErrorObj":                    ApiTest = ErrorObj.Test(Solver)
    Case "TextObj":                     ApiTest = TextObj.Test(Solver)
    Case "MergedVarsSubset":            ApiTest = MergedVarsSubset.Test(Solver)
    Case "MergedVarsOK":                ApiTest = MergedVarsOK.Test(Solver)
    Case "MergedVarsBad":               ApiTest = MergedVarsBad.Test(Solver)
    Case "MergedRHS_OK":                ApiTest = MergedRHS_OK.Test(Solver)
    Case "MergedLHS_OK":                ApiTest = MergedLHS_OK.Test(Solver)
    Case "MergedLHSMultiple_OK":        ApiTest = MergedLHSMultiple_OK.Test(Solver)
    Case "BinIntOverlap":               ApiTest = BinIntOverlap.Test(Solver)
    Case "ComplexLayout":               ApiTest = ComplexLayout.Test(Solver)
    Case "ConFormula":                  ApiTest = ConFormula.Test(Solver)
    Case "ConFormulaInf":               ApiTest = ConFormulaInf.Test(Solver)
    Case "BadName!":                    ApiTest = BadName.Test(Solver)
    Case "InfModel":                    ApiTest = InfModel.Test(Solver)
    Case "DeletedConRef":               ApiTest = DeletedConRef.Test(Solver)
    Case "DeletedObjRef":               ApiTest = DeletedObjRef.Test(Solver)
    Case "InfConstConstraint":          ApiTest = InfConstConstraint.Test(Solver)
    Case "FeasConstConstraint":         ApiTest = FeasConstConstraint.Test(Solver)
    Case "Unbounded":                   ApiTest = Unbounded.Test(Solver)
    Case "ProtectedSheet":              ApiTest = ProtectedSheet.Test(Solver)
    Case "NonLin":                      ApiTest = NonLin.Test(Solver)
    Case "NonLinObj":                   ApiTest = NonLinObj.Test(Solver)
    Case "NonLinSimple":                ApiTest = NonLinSimple.Test(Solver)
    Case "NonLinLarge":                 ApiTest = NonLinLarge.Test(Solver)
    Case "IndirectLBs":                 ApiTest = IndirectLBs.Test(Solver)
    Case "DirectLBs":                   ApiTest = DirectLBs.Test(Solver)
    Case "SingleRangeLB":               ApiTest = SingleRangeLB.Test(Solver)
    Case "BinLB":                       ApiTest = BinLB.Test(Solver)
    Case "BinIntLB":                    ApiTest = BinIntLB.Test(Solver)
    Case "Relaxation":                  ApiTest = Relaxation.Test(Solver)
    Case "LargeRangeLB":                ApiTest = LargeRangeLB.Test(Solver)
    Case "OverlapRangeLB":              ApiTest = OverlapRangeLB.Test(Solver)
    Case "OverlapRangeLBwithFormula":   ApiTest = OverlapRangeLBwithFormula.Test(Solver)
    Case "ConstFormulaRangeLB":         ApiTest = ConstFormulaRangeLB.Test(Solver)
    Case "VariableFormulaRangeLB":      ApiTest = VariableFormulaRangeLB.Test(Solver)
    Case "ConstraintLB":                ApiTest = ConstraintLB.Test(Solver)
    Case "ConstantLB":                  ApiTest = ConstantLB.Test(Solver)
    Case "FormulaLB":                   ApiTest = FormulaLB.Test(Solver)
    Case "VarConstraintLB":             ApiTest = VarConstraintLB.Test(Solver)
    Case "IterativeCalc":               ApiTest = IterativeCalc.Test(Solver)
    Case "NoObj":                       ApiTest = NoObj.Test(Solver)
    Case "QuickSolve":                  ApiTest = QuickSolve.Test(Solver)
    Case "SolverParameters":            ApiTest = SolverParameters.Test(Solver)
    Case "Sensitivity":                 ApiTest = Sensitivity.Test(Solver)
    Case "NamedRanges":                 ApiTest = NamedRanges.Test(Solver)
    Case "NonLin2":                     ApiTest = NonLin2.Test(Solver)
    Case "NonLinPruning":               ApiTest = NonLinPruning.Test(Solver)
    Case "NonLinMinMax":                ApiTest = NonLinMinMax.Test(Solver)
    Case "NonLinConstraints":           ApiTest = NonLinConstraints.Test(Solver)
    Case "NonLin3":                     ApiTest = NonLin3.Test(Solver)
    Case "SensitivityNames":            ApiTest = SensitivityNames.Test(Solver)
    Case "NonLin4":                     ApiTest = NonLin4.Test(Solver)
    Case "NonLin5":                     ApiTest = NonLin5.Test(Solver)
    Case "NonLin6":                     ApiTest = NonLin6.Test(Solver)
    Case "FractionalCoeffs":            ApiTest = FractionalCoeffs.Test(Solver)
    Case "SeekObj":                     ApiTest = SeekObj.Test(Solver)
    Case "SeekObjInf":                  ApiTest = SeekObjInf.Test(Solver)
    Case "DiffSheetObj":                ApiTest = DiffSheetObj.Test(Solver)
    Case "EscapeSheetName(1)+2-1":      ApiTest = EscapeSheetName.Test(Solver)
    Case "NonLinBinary":                ApiTest = NonLinBinary.Test(Solver)
    Case "Warmstart":                   ApiTest = Warmstart.Test(Solver)
'    Case "Highlighting":                ApiTest = Highlighting.Test(Solver)
    Case "AutoModel":                   ApiTest = AutoModel.Test(Solver)
    Case Else:                          ApiTest = NA
    End Select
End Function


