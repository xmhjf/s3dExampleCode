Attribute VB_Name = "ControlledInputs"
Option Explicit

Private Type InputInfo
    iIndex As Integer
    sName As String
    sPrompt As String
    sFilter As String
    lMinConnected As Long
    lMaxConnected As Long
    sComputeIIDs As String
End Type
Private Type ControlledInputInfo
    iIndex As Integer
    sName As String
    sComputeIIDs As String
End Type
Private Type ParameterInfo
    iIndex As Integer
    sName As String
    sUserName As String
    eParameterType As GCParameterType
    eUnitsOfPrimaryUOM As Units
    eUnitsOfSecondaryUOM As Units
    eUnitsOfTertiaryUOM As Units
    eUnitTypes As UnitTypes
    vDefaultValue As Variant
    bIsDisplayedOnRibbonBar As Boolean
    sCategoryName As String
 End Type
 Private Type ErrorInfo
    iIndex As Integer
    lErrorNumber As Long
    sShortDescription As String
    sLongDescription As String
    sCodeListTableName As String
    bIsForeign As Boolean
End Type
 
Private Type ParameterValueInfo
    iIndex As Integer
    sName As String
    sKey As String
    sCodeListValue As String
    lCodeList As Long
End Type
'
' management of controlled inputs
'
Public Sub GeometricConstructionDefinition_CopyInputs(ByVal pGCDefinitionOfTarget As GeometricConstructionDefinition, _
                                                      ByVal pGCDefinitionOfSource As GeometricConstructionDefinition, _
                                                      Optional ByVal sInputSubstitutionsList As String = "")
    ' split the substitution list
    Dim iCountOfSubstitutions As Integer: Let iCountOfSubstitutions = 0
    Dim sNamesOfTargetInputs() As String
    Dim sNamesOfSourceInputs() As String
    If sInputSubstitutionsList <> "" Then
        Dim sTokens() As String: Let sTokens = Split(sInputSubstitutionsList, ",")
        Let iCountOfSubstitutions = UBound(sTokens, 1) - LBound(sTokens, 1) + 1
        ReDim sNamesOfTargetInputs(1 To iCountOfSubstitutions)
        ReDim sNamesOfSourceInputs(1 To iCountOfSubstitutions)
        Dim i As Integer
        For i = 1 To iCountOfSubstitutions
            Dim sNames() As String: Let sNames = Split(sTokens(i - 1), "|")
            Dim iCountOfNames As Integer: Let iCountOfNames = UBound(sNames, 1) - LBound(sNames, 1) + 1
            If iCountOfNames <> 2 Then
                'MsgBox ("GeometricConstruction_CopyInputsAndParameters: each token should contain 2 na,es separated by a pipe")
                sNamesOfTargetInputs(i) = ""
                sNamesOfSourceInputs(i) = ""
            Else
                sNamesOfTargetInputs(i) = sNames(0)
                sNamesOfSourceInputs(i) = sNames(1)
            End If
        Next
    End If
    
    ' copy inputs
    For i = 1 To pGCDefinitionOfSource.InputCount
        Dim InputInfo As InputInfo: Let InputInfo = GeometricConstructionDefinition_GetInputInfoByIndex(pGCDefinitionOfSource, i)
        If iCountOfSubstitutions <> 0 Then
            Dim j As Integer
            For j = 1 To iCountOfSubstitutions
                If sNamesOfSourceInputs(j) = InputInfo.sName Then
                    InputInfo.sName = sNamesOfTargetInputs(j)
                    Exit For
                End If
            Next
        End If
        
        ' no target name indicates an input which nedds to be skipped
        If InputInfo.sName <> "" Then Call GeometricConstructionDefinition_AddInput(pGCDefinitionOfTarget, InputInfo)
    Next
End Sub
Public Sub GeometricConstructionDefinition_CopyControlledInputsEx(ByVal pGCDefinitionOfTarget As GeometricConstructionDefinition, _
                                                                   ByVal sProgid As String, _
                                                                   ByVal sSuffix As String)
    Dim pGCDefinitionService As IJGeometricConstructionDefinitionService
    Set pGCDefinitionService = CreateObject(sProgid)
    Dim pGCDefinition As IJGeometricConstructionDefinition
    Set pGCDefinition = New GeometricConstructionDefinition
    Call pGCDefinitionService.Initialize(pGCDefinition)
    
    Call GeometricConstructionDefinition_CopyControlledInputs(pGCDefinitionOfTarget, pGCDefinition, sSuffix)
End Sub
Public Sub GeometricConstructionDefinition_CopyControlledInputs(ByVal pGCDefinitionOfTarget As GeometricConstructionDefinition, _
                                                                ByVal pGCDefinitionOfSource As GeometricConstructionDefinition, _
                                                                ByVal sSuffix As String)
    ' copy controlled inputs
    Dim i As Integer
    For i = 1 To pGCDefinitionOfSource.ControlledInputCount
        Dim ControlledInputInfo As ControlledInputInfo
        ControlledInputInfo = GeometricConstructionDefinition_GetControlledInputInfoByIndex(pGCDefinitionOfSource, i)
        ControlledInputInfo.sName = ControlledInputInfo.sName + sSuffix
        Call GeometricConstructionDefinition_AddControlledInput(pGCDefinitionOfTarget, ControlledInputInfo)
    Next
End Sub
Public Sub GeometricConstruction_PropagateControlledInputs(ByVal pGCOfTarget As GeometricConstruction, _
                                                           ByVal pGCOfSource As GeometricConstruction, _
                                                           ByVal sSuffix As String)

    ' retrieve GCDefinition
    Dim pGCDefinitionOfSource As GeometricConstructionDefinition
    Set pGCDefinitionOfSource = pGCOfSource.Definition

    ' propagate controlled inputs
    Dim i As Integer
    For i = 1 To pGCDefinitionOfSource.ControlledInputCount
        Dim ControlledInputInfo As ControlledInputInfo
        ControlledInputInfo = GeometricConstructionDefinition_GetControlledInputInfoByIndex(pGCDefinitionOfSource, i)
        If ControlledInputInfo.sName <> "_Parameters" Then
            Call pGCOfTarget.ControlledInputs(ControlledInputInfo.sName + sSuffix).Clear
            If pGCOfSource.ControlledInputs(ControlledInputInfo.sName).Count > 0 Then
                Call pGCOfTarget.ControlledInputs(ControlledInputInfo.sName + sSuffix).AddElements(pGCOfSource.ControlledInputs(ControlledInputInfo.sName))
            End If
        End If
    Next
End Sub
Public Sub GeometricConstruction_ClearControlledInputs(ByVal pGCOfTarget As GeometricConstruction, _
                                                       ByVal pGCDefinitionOfSource As IJGeometricConstructionDefinition, _
                                                       ByVal sSuffix As String)
    ' clear controlled inputs
    Dim i As Integer
    For i = 1 To pGCDefinitionOfSource.ControlledInputCount
        Dim ControlledInputInfo As ControlledInputInfo
        ControlledInputInfo = GeometricConstructionDefinition_GetControlledInputInfoByIndex(pGCDefinitionOfSource, i)
        If ControlledInputInfo.sName <> "_Parameters" Then
            Call pGCOfTarget.ControlledInputs(ControlledInputInfo.sName + sSuffix).Clear
        End If
    Next
End Sub
Private Function GeometricConstructionDefinition_GetInputInfoByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                                     ByVal iIndex As Integer) As InputInfo
    ' prepare result
    Dim InputInfo As InputInfo

    InputInfo.iIndex = iIndex
    InputInfo.sName = ""
    InputInfo.sPrompt = ""
    InputInfo.sFilter = ""
    InputInfo.sComputeIIDs = ""
    Call pGCDefinition.GetInputInfoByIndex(iIndex, _
                                           InputInfo.sName, _
                                           InputInfo.sPrompt, _
                                           InputInfo.sFilter, _
                                           InputInfo.lMinConnected, _
                                           InputInfo.lMaxConnected, _
                                           InputInfo.sComputeIIDs)

    ' return result
    GeometricConstructionDefinition_GetInputInfoByIndex = InputInfo
End Function
Private Function GeometricConstructionDefinition_GetControlledInputInfoByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                                               ByVal iIndex As Integer) As ControlledInputInfo
    ' prepare result
    Dim ControlledInputInfo As ControlledInputInfo

    ControlledInputInfo.iIndex = iIndex
    ControlledInputInfo.sName = ""
    ControlledInputInfo.sComputeIIDs = ""
    Call pGCDefinition.GetControlledInputInfoByIndex(iIndex, _
                                                     ControlledInputInfo.sName, _
                                                     ControlledInputInfo.sComputeIIDs)

    ' return result
    GeometricConstructionDefinition_GetControlledInputInfoByIndex = ControlledInputInfo
End Function
Public Sub GeometricConstructionDefinition_AddInput(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                    ByRef InputInfo As InputInfo)
    If InputInfo.sComputeIIDs = "" Then
        Call pGCDefinition.AddInput(InputInfo.sName, _
                                    InputInfo.sPrompt, _
                                    InputInfo.sFilter, _
                                    InputInfo.lMinConnected, _
                                    InputInfo.lMaxConnected, _
                                    "IJGeometry")
    Else
        Call pGCDefinition.AddInput(InputInfo.sName, _
                                    InputInfo.sPrompt, _
                                    InputInfo.sFilter, _
                                    InputInfo.lMinConnected, _
                                    InputInfo.lMaxConnected, _
                                    InputInfo.sComputeIIDs)
    End If
End Sub
Private Sub GeometricConstructionDefinition_AddControlledInput(ByVal pGCDefinition As GeometricConstructionDefinition, _
                                                               ByRef ControlledInputInfo As ControlledInputInfo)
    If ControlledInputInfo.sComputeIIDs = "" Then
        Call pGCDefinition.AddControlledInput(ControlledInputInfo.sName, "IJGeometry")
    Else
        Call pGCDefinition.AddControlledInput(ControlledInputInfo.sName, _
                                              ControlledInputInfo.sComputeIIDs)
    End If
End Sub
Private Sub GeometricConstructionDefinition_AddErrorValue(ByVal pGCDefinition As GeometricConstructionDefinition, _
                                                          ByRef ErrorInfo As ErrorInfo)
    Call pGCDefinition.AddErrorValue(ErrorInfo.lErrorNumber, _
                                     ErrorInfo.sShortDescription, _
                                     ErrorInfo.sLongDescription)
End Sub
Public Sub GeometricConstructionDefinition_CopyParameters(ByVal pGCDefinitionOfTarget As GeometricConstructionDefinition, _
                                                          ByVal pGCDefinitionOfSource As GeometricConstructionDefinition, _
                                                          ByVal sSuffix As String)
    Dim i As Integer
    For i = 1 To pGCDefinitionOfSource.ParameterCount
        Dim ParameterInfo As ParameterInfo: ParameterInfo = GeometricConstructionDefinition_GetParameterInfoByIndex(pGCDefinitionOfSource, i)
        Call GeometricConstructionDefinition_AddParameter(pGCDefinitionOfTarget, ParameterInfo, pGCDefinitionOfSource, sSuffix)
    Next
End Sub
Public Sub GeometricConstructionDefinition_CopyErrors(ByVal pGCDefinitionOfTarget As GeometricConstructionDefinition, _
                                                      ByVal pGCDefinitionOfSource As GeometricConstructionDefinition, _
                                                      Optional ByVal iOffset As Integer = 0)
    Dim i As Integer
    For i = 1 To pGCDefinitionOfSource.ErrorsCount
        Dim ErrorInfo As ErrorInfo: ErrorInfo = GeometricConstructionDefinition_GetErrorInfoByIndex(pGCDefinitionOfSource, i)
        ErrorInfo.lErrorNumber = ErrorInfo.lErrorNumber + iOffset
        Call GeometricConstructionDefinition_AddErrorValue(pGCDefinitionOfTarget, ErrorInfo)
    Next
End Sub
 Public Function GeometricConstructionDefinition_GetParameterInfoByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                                         ByVal iIndex As Integer, _
                                                                         Optional ByVal iOffset As Integer = 0) As ParameterInfo
    ' prepare result
    Dim ParameterInfo As ParameterInfo

    ParameterInfo.iIndex = iIndex + iOffset
    ParameterInfo.sName = ""
    ParameterInfo.sUserName = ""
    ParameterInfo.vDefaultValue = ""
    Call pGCDefinition.GetParameterInfoByIndex(iIndex, _
                                               ParameterInfo.sName, _
                                               ParameterInfo.sUserName, _
                                               ParameterInfo.eParameterType, _
                                               ParameterInfo.eUnitsOfPrimaryUOM, _
                                               ParameterInfo.eUnitsOfSecondaryUOM, _
                                               ParameterInfo.eUnitsOfTertiaryUOM, _
                                               ParameterInfo.eUnitTypes, _
                                               ParameterInfo.vDefaultValue, _
                                               ParameterInfo.bIsDisplayedOnRibbonBar, _
                                               ParameterInfo.sCategoryName)

    ' return result
    GeometricConstructionDefinition_GetParameterInfoByIndex = ParameterInfo
End Function
 Public Function GeometricConstructionDefinition_GetErrorInfoByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                                     ByVal iIndex As Integer) As ErrorInfo
    ' prepare result
    Dim ErrorInfo As ErrorInfo

    ErrorInfo.iIndex = iIndex
    ErrorInfo.lErrorNumber = -1
    ErrorInfo.sShortDescription = ""
    ErrorInfo.sLongDescription = ""
    ErrorInfo.sCodeListTableName = ""
    ErrorInfo.bIsForeign = False
    
    Call pGCDefinition.GetErrorInfoByIndex(iIndex, _
                                           ErrorInfo.lErrorNumber, _
                                           ErrorInfo.sShortDescription, _
                                           ErrorInfo.sLongDescription, _
                                           ErrorInfo.sCodeListTableName, _
                                           ErrorInfo.bIsForeign)
    ' return result
    GeometricConstructionDefinition_GetErrorInfoByIndex = ErrorInfo
End Function
Public Sub GeometricConstructionDefinition_AddParameter(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                        ByRef ParameterInfo As ParameterInfo, _
                                                        ByVal pGCDefinitionOfWrapper As IJGeometricConstructionDefinition, _
                                                        Optional ByVal sSuffix As String = "")
    Call pGCDefinition.AddParameter(ParameterInfo.sName + sSuffix, _
                                    ParameterInfo.sUserName, _
                                    ParameterInfo.eParameterType, _
                                    ParameterInfo.eUnitTypes, _
                                    ParameterInfo.eUnitsOfPrimaryUOM, _
                                    ParameterInfo.eUnitsOfSecondaryUOM, _
                                    ParameterInfo.eUnitsOfTertiaryUOM, _
                                    ParameterInfo.vDefaultValue, _
                                    ParameterInfo.bIsDisplayedOnRibbonBar)

    If ParameterInfo.eParameterType = GCParameterType.GCCodeList Then
        Dim i As Integer
        For i = 1 To pGCDefinitionOfWrapper.ParameterValueCount
            Dim ParameterValueInfo As ParameterValueInfo: ParameterValueInfo = GeometricConstructionDefinition_GetParameterValueInfoByIndex(pGCDefinitionOfWrapper, i)
            If ParameterValueInfo.sName = ParameterInfo.sName Then
                Call pGCDefinition.AddParameterValue(ParameterValueInfo.sName + sSuffix, _
                                                     ParameterValueInfo.sCodeListValue, _
                                                     ParameterValueInfo.lCodeList)
            End If
        Next
    End If
End Sub
    Public Function GeometricConstructionDefinition_GetParameterValueInfoByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, _
                                                                                 ByVal iIndex As Integer) As ParameterValueInfo
        ' prepare result
        Dim ParameterValueInfo As ParameterValueInfo

        ParameterValueInfo.iIndex = iIndex
        ParameterValueInfo.sKey = ""
        ParameterValueInfo.sCodeListValue = ""
        Call pGCDefinition.GetParameterValueInfoByIndex(iIndex, _
                                                        ParameterValueInfo.sName, _
                                                        ParameterValueInfo.sKey, _
                                                        ParameterValueInfo.sCodeListValue, _
                                                        ParameterValueInfo.lCodeList)
        ' return result
        GeometricConstructionDefinition_GetParameterValueInfoByIndex = ParameterValueInfo
    End Function
Public Sub GeometricConstruction_CopyParameters(ByVal pGCOfTarget As GeometricConstruction, _
                                                ByVal pGCOfSource As GeometricConstruction, _
                                                ByVal sSuffix As String)
    Dim pGCDefinitionOfTarget As GeometricConstructionDefinition: Set pGCDefinitionOfTarget = pGCOfTarget.Definition
    Dim pGCDefinitionOfSource As GeometricConstructionDefinition: Set pGCDefinitionOfSource = pGCOfSource.Definition
    Dim i As Integer
    For i = 1 To pGCDefinitionOfTarget.ParameterCount
        Dim sNameOfTargetParameter As String: Let sNameOfTargetParameter = GeometricConstructionDefinition_GetParameterNameByIndex(pGCDefinitionOfTarget, i)
        Dim sNameOfSourceParameter As String: Let sNameOfSourceParameter = sNameOfTargetParameter + sSuffix
        pGCOfTarget.Parameter(sNameOfTargetParameter) = pGCOfSource.Parameter(sNameOfSourceParameter)
    Next
End Sub
Private Function GeometricConstructionDefinition_GetParameterNameByIndex(ByVal pGCDefinition0 As IJGeometricConstructionDefinition, ByVal iIndex As Integer) As String
    Dim sName As String: Let sName = ""
    Dim sUserName As String
    Dim eParameterType As GCParameterType
    Dim eUnitsOfPrimaryUOM As Units
    Dim eUnitsOfSecondaryUOM As Units
    Dim eUnitsOfTertiaryUOM As Units
    Dim eUnitTypes As UnitTypes
    Dim vValue As Variant
    Dim bIsDisplayedOnRBB As Boolean
    Dim sCategoryName As String
    Call pGCDefinition0.GetParameterInfoByIndex(iIndex, sName, sUserName, eParameterType, eUnitsOfPrimaryUOM, eUnitsOfSecondaryUOM, eUnitsOfTertiaryUOM, eUnitTypes, vValue, bIsDisplayedOnRBB, sCategoryName)
    GeometricConstructionDefinition_GetParameterNameByIndex = sName
End Function
Public Sub GeometricConstruction_CopyInputs(ByVal pGCOfTarget As GeometricConstruction, _
                                            ByVal pGCOfSource As GeometricConstruction, _
                                            ByRef pGCs() As GeometricConstruction, _
                                            Optional ByVal sSubstitutionsList As String = "")
    Dim iCountOfSubstitutions As Integer: Let iCountOfSubstitutions = 0
    Dim sNamesOfTargetInputs() As String
    Dim sNamesOfSourceInputs() As String
    If sSubstitutionsList <> "" Then
        ' split the substitution list
        Dim sTokens() As String: Let sTokens = Split(sSubstitutionsList, ",")
        Let iCountOfSubstitutions = UBound(sTokens, 1) - LBound(sTokens, 1) + 1
        ReDim sNamesOfTargetInputs(1 To iCountOfSubstitutions)
        ReDim sNamesOfSourceInputs(1 To iCountOfSubstitutions)
        Dim i As Integer
        For i = 0 To iCountOfSubstitutions - 1
            Dim sNames() As String: Let sNames = Split(sTokens(i), "|")
            Dim iCountOfNames As Integer: Let iCountOfNames = UBound(sNames, 1) - LBound(sNames, 1) + 1
            If iCountOfNames <> 2 Then
                MsgBox ("GeometricConstruction_CopyInputsAndParameters: each token should contain 2 names separated by a pipe")
                sNamesOfTargetInputs(i + 1) = ""
                sNamesOfSourceInputs(i + 1) = ""
            Else
                sNamesOfTargetInputs(i + 1) = sNames(0)
                sNamesOfSourceInputs(i + 1) = sNames(1)
            End If
        Next i
    End If
    
    Dim pGCDefinitionOfTarget As GeometricConstructionDefinition: Set pGCDefinitionOfTarget = pGCOfTarget.Definition
    Dim pGCDefinitionOfSource As GeometricConstructionDefinition: Set pGCDefinitionOfSource = pGCOfSource.Definition
    For i = 1 To pGCDefinitionOfTarget.InputCount
        Dim sNameOfTargetInput As String: Let sNameOfTargetInput = GeometricConstructionDefinition_GetInputNameByIndex(pGCDefinitionOfTarget, i)
        Dim sNameOfSourceInput As String: Let sNameOfSourceInput = sNameOfTargetInput
        
        If iCountOfSubstitutions <> 0 Then
            Dim j As Integer
            For j = 1 To iCountOfSubstitutions
                If sNamesOfTargetInputs(j) = sNameOfTargetInput Then
                    sNameOfSourceInput = sNamesOfSourceInputs(j)
                    Exit For
                End If
            Next j
        End If

        ' decode name of source input
        Let sTokens = Split(sNameOfSourceInput, ".")
        If UBound(sTokens, 1) = 0 Then
            ' copy from input
            If GeometricConstructionDefinition_GetInputIndexByName(pGCDefinitionOfSource, sNameOfSourceInput) <> 0 Then
                If pGCOfSource.Inputs(sNameOfSourceInput).Count > 0 Then pGCOfTarget.Input(sNameOfTargetInput) = pGCOfSource.Input(sNameOfSourceInput)
            End If
        Else
            ' copy from output
            Dim pGCOfSourceTemp As GeometricConstruction: Set pGCOfSourceTemp = pGCs(Val(sTokens(0)))
            Dim pGCDefinitionOfSourceTemp As GeometricConstructionDefinition: Set pGCDefinitionOfSourceTemp = pGCOfSourceTemp.Definition
            Dim sNameOfSourceOutputTemp As String: Let sNameOfSourceOutputTemp = sTokens(1)
            If sTokens(1) = "" Then
                pGCOfTarget.Input(sNameOfTargetInput) = pGCOfSourceTemp
            Else
                Dim sKeyOfSourceOutputTemp As String: Let sKeyOfSourceOutputTemp = ""
                If UBound(sTokens, 1) = 2 Then Let sKeyOfSourceOutputTemp = sTokens(2)
    
                Dim pGCMacroOfSourceTemp As IJGeometricConstructionMacro: Set pGCMacroOfSourceTemp = pGCOfSourceTemp
                If pGCMacroOfSourceTemp.Outputs(sNameOfSourceOutputTemp).Count > 0 Then
                    If sKeyOfSourceOutputTemp = "" Then
                        pGCOfTarget.Input(sNameOfTargetInput) = pGCMacroOfSourceTemp.Outputs(sNameOfSourceOutputTemp)(1)
                    Else
                        pGCOfTarget.Input(sNameOfTargetInput) = pGCMacroOfSourceTemp.Outputs(sNameOfSourceOutputTemp)(sKeyOfSourceOutputTemp)
                    End If
                End If
            End If
        End If
    Next i
End Sub
Public Sub GeometricConstruction_CopyOutputs(ByVal pGCOfTarget As GeometricConstruction, _
                                             ByVal pGCOfSource As GeometricConstruction, _
                                             Optional ByVal sSubstitutionsList As String = "")
    Dim iCountOfSubstitutions As Integer: Let iCountOfSubstitutions = 0
    Dim sNamesOfTargetOutputs() As String
    Dim sNamesOfSourceOutputs() As String
    If sSubstitutionsList <> "" Then
        ' split the substitution list
        Dim sTokens() As String: Let sTokens = Split(sSubstitutionsList, ",")
        Let iCountOfSubstitutions = UBound(sTokens, 1) - LBound(sTokens, 1) + 1
        ReDim sNamesOfTargetOutputs(1 To iCountOfSubstitutions)
        ReDim sNamesOfSourceOutputs(1 To iCountOfSubstitutions)
        Dim i As Integer
        For i = 0 To iCountOfSubstitutions - 1
            Dim sNames() As String: Let sNames = Split(sTokens(i), "|")
            Dim iCountOfNames As Integer: Let iCountOfNames = UBound(sNames, 1) - LBound(sNames, 1) + 1
            If iCountOfNames <> 2 Then
                MsgBox ("GeometricConstruction_CopyOutputs: each token should contain 2 names separated by a pipe")
                sNamesOfTargetOutputs(i + 1) = ""
                sNamesOfSourceOutputs(i + 1) = ""
            Else
                sNamesOfTargetOutputs(i + 1) = sNames(0)
                sNamesOfSourceOutputs(i + 1) = sNames(1)
            End If
        Next
    End If
    
    Dim pGCDefinitionOfTarget As GeometricConstructionDefinition: Set pGCDefinitionOfTarget = pGCOfTarget.Definition
    Dim pGCDefinitionOfSource As GeometricConstructionDefinition: Set pGCDefinitionOfSource = pGCOfSource.Definition
    For i = 1 To pGCDefinitionOfTarget.OutputCount
        Dim sNameOfTargetOutput As String: Let sNameOfTargetOutput = GeometricConstructionDefinition_GetOutputNameByIndex(pGCDefinitionOfTarget, i)
        Dim sNameOfSourceOutput As String: Let sNameOfSourceOutput = sNameOfTargetOutput
        If iCountOfSubstitutions <> 0 Then
            Dim j As Integer
            For j = 1 To iCountOfSubstitutions
                If sNamesOfTargetOutputs(j) = sNameOfTargetOutput Then
                    sNameOfSourceOutput = sNamesOfSourceOutputs(j)
                    Exit For
                End If
            Next
        End If
        
        If GeometricConstructionDefinition_GetOutputIndexByName(pGCDefinitionOfSource, sNameOfSourceOutput) <> 0 Then
            Dim pGCMacroOfSource As IJGeometricConstructionMacro: Set pGCMacroOfSource = pGCOfSource
            Dim pGCMacroOfTarget As IJGeometricConstructionMacro: Set pGCMacroOfTarget = pGCOfTarget
            'If pGCMacroOfSource.Outputs(sNameOfSourceOutput).Count > 0 Then Call pGCMacroOfTarget.Outputs(sNameOfTargetOutput).AddElements(pGCMacroOfSource.Outputs(sNameOfSourceOutput))
            For j = 1 To pGCMacroOfSource.Outputs(sNameOfSourceOutput).Count
                Dim oSourceOutput As Object: Set oSourceOutput = pGCMacroOfSource.Outputs(sNameOfSourceOutput)(j)
                Dim sKey As String: sKey = pGCMacroOfSource.Outputs(sNameOfSourceOutput).GetKey(oSourceOutput)
                pGCOfTarget.Output(sNameOfTargetOutput, sKey) = oSourceOutput
            Next
        End If
    Next
End Sub
Private Function GeometricConstructionDefinition_GetInputNameByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, ByVal iIndex As Integer) As String
    Dim sName As String: Let sName = ""
    
    Dim sPrompt As String
    Dim sFilter As String
    Dim lMinConnected As Long, lMaxConnected As Long
    Dim sComputeIIDs As String
    Call pGCDefinition.GetInputInfoByIndex(iIndex, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
    
    GeometricConstructionDefinition_GetInputNameByIndex = sName
End Function
Private Function GeometricConstructionDefinition_GetOutputNameByIndex(ByVal pGCDefinition As IJGeometricConstructionDefinition, ByVal iIndex As Integer) As String
    Dim sName As String: Let sName = ""
    
    Dim eOutputType As GCOutputType
    Call pGCDefinition.GetOutputInfoByIndex(iIndex, sName, eOutputType)
    
    GeometricConstructionDefinition_GetOutputNameByIndex = sName
End Function

Private Function GeometricConstructionDefinition_GetInputIndexByName(ByVal pGCDefinition As IJGeometricConstructionDefinition, ByVal sName0 As String) As Integer
     Dim iIndex As Integer: Let iIndex = 0

     Dim i As Integer
     For i = 1 To pGCDefinition.InputCount
         Dim sName As String: Let sName = ""
         Dim sPrompt As String
         Dim sFilter As String
         Dim lMinConnected As Long, lMaxConnected As Long
         Dim sComputeIIDs As String
         Call pGCDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
         If sName = sName0 Then
             iIndex = i
             Exit For
         End If
     Next

     GeometricConstructionDefinition_GetInputIndexByName = iIndex
 End Function
 Private Function GeometricConstructionDefinition_GetOutputIndexByName(ByVal pGCDefinition As IJGeometricConstructionDefinition, ByVal sName0 As String) As Integer
     Dim iIndex As Integer: Let iIndex = 0

     Dim i As Integer
     For i = 1 To pGCDefinition.OutputCount
         Dim sName As String: Let sName = ""
         Dim eOutputType As GCOutputType
         Call pGCDefinition.GetOutputInfoByIndex(i, sName, eOutputType)
         If sName = sName0 Then
             iIndex = i
             Exit For
         End If
     Next

     GeometricConstructionDefinition_GetOutputIndexByName = iIndex
 End Function
 Public Function GeometricConstruction_GetInput(ByVal pGC As GeometricConstruction, _
                                                ByVal sName As String) As Object
    Dim oInput As Object: Set oInput = Nothing
    
    If pGC.Inputs(sName).Count = 1 Then
        Set oInput = pGC.Input(sName)
    End If
    
    Set GeometricConstruction_GetInput = oInput
End Function
Public Function GetSuffixOfComponent(iIndexOfComponent As Integer) As String
    GetSuffixOfComponent = "_" + CStr(iIndexOfComponent)
End Function
Public Sub GeometricConstructionCMacro_RemoveBoundary(pGCMacro As IJGeometricConstructionMacro, sBoundaryName As String)
    Dim i As Integer
    For i = 1 To pGCMacro.Outputs("Boundary").Count
        Dim oBoundary As Object: Set oBoundary = pGCMacro.Outputs("Boundary")(i)
        Dim sKey As String: Let sKey = pGCMacro.Outputs("Boundary").GetKey(oBoundary)
        If sKey = sBoundaryName Then
            Call pGCMacro.Outputs("Boundary").Remove(sKey)
            Dim pObject As IJDObject: Set pObject = oBoundary
            pObject.Remove
            Exit For
        End If
    Next
End Sub
Function GeometricConstructionDefinition_GetFromTypeName(sTypeName As String) As IJGeometricConstructionDefinition
    Dim pGCEntitiesFactory As IJGeometricConstructionEntitiesFactory: Set pGCEntitiesFactory = New GeometricConstructionEntitiesFactory
    Dim pGC As IJGeometricConstruction: Set pGC = pGCEntitiesFactory.CreateEntity(sTypeName, Nothing)
    Dim pGCType As IJGeometricConstructionType: Set pGCType = pGC.Type
    Set GeometricConstructionDefinition_GetFromTypeName = pGCType.GetInitializedDefinition
End Function
Public Function GetDefinitionFromTypeName(sTypeName As String) As IJGeometricConstructionDefinition
      ' prerare result
      Dim pGCDefinition As IJGeometricConstructionDefinition: Set pGCDefinition = New GeometricConstructionDefinition
      'MsgBox "sTypeName= " + sTypeName
      ' retrieve the GCType
      Dim pGCEntitiesFactory As IJGeometricConstructionEntitiesFactory: Set pGCEntitiesFactory = New GeometricConstructionEntitiesFactory
      Dim pGC As IJGeometricConstruction: Set pGC = pGCEntitiesFactory.CreateEntity(sTypeName, Nothing)
      Dim pGCType As IJGeometricConstructionType: Set pGCType = pGC.Type

      ' retrieve the definitionService
      Dim pGCDefinitionService As IJGeometricConstructionDefinitionService: Set pGCDefinitionService = pGCType.GetDefinitionService()
      
      ' initialize the definition
      Call pGCDefinitionService.Initialize(pGCDefinition)

      ' return result
      'MsgBox "Definition retrieved"
      Set GetDefinitionFromTypeName = pGCDefinition
  End Function

