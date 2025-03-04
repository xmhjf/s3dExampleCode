Attribute VB_Name = "GCTools"
'
' 21-mar-2008 : Patrice Blanchard : Initial implementation
'
' depends on Revision (IJRevision)
' depends on GeometricConstruction (IJGeometricConstruction)

Option Explicit

'''Public Function GeometricConstruction_SharedParameter(pGeometricConstruction As IJGeometricConstruction, sName As String) As Double
'''    Dim dValue As Double: Let dValue = CDbl(pGeometricConstruction.Parameter(sName))
''''    If pGeometricConstruction.Inputs("Parameter").Count = 1 Then
''''        Dim pGeometricConstructionOfParameter As IJGeometricConstruction
''''        Set pGeometricConstructionOfParameter = pGeometricConstruction.Input("Parameter")
''''        Let dValue = pGeometricConstructionOfParameter.Parameter("Value")
''''    End If
'''    Let GeometricConstruction_SharedParameter = dValue
'''End Function
Private m_iIndexOfGCs As Integer

Public Sub ResetEntityCount()
    m_iIndexOfGCs = 0
End Sub
Public Function CreateEntity(pGeometricConstructionEntitiesFactory As IJGeometricConstructionEntitiesFactory, sType As String, pPOM As IJDPOM) As Object
    Let m_iIndexOfGCs = m_iIndexOfGCs + 1
    Dim sIndex As String: Let sIndex = CStr(m_iIndexOfGCs)
    Dim sPrefix As String
    If Len(sIndex) = 1 Then Let sPrefix = "00" + sIndex
    If Len(sIndex) = 2 Then Let sPrefix = "0" + sIndex
    If Len(sIndex) = 3 Then Let sPrefix = sIndex
    
    Set CreateEntity = pGeometricConstructionEntitiesFactory.CreateEntity(sType, pPOM, sPrefix + "-" + sType)
End Function
Public Function GeometricConstruction_Input(pGeometricConstruction As IJGeometricConstruction, sName As String) As Object
    Set GeometricConstruction_Input = Nothing
    If pGeometricConstruction.Inputs(sName).Count = 1 Then
        Set GeometricConstruction_Input = pGeometricConstruction.Input(sName)
    End If
End Function
Public Sub GeometricConstructionDefinition_AddParameterValues(pGeometricConstructionDefinition As IJGeometricConstructionDefinition, sCodeListName As String)
    Select Case sCodeListName
        Case sTRACK_FLAG:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Near", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Far", 2)
        Case sLOOKING_AXIS:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "X", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Y", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Z", 3)
        Case sINTERSECTING_PLANE:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Not defined", 0)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(x,y)", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(y,z)", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(z,x)", 3)
        Case sSWEEP_ANGLE:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Small", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Large", 2)
        Case sAXES_ROLES:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(x,y)", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(y,z)", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(z,x)", 3)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(y,x)", 4)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(z,y)", 5)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "(x,z)", 6)
        Case sCS_ORIENTATION:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Direct", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Indirect", 2)
        Case sLINE_JUSTIFICATION:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Not-Centered", 0)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Centered", 1)
        Case sPOINT_LOCATION:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Inside", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Outside", 2)
        Case sCONTEXT1:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Start-Start", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Start-End", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "End-Start", 3)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "End-End", 4)
        Case sCONTEXT2:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Start-Start", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Start-End", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "End-Start", 3)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "End-End", 4)
        Case sGEOMETRY_SELECTOR:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Initial", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Current", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Stable", 4)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Bounded", 5)
        Case sFACES_CONTEXT:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Capping", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Lateral", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Side", 3)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Internal", 4)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "External", 5)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Fillet", 6)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Any face", 7)
        Case sSURFACE_TYPE:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "General", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "Planar", 2)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "NonPlanar", 3)
        Case sAXIS_DIRECTION, sAXIS1_DIRECTION, sAXIS2_DIRECTION:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "GCAsDefined", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "GCOpposite", 2)
         Case sDIRECTION:
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "GCPositive", 1)
            Call pGeometricConstructionDefinition.AddParameterValue(sCodeListName, "GCNegative", 2)
          
        Case Else
            MsgBox "Invalid CodeListName= " + sCodeListName
    End Select
End Sub
Function GeometricConstruction_Create(pUnknownOfResourceManager As IUnknown, sType As String) As IJGeometricConstruction
    ' create factory
    Dim pGeometricConstructionFactory As IJGeometricConstructionEntitiesFactory
    Set pGeometricConstructionFactory = New GeometricConstructionEntitiesFactory

    ' create custom construction
    Dim pGeometricConstruction As IJGeometricConstruction
    Set pGeometricConstruction = pGeometricConstructionFactory.CreateEntity(sType, pUnknownOfResourceManager)

    ' return result
    Set GeometricConstruction_Create = pGeometricConstruction
End Function
Sub GeometricConstruction_ConsumeObject(pGeometricConstruction As IJGeometricConstruction, oObject As Object)
    If Not TypeOf oObject Is IJGeometricConstruction Then Exit Sub

    Dim pElementsConsumed As IJElements: Set pElementsConsumed = pGeometricConstruction.ConsumedList
    
    If pElementsConsumed.Count = 1 Then
        If Not pElementsConsumed.Contains(oObject) Then
            Call pElementsConsumed.Clear
        End If
    End If
    
    If pElementsConsumed.Count = 0 Then
        Call pElementsConsumed.Add(oObject)
    End If
End Sub
Sub GeometricConstruction_UnConsumeObject(pGeometricConstruction As IJGeometricConstruction, oObject As Object)
    If Not TypeOf oObject Is IJGeometricConstruction Then Exit Sub

    Dim pElementsConsumed As IJElements: Set pElementsConsumed = pGeometricConstruction.ConsumedList
    If pElementsConsumed.Contains(oObject) Then Call pElementsConsumed.Remove(oObject)
End Sub
Function GeometricConstruction_IsObjectConsumed(pGeometricConstruction As IJGeometricConstruction, oObject As Object)
    Dim pElementsConsumed As IJElements: Set pElementsConsumed = pGeometricConstruction.ConsumedList
    Let GeometricConstruction_IsObjectConsumed = pElementsConsumed.Contains(oObject)
End Function
Sub GeometricConstruction_ConsumeElements(pGeometricConstruction As IJGeometricConstruction, pElements As IJElements)
    Dim pElementsConsumed As IJElements: Set pElementsConsumed = pGeometricConstruction.ConsumedList
    
    If pElementsConsumed.Count <> pElements.Count Then
        Call pElementsConsumed.Clear
    Else
        Dim i As Integer
        For i = 1 To pElements.Count
            If pElementsConsumed.Contains(pElements(i)) Then
                Call pElementsConsumed.Clear
                Exit For
            End If
        Next
    End If
    
    If pElementsConsumed.Count = 0 Then
        For i = 1 To pElements.Count
            Dim oElement As Object: Set oElement = pElements(i)
            If TypeOf oElement Is IJGeometricConstruction Then
                Call pElementsConsumed.Add(oElement)
            End If
        Next
    End If
End Sub
Public Function GeometricConstruction_GetNamesOfControlledInputs(pGeometricConstruction As IJGeometricConstruction) As Collection
    ' prepare a collection to be filled
    Dim pCollection As New Collection
    
    ' fill the collection with the names of the controlled inputs
    If True Then
        ' retrieve the definition
        Dim pGeometricConstructionDefinition As IJGeometricConstructionDefinition: Set pGeometricConstructionDefinition = pGeometricConstruction.definition
        
        ' loop on the controlled inputs
        Dim i As Integer
        For i = 1 To pGeometricConstructionDefinition.ControlledInputCount
            ' retrieve the name of the controlled input
            Dim sName As String
            If True Then
                Dim sComputeIIDs As String
                
                ' retrieve the controlled input
                Call pGeometricConstructionDefinition.GetControlledInputInfoByIndex(i, sName, sComputeIIDs)
            End If
            
            ' add the name to the collection
            Call pCollection.Add(sName)
        Next
    End If
    
    ' return result
    Set GeometricConstruction_GetNamesOfControlledInputs = pCollection
End Function
Public Sub GeometricConstructionMacro_DeleteOutput(pGeometricConstructionMacro As IJGeometricConstructionMacro, sOutputName As String, vKey As Variant)
    Dim pElementsOfOutputs As IJElements: Set pElementsOfOutputs = pGeometricConstructionMacro.Outputs(sOutputName)
    Dim i As Integer
    For i = 1 To pElementsOfOutputs.Count
        Dim oOutput As Object: Set oOutput = pElementsOfOutputs(i)
        If pElementsOfOutputs.GetKey(oOutput) = "FacePort2" Then
            pElementsOfOutputs.Remove ("FacePort2")
            Dim pObject As IJDObject: Set pObject = oOutput
            Call pObject.Remove
            Exit For
        End If
    Next
End Sub
Public Function GeometricConstructionMacro_IsOutputUnused(pGeometricConstructionMacro As IJGeometricConstructionMacro, pGeometricConstruction As IJGeometricConstruction) As Boolean
    Let GeometricConstructionMacro_IsOutputUnused = pGeometricConstructionMacro.DisableOutput(pGeometricConstruction)
End Function
Function MacroDefinition_GetInputIndex(pGeometricConstructionDefinition As IJGeometricConstructionDefinition, sInputName As String) As Integer
    Call DebugIn("MacroDefinition_GetInputIndex")
    Call DebugInput("InputName", sInputName)
    
    ' prepare result
    Dim iInputIndex As Integer: Let iInputIndex = 0
    
    Dim i As Integer
    For i = 1 To pGeometricConstructionDefinition.InputCount
        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
        Call pGeometricConstructionDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
        If sName = sInputName Then
            ' found result
            Let iInputIndex = i
            Exit For
        End If
    Next

    ' return result
    Let MacroDefinition_GetInputIndex = iInputIndex
    
    Call DebugOutput("InputIndex", MacroDefinition_GetInputIndex)
    Call DebugOut
End Function
Function GeometricConstruction_GetElementsOfAllInputsByNameWithExtendedKey(pGC As IJGeometricConstruction, sPrefixOfName As String) As IJElements
    ' prepare result
    Dim pElementsOfAllInputsByNameWithExtendedKey As IJElements: Set pElementsOfAllInputsByNameWithExtendedKey = New JObjectCollection
    
    ' retrieve definition of the parent Macro
    Dim pGCDefinition As IJGeometricConstructionDefinition: Set pGCDefinition = pGC.definition
    
    Dim i As Integer
    For i = 1 To pGCDefinition.InputCount
        Dim sNameOfInput As String: Let sNameOfInput = GeometricConstructionDefinition_GetInputNameByIndex(pGCDefinition, i)
        If Mid(sNameOfInput, 1, Len(sPrefixOfName)) = sPrefixOfName Then
            Dim pElementsOfInputsByName As IJElements: Set pElementsOfInputsByName = pGC.Inputs(sNameOfInput)
            Dim j As Integer
            For j = 1 To pElementsOfInputsByName.Count
                Dim oInput As Object: Set oInput = pElementsOfInputsByName(j)
                
                ' retrieve the
                ' build the extended key by concatenating the input number and the regular key
                Dim sExtendedKey As String: Let sExtendedKey = CStr(i) + "." + pElementsOfInputsByName.GetKey(oInput)
                
                ' add the member to the global collection
                Call pElementsOfAllInputsByNameWithExtendedKey.Add(oInput, sExtendedKey)
            Next j
        End If
    Next i
  
    ' return result
    Set GeometricConstruction_GetElementsOfAllInputsByNameWithExtendedKey = pElementsOfAllInputsByNameWithExtendedKey
End Function
Function GeometricConstruction_GetElementsOfAllControlledInputsByNameWithNameAsKey(pGC As IJGeometricConstruction, sPrefixOfName As String) As IJElements
    ' prepare result
    Dim pElementsOfAllControlledInputsByNameWithNameAsKey As IJElements: Set pElementsOfAllControlledInputsByNameWithNameAsKey = New JObjectCollection
    
    ' retrieve definition of the parent Macro
    Dim pGCDefinition As IJGeometricConstructionDefinition: Set pGCDefinition = pGC.definition
    
    Dim i As Integer
    For i = 1 To pGCDefinition.ControlledInputCount
        Dim sNameOfControlledInput As String: Let sNameOfControlledInput = GeometricConstructionDefinition_GetControlledInputNameByIndex(pGCDefinition, i)
        Dim bIsControlledInputFound As Boolean: Let bIsControlledInputFound = False
        If sPrefixOfName = "" Then
            Let bIsControlledInputFound = True
        ElseIf Mid(sNameOfControlledInput, 1, Len(sPrefixOfName)) = sPrefixOfName Then
            Let bIsControlledInputFound = True
        End If
        
        If bIsControlledInputFound Then
            Dim pElementsOfControlledInputsByName As IJElements: Set pElementsOfControlledInputsByName = pGC.ControlledInputs(sNameOfControlledInput)
            Dim j As Integer
            For j = 1 To pElementsOfControlledInputsByName.Count
                Dim oControlledInput As Object: Set oControlledInput = pElementsOfControlledInputsByName(j)
                Dim sKey As String: Let sKey = pElementsOfControlledInputsByName.GetKey(oControlledInput)
                Dim sNameAndKey As String: Let sNameAndKey = sNameOfControlledInput
                If Not sKey = "" Then Let sNameAndKey = sNameAndKey + "." + sKey
                
                ' add the controlled input to the global collection with its name as key
                Call pElementsOfAllControlledInputsByNameWithNameAsKey.Add(oControlledInput, sNameAndKey)
            Next j
        End If
    Next i
  
    ' return result
    Set GeometricConstruction_GetElementsOfAllControlledInputsByNameWithNameAsKey = pElementsOfAllControlledInputsByNameWithNameAsKey
End Function
Function GeometricConstructionMacro_GetOutputs(pGCMacro As IJGeometricConstructionMacro, sName As String) As IJElements
    Dim pElementsOfOutputs As IJElements
    On Error Resume Next
    Set pElementsOfOutputs = pGCMacro.Outputs(sName)
    On Error GoTo 0
    
    ' if no elements, then return an enpty collection
    If pElementsOfOutputs Is Nothing Then Set pElementsOfOutputs = New JObjectCollection

    ' return result
    Set GeometricConstructionMacro_GetOutputs = pElementsOfOutputs
End Function
Sub GeometricConstruction_ClearAllInputs(pGC As IJGeometricConstruction, sPrefixOfInputName As String)
    ' retrieve definition of the parent Macro
    Dim pGCDefinition As IJGeometricConstructionDefinition: Set pGCDefinition = pGC.definition
    
    Dim i As Integer
    For i = 1 To pGCDefinition.InputCount
        Dim sNameOfInput As String: Let sNameOfInput = GeometricConstructionDefinition_GetInputNameByIndex(pGCDefinition, i)
        If Mid(sNameOfInput, 1, Len(sPrefixOfInputName)) = sPrefixOfInputName Then
            Call pGC.Inputs(sNameOfInput).Clear
        End If
    Next i
  
End Sub
Function GeometricConstructionDefinition_GetInputIndexByName(pGCDefinition As IJGeometricConstructionDefinition, sName0 As String) As Integer
    Dim iIndex As Integer: Let iIndex = 0
    
    Dim i As Integer
    For i = 1 To pGCDefinition.InputCount
        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
        Call pGCDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
        If sName = sName0 Then
            Let iIndex = i
            Exit For
        End If
    Next
    
    Let GeometricConstructionDefinition_GetInputIndexByName = iIndex
End Function
Function GeometricConstructionDefinition_GetParameterIndexByName(pGCDefinition As IJGeometricConstructionDefinition, sName0 As String) As Integer
    Dim iIndex As Integer: Let iIndex = 0
    
    Dim i As Integer
    For i = 1 To pGCDefinition.ParameterCount
        Dim sName As String, sUserName As String, eParameterType As GCParameterType
        Dim ePrimaryUnits As Units, eSecondaryUnits As Units, eTertiaryUnits As Units, eUnitTypes As UnitTypes
        Dim vDefaultValue As Variant, bIsDisplayedInRBB As Boolean
        Dim sCategoryName As String
        
        Call pGCDefinition.GetParameterInfoByIndex(i, sName, sUserName, eParameterType, ePrimaryUnits, eSecondaryUnits, eTertiaryUnits, eUnitTypes, vDefaultValue, bIsDisplayedInRBB, sCategoryName)
        If sName = sName0 Then
            Let iIndex = i
            Exit For
        End If
    Next
    
    Let GeometricConstructionDefinition_GetParameterIndexByName = iIndex
End Function

Private Function GeometricConstructionDefinition_GetInputNameByIndex(pGCDefinition0 As IJGeometricConstructionDefinition, iIndex As Integer) As String
    Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
    Call pGCDefinition0.GetInputInfoByIndex(iIndex, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
    Let GeometricConstructionDefinition_GetInputNameByIndex = sName
End Function
Private Function GeometricConstructionDefinition_GetControlledInputNameByIndex(pGCDefinition0 As IJGeometricConstructionDefinition, iIndex As Integer) As String
    Dim sName As String, sComputeIIDs As String
    Call pGCDefinition0.GetControlledInputInfoByIndex(iIndex, sName, sComputeIIDs)
    Let GeometricConstructionDefinition_GetControlledInputNameByIndex = sName
End Function
Sub GeometricConstructionMacro_UpdateOutputAsModelBody(pGeometricConstructionMacro As IJGeometricConstructionMacro, sOutputName As String, vKey As Variant, pModelBody1 As IJDModelBody)
    ' check if output with key already exists
    Dim bIsBodyUpToDate As Boolean: Let bIsBodyUpToDate = False
    If True Then
        Dim pElementsOfOutputs As IJElements: Set pElementsOfOutputs = pGeometricConstructionMacro.Outputs(sOutputName)
        Dim i As Integer
        For i = 1 To pElementsOfOutputs.Count
            Dim sKey As String: Let sKey = pElementsOfOutputs.GetKey(pElementsOfOutputs(i))
            If sKey = "" Or sKey = vKey Then
                Let bIsBodyUpToDate = AreModelBodiesSame(pGeometricConstructionMacro.Outputs(sOutputName)(vKey), pModelBody1)
                Exit For
            End If
        Next i
    End If
    
    ' update the output model body
    If Not bIsBodyUpToDate Then pGeometricConstructionMacro.Output(sOutputName, vKey) = pModelBody1
End Sub

'Function GetPortOfAllLateralFaces(pStructGraphConnectable As IJStructGraphConnectable) As IJSurfaceBody
'    ' retrieve visible struct geometry
'    Dim oStructGeometry As Object
'    Set oStructGeometry = Entity_GetGeometry(pStructGraphConnectable)
'
'    ' get port moniker for all lateral faces
'    Dim pMonikerRightOfPort As IMoniker
'    If True Then
'        Dim JS_TOPOLOGY_FILTER_LCONNECT_PRT_LATERAL_LFACES As Integer
'        Let JS_TOPOLOGY_FILTER_LCONNECT_PRT_LATERAL_LFACES = 22
'
'        Dim pMonikerElementsOfPorts As IJMonikerElements
'        Set pMonikerElementsOfPorts = GetPortHelper ().EnumNamedElements(oStructGeometry, JS_TOPOLOGY_FILTER_LCONNECT_PRT_LATERAL_LFACES)
'        If Not pMonikerElementsOfPorts.Count = 1 Then
'            MsgBox "Port for all lateral faces not found"
'            Exit Function
'        End If
'        Set pMonikerRightOfPort = pMonikerElementsOfPorts.Item(1)
'
'        Set pMonikerElementsOfPorts = Nothing
'    End If
'
'    ' create composite moniker
'    Dim pMonikerCompositeOfPort As IMoniker
'    Set pMonikerCompositeOfPort = GetPortHelper ().CreateCompositeMonikerEx(pMonikerRightOfPort)
'
'    ' bind port moniker
'    Dim pPort As IJPort
'    Set pPort = GetPortHelper ().GetPort(pStructGraphConnectable, pMonikerCompositeOfPort)
'
'    ' clean port geometry
'    Dim pModelBody As IJDModelBody
'    Set pModelBody = pPort.Geometry
'
'    Call pModelBody.DebugToSATFile("S:\CommonStruct\Testing\Data\AllLateralFaces.sat")
'    Call SatFile_Clean("S:\CommonStruct\Testing\Data\AllLateralFaces.sat", _
'                       "S:\CommonStruct\Testing\Data\AllLateralFaces-Cleaned.sat")
'
''    Set GetPortOfAllLateralFaces = pPort
'    Set GetPortOfAllLateralFaces = GetAcisHelper().ImportFromSatFile("S:\CommonStruct\Testing\Data\AllLateralFaces-Cleaned.sat")
'End Function
'Private Function Entity_GetGeometry(pAssocRelations As IJDAssocRelation) As Object
'    Dim pRelationshipCol As IJDRelationshipCol
'    Set pRelationshipCol = pAssocRelations.CollectionRelations("IJStructEntityToGeometry", "StructEntityGeometry_ORIG")
'
'    Dim pRelationship As IJDRelationship
'    Set pRelationship = pRelationshipCol.Item(1)
'
'    Set Entity_GetGeometry = pRelationship.Target
'End Function
'Sub SatFile_Clean(sFileIn As String, sFileOut As String)
'    Open sFileIn For Input As #1
'    Open sFileOut For Output As #2
'    While Not EOF(1)
'        Dim sLine As String
'        Line Input #1, sLine
'
'        Dim sStrings() As String
'        sStrings = Split(sLine, " ", 3, vbTextCompare)
'
'        If sStrings(1) = "integer_attrib-name_attrib-gen-attrib" _
'        Or sStrings(1) = "real_attrib-name_attrib-gen-attrib" Then
''            MsgBox sLine
'        Else
'            Print #2, sLine
'        End If
'    Wend
'    Close #1
'    Close #2
'End Sub
Function GeometricConstruction_GetInputToIntersectWithOrToProjectOn(pGC As IJGeometricConstruction) As Object
    ' prepare result
    Dim oInput As Object: Set oInput = Nothing
    
    If pGC.TypeName <> "" Then
        Dim pGCDefinition As IJGeometricConstructionDefinition: Set pGCDefinition = pGC.definition
        Dim i As Integer
        For i = 1 To pGCDefinition.InputCount
            Dim sRoles As String
            Dim sDum1 As String, sDum2 As String, sDum3 As String
            Dim oDum As Object
            Dim bDum As Boolean
            Call pGCDefinition.GetInputGUIstaticInfoByIndex(i, sDum1, oDum, sDum2, bDum, sDum3, sRoles)
            If sRoles <> "" Then
                Dim sArrayOfRoles() As String: sArrayOfRoles = Split(sRoles, ",")
                Dim j As Integer
                For j = 0 To UBound(sArrayOfRoles, 1)
                    Dim sRole As String: Let sRole = sArrayOfRoles(j)
                    If sRole = sROLE_SURFACE_TO_INTERSECT_WITH _
                    Or sRole = sROLE_SURFACE_TO_PROJECT_ON Then
                        Dim sName As String, sPrompt As String, sFilter As String, lMinConnected As Long, lMaxConnected As Long, sComputeIIDs As String
                        Call pGCDefinition.GetInputInfoByIndex(i, sName, sPrompt, sFilter, lMinConnected, lMaxConnected, sComputeIIDs)
                        If pGC.Inputs(sName).Count = 1 Then Set oInput = pGC.Input(sName)
                        Exit For
                    End If
                Next j
            End If
            If Not oInput Is Nothing Then Exit For
        Next i
    End If
    
    ' return result
    Set GeometricConstruction_GetInputToIntersectWithOrToProjectOn = oInput
End Function

'This function is checking if the GC is being mirrored (the flag CTL_FLAG_MIRROR_IN_PROGRESS is set when the GC is being mirrored in the method
'CGeometricConstruction::Adapt() and is reset in the method CGeometricConstruction::Evaluate() after the Definition evaluate)
Public Function GeometricConstructionMacro_IsBeingMirrored(pGC As IJGeometricConstruction) As Boolean
    GeometricConstructionMacro_IsBeingMirrored = False
    Dim oControlFlag As IJControlFlags
    Set oControlFlag = pGC
    If oControlFlag.ControlFlags(CTL_FLAG_MIRROR_IN_PROGRESS) = CTL_FLAG_MIRROR_IN_PROGRESS Then
        GeometricConstructionMacro_IsBeingMirrored = True
    End If
    
End Function
