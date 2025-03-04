VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberEncaseSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'History
'    SS          10/29/08      Creation
'    SSuru       12/17/08      Set "Undefined" and proper compute status when rule cannot find an item
'---------------------------------------------------------------------------------------


Option Explicit

Const m_SelectorProgid As String = "SPSEncasementRule.MemberEncaseSel"
Const m_SelectorName As String = "SPSEncasementRule.MemberEncaseSel"
Private Const MODULE = "MemberEncasementSel"

Implements IJDUserSymbolServices
Implements IStructInsulationGraphicInputFilter
Implements IStructInsulationGraphicInputHelper


Public Sub SelectorInputs(pIH As IJDInputsHelper)
Const METHOD = "SelectorInputs"
On Error GoTo ErrorHandler
  
  Exit Sub
  
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
Const METHOD = "SelectorQuestions"
On Error GoTo ErrorHandler
  
  Exit Sub
  
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    Const METHOD = "SelectorLogic"
    On Error GoTo ErrorHandler
    Dim oRC As IJDReferencesCollection
    Dim oObj As Object
    Dim oPart As ISPSMemberPartPrismatic

    Set oRC = GetRefCollection(pSL.SmartOccurrence)
    If oRC.IJDEditJDArgument.GetCount > 2 Then
        Set oObj = oRC.IJDEditJDArgument.GetEntityByIndex(3)
        If Not oObj Is Nothing Then
            If TypeOf oObj Is ISPSMemberPartPrismatic Then
                Set oPart = oObj
            End If
        End If
    End If
    If oPart Is Nothing Then
        MsgBox METHOD & "Cannot get memberpart"
        pSL.Add ""
        Exit Sub
    End If
    
    Dim lInputMemberType As Long
    Dim strInputMemberCrossSec As String
    
    lInputMemberType = oPart.MemberType.TypeCategory
    strInputMemberCrossSec = oPart.CrossSection.CrossSectionTypeName
            
    Dim pRuleSO As IJSmartOccurrence
    Set pRuleSO = pSL.SmartOccurrence
        
    Dim oSmartClass As IJSmartClass
    If Not pRuleSO Is Nothing Then
        Set oSmartClass = pRuleSO.RootSelectionObject
    End If
    
    If oSmartClass Is Nothing Then
        'if here and no smart class, exit
        pSL.Add ""
        Exit Sub
    End If
        
    ' get all the items under smartclass and check the attributes
    ' with respect to input member type and cross-section
    Dim pIRelationHelper    As IMSRelation.DRelationHelper
    Set pIRelationHelper = oSmartClass
    
    Dim pCollOfSmartItems As IMSRelation.DCollectionHelper
'    7FAA6155-07BE-11D2-BC6B-0800360DCD02
    Set pCollOfSmartItems = pIRelationHelper.CollectionRelations(IJDPartClass, RELATED_PARTS)

    Dim bFoundItemForInput As Boolean
    bFoundItemForInput = False
    
    Dim bFoundCurrentItemWithinClass As Boolean
    Dim strCurrentItemName As String
    
    Dim vAttrValue As Variant
    Dim oEncaseItem As IJSmartItem
    Dim oEncaseAttributes As Object
    Dim lMemTypeCategory As Long
    Dim MemCrossSection As String
    Dim lIndex As Long
    Dim oDefAttrs As IJDAttributes
    Dim oStructInsulation As IStructInsulation
    
    Dim oRuleSOAttrs As IJDAttributes
    Set oRuleSOAttrs = pRuleSO
    
    bFoundCurrentItemWithinClass = False
    Set oStructInsulation = pSL.SmartOccurrence
    strCurrentItemName = oStructInsulation.DefinitionName
    
    For lIndex = 1 To pCollOfSmartItems.Count
    
        Set oEncaseItem = pCollOfSmartItems.Item(lIndex)
        
        If Not bFoundCurrentItemWithinClass Then
            If strCurrentItemName = oEncaseItem.Name Then
                bFoundCurrentItemWithinClass = True
            End If
        End If

        ' get the default attributes for copy to occurrance later
        Set oDefAttrs = oEncaseItem
        If Not oEncaseItem Is Nothing Then
            Set oEncaseAttributes = GetAttributeCollection(oEncaseItem, INSULATION_FP_CRITERIA_IFACE)
        End If
        
        If GetAttributeValue(oEncaseAttributes, MEMBER_TYPECATGORY, vAttrValue) Then
            lMemTypeCategory = vAttrValue
        Else
'            MsgBox "attribute IJUAStrMemberFPCriteria::MemberTypeCategory not found"
        End If
        
        If GetAttributeValue(oEncaseAttributes, CROSSSECTION_TYPE, vAttrValue) Then
            MemCrossSection = vAttrValue
        Else
'            MsgBox "attribute IJUAStrMemberFPCriteria::MemberTypeCategory not found"
        End If
                
        If lMemTypeCategory = lInputMemberType And LCase(Trim(MemCrossSection)) = LCase(Trim(strInputMemberCrossSec)) Then
            'set the item
            pSL.Add oEncaseItem.Name
            ' copy the values from Item to Occ
            If Not IsSOOverridden(oRuleSOAttrs.CollectionOfAttributes(SETBACK_INTERFACE)) Then
                CopyValuesFromItemToSO oRuleSOAttrs.CollectionOfAttributes(SETBACK_INTERFACE), oDefAttrs.CollectionOfAttributes(SETBACK_INTERFACE)
            End If
            bFoundItemForInput = True
            Exit For
        Else
        
        End If
        Set oEncaseAttributes = Nothing
        Set oDefAttrs = Nothing
        Set oEncaseItem = Nothing
        
    Next
    
    If bFoundItemForInput = False Then
        ' if not found check by TypeCategory only
        For lIndex = 1 To pCollOfSmartItems.Count
        
            Set oEncaseItem = pCollOfSmartItems.Item(lIndex)
            ' get the default attributes for copy to occurrance later
            Set oDefAttrs = oEncaseItem

            If Not oEncaseItem Is Nothing Then
                Set oEncaseAttributes = GetAttributeCollection(oEncaseItem, INSULATION_FP_CRITERIA_IFACE)
            End If
            
            If GetAttributeValue(oEncaseAttributes, MEMBER_TYPECATGORY, vAttrValue) Then
                lMemTypeCategory = vAttrValue
            Else
'                MsgBox "attribute IJUAStrMemberFPCriteria::MemberTypeCategory not found"
            End If
                    
            If GetAttributeValue(oEncaseAttributes, CROSSSECTION_TYPE, vAttrValue) Then
                MemCrossSection = vAttrValue
            Else
    '            MsgBox "attribute IJUAStrMemberFPCriteria::MemberTypeCategory not found"
            End If
            
            If lMemTypeCategory = lInputMemberType And LCase(Trim(MemCrossSection)) = "*" Then
                'set the item
                pSL.Add oEncaseItem.Name
                ' copy the values from Item to Occ
                If Not IsSOOverridden(oRuleSOAttrs.CollectionOfAttributes(SETBACK_INTERFACE)) Then
                    CopyValuesFromItemToSO oRuleSOAttrs.CollectionOfAttributes(SETBACK_INTERFACE), oDefAttrs.CollectionOfAttributes(SETBACK_INTERFACE)
                End If
                bFoundItemForInput = True
                Exit For
            Else
            
            End If
            Set oEncaseAttributes = Nothing
            Set oDefAttrs = Nothing
            Set oEncaseItem = Nothing

        Next
        
        If bFoundItemForInput = False Then

            oStructInsulation.ComputeStatus = StructInsulationInputHelper_CannotFindCompatibleEncasement

            If oStructInsulation.DefinitionName = oStructInsulation.DefinitionClassName Then    ' placement by rule.
                pSL.Add "Undefined"
            ElseIf bFoundCurrentItemWithinClass Then        ' okay to keep current item
                pSL.Add oStructInsulation.DefinitionName
            Else                                            ' spec changed and current item not in current class
                SPSToDoErrorNotify INSULATIONERROR_TDLCODELISTNAME, INSULATIONERROR_NOSELECTION, oStructInsulation, Nothing
                Err.Raise E_FAIL
            End If
        
        End If
        
    End If
    Exit Sub
    
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        HandleError MODULE, METHOD
    End If
End Sub
  
Public Sub CMSelector(pRep As IJDRepresentation)
Const METHOD = "SelectorLogic"
On Error GoTo ErrorHandler
    
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
End Sub


Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String

  IJDUserSymbolServices_GetDefinitionName = m_SelectorName

End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
    
    Dim inputsProp As IMSDescriptionProperties
    inputsProp = pSelector.IJDInputs.Property
    pSelector.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE
   
   Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD

End Sub


Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

    Dim pDefFact As New DefinitionFactory
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function


Private Function IStructInsulationGraphicInputFilter_FilterGraphicInput(ByVal oGraphicObject As Object, ByVal oDefinition As Object) As Long
Const METHOD = "IStructInsulationGraphicInputFilter_FilterGraphicInput"
On Error GoTo ErrorHandler

    If oGraphicObject Is Nothing Then
        IStructInsulationGraphicInputFilter_FilterGraphicInput = 0
    ElseIf TypeOf oGraphicObject Is ISPSMemberPartPrismatic Then
        IStructInsulationGraphicInputFilter_FilterGraphicInput = 1
    Else
        IStructInsulationGraphicInputFilter_FilterGraphicInput = 0
    End If

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria() As String
    IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria = "ISPSMemberPartPrismatic"
End Function

Private Function IStructInsulationGraphicInputFilter_SingleGraphicInput() As Boolean
    IStructInsulationGraphicInputFilter_SingleGraphicInput = True
End Function

Private Function IStructInsulationGraphicInputHelper_GetGraphicInputs(ByVal pInsulation As StructInsulations.IStructInsulation, ByRef ppElesInputObjects As IJElements) As StructInsulations.StructInsulationGraphicInputHelperStatus
Const METHOD = "IStructInsulationGraphicInputHelper_GetGraphicInputs"
On Error GoTo ErrorHandler

    Dim status As StructInsulations.StructInsulationGraphicInputHelperStatus
    Dim oRC As IJDReferencesCollection
    Dim oObj As Object

    status = StructInsulationInputHelper_UnexpectedError

    Set oRC = GetRefCollection(pInsulation)
    If oRC.IJDEditJDArgument.GetCount > 2 Then
        Set oObj = oRC.IJDEditJDArgument.GetEntityByIndex(3)
        If TypeOf oObj Is ISPSMemberPartPrismatic Then
            Set ppElesInputObjects = New JObjectCollection
            ppElesInputObjects.Add oObj
            status = StructInsulationInputHelper_Ok
        Else
            status = StructInsulationInputHelper_InvalidTypeOfObject
        End If
    Else
        status = StructInsulationInputHelper_BadNumberOfObjects
    End If
    
    IStructInsulationGraphicInputHelper_GetGraphicInputs = status
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IStructInsulationGraphicInputHelper_InitializeEntity(ByVal pInsulation As StructInsulations.IStructInsulation)
Const METHOD = "IStructInsulationGraphicInputHelper_InitializeEntity"
On Error GoTo ErrorHandler
    pInsulation.DisplayAspects = 32     ' initialize the display to use only the insulation aspect
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IStructInsulationGraphicInputHelper_SetGraphicInputs(ByVal pInsulation As StructInsulations.IStructInsulation, ByVal elesInputObjects As IJElements) As StructInsulations.StructInsulationGraphicInputHelperStatus
Const METHOD = "IStructInsulationGraphicInputHelper_SetGraphicInputs"
On Error GoTo ErrorHandler

    Dim status As StructInsulations.StructInsulationGraphicInputHelperStatus
    Dim oRC As IJDReferencesCollection
    Dim oObj As Object

    status = StructInsulationInputHelper_UnexpectedError

    Set oRC = GetRefCollection(pInsulation)
    If elesInputObjects.Count > 0 Then
        Set oObj = elesInputObjects(1)
        If TypeOf oObj Is ISPSMemberPartPrismatic Then
    
            oRC.IJDEditJDArgument.SetEntity 3, oObj, IJGeometry, "StructInsulGeometryRefColl_D"
            oRC.IJDEditJDArgument.SetEntity 4, oObj, ISPSMemberType, "MemberTypeRC_DEST"
            
            AddRelationship oObj, IStructHasInsulation, pInsulation, "Insulation_O", vbNull, False

            status = StructInsulationInputHelper_Ok

        Else
            status = StructInsulationInputHelper_InvalidTypeOfObject
        End If
        
    Else
        status = StructInsulationInputHelper_BadNumberOfObjects
    End If

    IStructInsulationGraphicInputHelper_SetGraphicInputs = status
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

