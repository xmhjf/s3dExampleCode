VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StructPenetrationSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.StructPenetrationSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.StructPenetrationSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StructPenetrationSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", strError).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Dim strError As String
    ' Manual Slot
    ' question/answer to control creation of Primary,Secondary Clips
    pQH.SetQuestion "AllowedClips", "Default", "AllowedClipsCodeList"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    
    strError = "Retrieving assembly connection from Smart Occurrence."
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence

    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance

    Dim sSectionAlias As String
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    
    If Not oSlotmappingRule Is Nothing Then
        oSlotmappingRule.GetSectionAlias oAssyConn.ConnectedObject2, oAssyConn.ConnectedObject1, sSectionAlias, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
        'Either Slot Mapping Rule Does Not Exist in the Catalog - Default to UnknownAlias
        ' or inverted LC
        sSectionAlias = "UnknownAlias"
    End If
    
    If oAssyConn.ConnectedObject2 Is oFlange Then
        'Do not create an AC if the Penetrating Object is the Flange Object
        'Need to update Slot between Penetrated object and Web object
        '
        Dim oFeatures As Collection
        
        If oAssyConn.ConnectedObject1Type = SDOBJECT_PLATE Then
           ' Get slots on plate
           Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
           
           Set oPlatePartWrapper.object = oAssyConn.ConnectedObject1
           Set oFeatures = oPlatePartWrapper.PlateFeatures
           
           Set oPlatePartWrapper = Nothing
        Else
           ' Get slots on profile
           Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
           
           Set oProfilePartWrapper.object = oAssyConn.ConnectedObject1
           Set oFeatures = oProfilePartWrapper.ProfileFeatures
           Set oProfilePartWrapper = Nothing
        End If
        
        If Not oFeatures Is Nothing Then
           Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools
           Dim oSlotWrapper As New StructDetailObjects.Slot
           Dim oFeature As Object
           Dim oStructFeature As IJStructFeature
           
           Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
           For Each oFeature In oFeatures
              If TypeOf oFeature Is IJStructFeature Then
                 Set oStructFeature = oFeature
                 Set oFeature = Nothing
                 If oStructFeature.get_StructFeatureType = SF_Slot Then
                    Set oSlotWrapper.object = oStructFeature
                    If oSlotWrapper.Penetrating Is oWeb Then
                       If oSlotWrapper.Penetrated Is oAssyConn.ConnectedObject1 Then
                          ' Update slot
                          oStructAssocTools.UpdateObject oStructFeature, _
                                                         "{6034AD40-FA0B-11d1-B2FD-080036024603}" ' IJStructGeometry
                          Set oStructFeature = Nothing
                          Exit For
                       End If
                    End If
                 End If
                 Set oStructFeature = Nothing
              End If
           Next
           Set oStructAssocTools = Nothing
           Set oSlotWrapper = Nothing
           Set oFeatures = Nothing
        End If
        
        Set oWeb = Nothing
        Set oFlange = Nothing
        
        Exit Sub
    End If
    
    If sSectionAlias = "UnknownAlias" Or oAssyConn.ConnectedObject2 Is oWeb Then

        strError = "Querying assembly connection for penetrated and penetrating object types."
        Select Case oAssyConn.ConnectedObject2Type  'Should be penetrating object
            Case SDOBJECT_PLATE
                Select Case oAssyConn.ConnectedObject1Type  'Should be penetrated object
                    Case SDOBJECT_PLATE
                        If sSectionAlias = "UnknownAlias" Then
                            pSL.Add "PlateThruPlate"
                        Else
                            If sSectionAlias = "FB" Then
                                pSL.Add "StiffenerThruPlate"
                                pSL.Add "PlateThruPlate"
                            Else
                                pSL.Add "StiffenerThruPlate"
                                pSL.Add "PlateThruPlate"
                            End If
                        End If
'                        If (CheckRegistryEntry("INTELLISHIP_NEWTRIM", 0) = 0) Then
'                            pSL.Add "PlateThruPlate"
'                        Else
'                            pSL.Add "PlateThruPlateCutout"
'                        End If
                    Case SDOBJECT_STIFFENER
                        ' Manual Slot, testing only
                        ' Allow Profile bounded by Plate case
                        pSL.Add "StiffenerThruPlate"
                    Case SDOBJECT_BEAM
    '                    pSL.Add "PlateThruBeam"
                End Select
            Case SDOBJECT_STIFFENER
                Select Case oAssyConn.ConnectedObject1Type  'Should be penetrated object
                    Case SDOBJECT_PLATE
                        pSL.Add "StiffenerThruPlate"
                    Case SDOBJECT_STIFFENER
                        pSL.Add "StiffenerThruStiffener"
                    Case SDOBJECT_BEAM
    '                    pSL.Add "StiffenerThruBeam"
                End Select
            Case SDOBJECT_BEAM
                Select Case oAssyConn.ConnectedObject1Type  'Should be penetrated object
                    Case SDOBJECT_PLATE
    '                    pSL.Add "BeamThruPlate"
                    Case SDOBJECT_STIFFENER
    '                    pSL.Add "BeamThruStiffener"
                    Case SDOBJECT_BEAM
    '                    pSL.Add "BeamThruBeam"
                End Select
            Case SDOBJECT_SEAM
                Select Case oAssyConn.ConnectedObject1Type  'Should be penetrated object
                    Case SDOBJECT_PLATE
                        pSL.Add "SeamThruPlate"
                    Case SDOBJECT_STIFFENER
                        pSL.Add "SeamThruStiffener"
                    Case SDOBJECT_BEAM
    '                    pSL.Add "SeamThruBeam"
                End Select
            Case Else
                strError = "Unknown value:  oAssyConn.ConnectedObject2Type=" & oAssyConn.ConnectedObject2Type
                GoTo ErrorHandler
        End Select
    
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
End Sub
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Private Function CheckRegistryEntry(strEntry As String, dNewValue As Long) As Long
On Error GoTo ErrorHandler

    Dim strError As String
    
    Dim oRegistry As IMSRegistryHandler.IJRegistry
    Dim oCurrNode As IMSRegistryHandler.IJRegistryNode
    Dim oCurrSettings As IMSRegistryHandler.IJRegistryItems
    Dim oItem As IMSRegistryHandler.IJRegistryItem
    Dim strDefaultValuesKey As String
    
    strError = "Getting initial node"
    'Initialisation
    Set oRegistry = New Registry
    strDefaultValuesKey = "\HKEY_LOCAL_MACHINE\Software\Intergraph\Applications\Environments\ShipStructure\Parameters"
    
    ' Get the node. If it doesnt exist, create it
    On Error Resume Next
    Set oCurrNode = oRegistry.GetNodeFromName(strDefaultValuesKey)
    On Error GoTo ErrorHandler
      
    If oCurrNode Is Nothing Then
        strError = "Adding node"
        Set oCurrNode = oRegistry.AddKey(strDefaultValuesKey)
        oCurrNode.Contents.Add strEntry, dNewValue, IMSREG_VALUE_DWORD
    Else
        strError = "Getting node Items"
        ' Get the different node items
        Set oCurrSettings = oCurrNode.Contents
    
        ' Get the corresponding item. If it doesn't exist, create it
        On Error Resume Next
        Set oItem = oCurrSettings.Item(strEntry)
        On Error GoTo ErrorHandler
        
        If oItem Is Nothing Then
            strError = "Adding key"
            oCurrSettings.Add strEntry, dNewValue, IMSREG_VALUE_DWORD
        Else
            strError = "Reading key"
            dNewValue = oItem.Value
        End If
    End If
    
    strError = "Return value"
    CheckRegistryEntry = dNewValue
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CheckRegistryEntry").Number

End Function



Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



