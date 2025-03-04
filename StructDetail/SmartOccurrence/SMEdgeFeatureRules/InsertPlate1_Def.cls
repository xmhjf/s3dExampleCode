VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsertPlate1_Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Item definition: basic template to define an item
'
Option Explicit

Const m_ItemProgid As String = CUSTOMERID + "EdgeFeatureRules.InsertPlate1_Def"
Const m_ItemName As String = CUSTOMERID + "EdgeFeatureRules.InsertPlate1_Def"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\SmartOccurrence\" + CUSTOMERID + "EdgeFeatureRules\InsertPlate1_Def.cls"

Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    pIH.SetInput "EdgeFeature"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim pPDs As IJDPropertyDescriptions

    pAD.SetCMFinalConstruct m_ItemProgid, "CMFinalConstructIP"
    
    ' The Smart Plate is dependent on the Output of the Symbol
    ' Therefore, we need to set the Property to force the FinalConstruct to be called AFTER_SYMBOL_UPDATE
    Set pPDs = pAD
    
    ' Set the material and grade, Need a valid material before setting Thickness
    sMsg = "Setting MaterialAndGrade property."
    pPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CM_MatAndGrade", imsCOOKIE_ID_USS_LIB, _
                     igPROCESS_PD_AFTER_SYMBOL_UPDATE
    
    ' Set the thickness
    sMsg = "Setting Thickness property."
    pPDs.AddProperty "Thickness", 1, IID_IJPlate, "CM_Thickness", imsCOOKIE_ID_USS_LIB, _
                     igPROCESS_PD_AFTER_SYMBOL_UPDATE
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    ' TODO - Add member description below
    Dim sMsg As String
    Dim sDispId As String
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    ' Insert Plate consist of (3) Member Items,
    ' Index(1) - Physical Connection: Edge Feature and Insert Plate
    ' Index(2) - Physical Connection(?)
    ' Index(2) - Physical Connection(?)
    For iDispId = 1 To 3
        sDispId = "InsertPlate_" & Trim(Str(iDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructItems", imsCOOKIE_ID_USS_LIB)
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsMemberItemNeeded"
        oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
        Set oMemDesc = Nothing
    Next iDispId

    
    Set oMemDesc = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructIP(ByVal pAggregatorDescription As IJDAggregatorDescription)
    On Error GoTo ErrorHandler

    Dim sMsg As String
    Dim sRepName As String
    Dim sSketchPlane As String
   
    Dim lXid1 As Long
    Dim lXid2 As Long
    Dim bOnWeb As Boolean
    Dim bPlateSystem As Boolean
    
    Dim iIndex As Long
    Dim iIndex1 As Long
    Dim nRepIds As Long
    
    Dim oPlate As IJPlate
    Dim oSymbol As IJDSymbol
    Dim oBoundary As Object
    Dim oEdgeFeature As Object
    Dim oInsertPlate As Object
    Dim oProfilePart As Object
    Dim oStiffenedPlate As Object
    Dim oInsertInputs As Collection
    Dim oStructCustomGeometry As IJDStructCustomGeometry
    
    Dim oSO As IJSmartOccurrence
    Dim oCustomCol As Collection
    Dim oBoundaries As Collection
    Dim oEnumPartsUnk As IEnumUnknown
    Dim oRepresentation As IJDRepresentation
    Dim oRepresentations As IJDRepresentations
    Dim oCollectionOfParts As JCmnShp_CollectionAlias
    
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_EdgeFeature As StructDetailObjects.EdgeFeature
    Dim oSDO_PlateSystem As StructDetailObjects.PlateSystem
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    
    Dim ConvertUtils As CCollectionConversions
    Dim oSDSmartPlateOps As GSCADSDCreateModifyUtilities.IJSDSmartPlateOperations
    Dim oStructDetailHelper As StructDetailHelper
    Dim oSDSmartPlateAttributes As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    
    Set oInsertPlate = pAggregatorDescription.CAO
    If oInsertPlate Is Nothing Then
        Exit Sub
    End If
    
    
    ' Need to determine if the Smart Plate symbol has been evaluated yet
    ' the Final Construct method is called ONLY ONCE by design BEFORE the symbol has been evaluated
    ' ... BUT for the Insert Plate, we want the Final Construct method to be call AFTER the symbol has been evaluated
    ' for the Insert Plate,
    ' ... the outputs from the symbol evaluation are used as the boundaries for the Create Plate Part
    ' ... to assist in determining if this is the first (Before evluation) or second (After evluation) call
    ' ... we will check the if the CreatePlatePart.CreatePlatePart_AE.1 relationships have been set
    Set oBoundaries = New Collection
    Set oSDSmartPlateOps = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    
    ' check/verify that oInsertPlate supports IJDStructCustomGeometry interface
    ' get the collcetion of objects for the CreatePlatePart.CreatePlatePart_AE.1 (inputs/outputs)
    If TypeOf oInsertPlate Is IJDStructCustomGeometry Then
        On Error Resume Next
        Set oCustomCol = New Collection
        Set oStructCustomGeometry = oInsertPlate
        oStructCustomGeometry.GetCustomGeometry "CreatePlatePart.CreatePlatePart_AE.1", oCustomCol
        On Error GoTo ErrorHandler
    
        ' CreatePlatePart.CreatePlatePart_AE.1 (inputs/outputs) NOT set ... Assume  first call (Before evluation)
        If oCustomCol.Count < 1 Then
            ' set dummy CreatePlatePart.CreatePlatePart_AE.1 (inputs/outputs)
            ' Reset SOProperties.SO_PROP_FINALCONSTRUCT bit
            ' exit
            oBoundaries.Add oInsertPlate
            oSDSmartPlateOps.PlaceInsertPlate oInsertPlate, oBoundaries
        
            Set oSO = pAggregatorDescription.CAO
            oSO.Properties = oSO.Properties Xor SOProperties.SO_PROP_FINALCONSTRUCT
        
            Exit Sub
        End If
    End If

    
    ' Expect Each Representation to have an Attributed WireBody as an output
    ' Expect the Attributed WireBody to define a boundary for the Smart Plate
    nRepIds = 0
    sRepName = ""
    Set oSDSmartPlateAttributes = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    oSDSmartPlateAttributes.GetRepresentations oInsertPlate, oRepresentations
    If oRepresentations Is Nothing Then
        Exit Sub
    End If
    
    nRepIds = oRepresentations.Count
    If nRepIds < 1 Then
        Exit Sub
    End If
    
    Set oRepresentation = oRepresentations.Item(1)
    If oRepresentation Is Nothing Then
        Exit Sub
    End If
    
    sRepName = oRepresentation.Name
    
    ' First item in the Boundary List should be the sketching Plane
    ' The sketching Plane should be defined on the First Representation
        
    ''' oSDSmartPlateAttributes.GetRepresentationName oInsertPlate, 1, sRepName
    
    ' First item in the Boundary List should be the sketching Plane
    ' The sketching Plane should be defined on the First Representation
    sSketchPlane = sRepName & ":Plane"
    Set oSymbol = oInsertPlate
    Set oBoundary = oSymbol.BindToOutput(sRepName, sSketchPlane)
    If Not oBoundary Is Nothing Then
        oBoundaries.Add oBoundary
        Set oBoundary = Nothing
    End If
    Set oRepresentation = Nothing
    
    ' Check if Defining Edge Feature on on the Profile's Web
    oSDSmartPlateAttributes.GetInputs_InsertPlate oInsertPlate, oInsertInputs
    Set oEdgeFeature = oInsertInputs.Item(1)
    bOnWeb = IsEdgeFeatureOnWeb(oEdgeFeature, lXid1, lXid2)
    
    ' From the Edge Feature: Get Profile Part, Profile Material, Stiffened Plate
    Set oSDO_EdgeFeature = New StructDetailObjects.EdgeFeature
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_EdgeFeature.object = oEdgeFeature
    Set oProfilePart = oSDO_EdgeFeature.GetPartObject
    
    Set oSDO_ProfilePart.object = oProfilePart
    oSDO_ProfilePart.GetStiffenedPlate oStiffenedPlate, bPlateSystem

    If bOnWeb Then
        ' For defining Web Edge Feature, use "Boundary_Inside" Representation and Stiffened Plate Part
        For iIndex = 1 To nRepIds
            sRepName = ""
            Set oRepresentation = oRepresentations.Item(iIndex)
            If Not oRepresentation Is Nothing Then
                sRepName = oRepresentation.Name
                If LCase(Trim(sRepName)) = LCase("Boundary_Inside") Then
                    oSDSmartPlateAttributes.GetAttributedWire oInsertPlate, sRepName, oBoundary
                    If Not oBoundary Is Nothing Then
                        oBoundaries.Add oBoundary
                        Set oBoundary = Nothing
                    End If
                End If
                Set oRepresentation = Nothing
            End If
        Next iIndex
        
        ' If Stiffened Plate is a Plate System
        ' Get the Plate System's Plate Parts to be use as boundaries
        If bPlateSystem Then
            Set oStructDetailHelper = New StructDetailHelper
            oStructDetailHelper.GetPartsDerivedFromSystem oStiffenedPlate, oEnumPartsUnk, True

            If Not oEnumPartsUnk Is Nothing Then
                Set ConvertUtils = New CCollectionConversions
                ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPartsUnk, oCollectionOfParts
            
                For iIndex1 = 1 To oCollectionOfParts.Count
                    oBoundaries.Add oCollectionOfParts.Item(iIndex1)
                Next iIndex1
            End If
            
        Else
            oBoundaries.Add oStiffenedPlate
        End If
        
    Else
        ' For defining Flange Edge Feature, use all Representations
        ' Loop thru each Representation, adding the Attributed WireBody to the Boundary List
        ' ... If not all Representations are to be used as boundaries
        ' ... Check the Representation Name to verify if it is to be included in the Boundary List
        For iIndex = 1 To nRepIds
            sRepName = ""
            Set oRepresentation = oRepresentations.Item(iIndex)
            If Not oRepresentation Is Nothing Then
                sRepName = oRepresentation.Name
                If Len(Trim(sRepName)) > 0 Then
                    oSDSmartPlateAttributes.GetAttributedWire oInsertPlate, sRepName, oBoundary
                    If Not oBoundary Is Nothing Then
                        oBoundaries.Add oBoundary
                        Set oBoundary = Nothing
                    End If
                End If
                Set oRepresentation = Nothing
            End If
        Next iIndex
    End If
    
    oSDSmartPlateOps.PlaceInsertPlate oInsertPlate, oBoundaries
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMFinalConstructIP").Number
End Sub

'***********************************************************************
Public Sub CM_IsMemberItemNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = m_ItemProgid & "::CM_IsMemberItemNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    Dim lDispId As Long
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper

    ' Get the Current Selection
    bIsNeeded = False
    lDispId = pMemberDescription.Dispid
    If lDispId = 1 Then
        ' Always: Physical Connection: Edge Feature and Insert Plate
        bIsNeeded = True
        Exit Sub
    End If
    
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_CopyQuestionsToMember(pMemberDescription As IJDMemberDescription)
Const METHOD = m_ItemProgid & "::CM_CopyQuestionsToMember"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lDispId As Long

    Dim oItemObject As Object
    
    sMsg = "Setting Questions on Insert Plate Items " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.Dispid)
    lDispId = pMemberDescription.Dispid
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_ConstructItems(ByVal pMemberDescription As IJDMemberDescription, _
                             ByVal pResourceManager As IUnknown, _
                             ByRef pObject As Object)
Const METHOD = m_ItemProgid & "::CM_ConstructItems"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lDispId As Long
    
    sMsg = "Creating Insert Plate Items " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.Dispid)
    lDispId = pMemberDescription.Dispid
    
    If lDispId = 1 Then
        ' Always:
        ' Physical Connection between Profile Edge Feature Port and Insert Plate Lateral Edge
        CreatePC_FeatureAndInsertPlate pMemberDescription, pResourceManager, _
                                       False, "ButtWeld", pObject
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_Thickness(pPD As IJDPropertyDescription, pObject As Object)
Const METHOD = m_ItemProgid & "::CM_Thickness"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dThickness As Double
    
    Dim oEdgeFeature As Object
    Dim oInsertPlate As Object
    Dim oInsertInputs As Collection
    
    Dim oSDSmartPlateAttributes As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    
    Set oInsertPlate = pPD.CAO
    Set oSDSmartPlateAttributes = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    oSDSmartPlateAttributes.GetInputs_InsertPlate oInsertPlate, oInsertInputs
    Set oEdgeFeature = oInsertInputs.Item(1)
    
    GetInsertPlateThicknessFromEdgeFeature oEdgeFeature, dThickness
    SetInsertPlateThickness oInsertPlate, dThickness
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_MatAndGrade(pPD As IJDPropertyDescription, pObject As Object)
Const METHOD = m_ItemProgid & "::CM_MatAndGrade"
    On Error GoTo ErrorHandler
    Dim sMsg As String
       
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
Const METHOD = m_ItemProgid & "::_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation
    
    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous member descriptions
    ItemMembers pMDs
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
Const METHOD = m_ItemProgid & "::InstanciateDefinition"
    On Error GoTo ErrorHandler
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function
