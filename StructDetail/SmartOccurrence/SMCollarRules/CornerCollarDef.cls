VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CornerCollarDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*********************************************************************************************
'  Copyright (C) 2013, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCollarRules
'  File        : CornerCollarDef.cls
'
'  Description :   Item definition: basic template to define an item
'
'
'  Author      : Alligators
'
'  History     :
'   ---------------------------------------------------------------------
'    DATE        |    AUTHOR    |          COMMENTS
'   ---------------------------------------------------------------------
'   01/APR/2013  -  mchandak    -    Created
'   31/May/2013  -  svsmylav    -    CR-232822 'CM_IsPhysConnNeeded' method is updated
'                                    to set conditionals for Support1/2 PCs.
'
'*********************************************************************************************

Option Explicit

Const m_ItemProgid As String = CUSTOMERID + "CollarRules.CornerCollarDef"
Const m_ItemName As String = CUSTOMERID + "CollarRules.CornerCollarDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "CollarRules\CornerCollarDef.cls"

Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_CORNERFEATURE
    pIH.SetInput INPUT_CORNERCUTOUTPORT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
''MsgBox MODULE & "::ItemAggregator"
    
    Dim sMsg As String
    Dim pPDs As IJDPropertyDescriptions

    pAD.SetCMFinalConstruct m_ItemProgid, "CMFinalConstructCC"
    
    ' The Smart Plate is dependent on the Output of the Symbol
    ' Therefore, we need to set the Property to force the FinalConstruct to be called AFTER_SYMBOL_UPDATE
    ' igPROCESS_PD_AFTER_SYMBOL_UPDATE
    Set pPDs = pAD
    
    ' This is called when CornerCollar is updated on IID_IJStructureMaterial interface
    ' ... Using igPROCESS_PD_AFTER_SYMBOL_UPDATE so that this is called AFTER the Rad Symbol has been processed
    sMsg = "Setting MaterialAndGrade property."
    pPDs.AddProperty "MatAndGrade", 1, IID_IJStructureMaterial, "CM_MatAndGrade", imsCOOKIE_ID_USS_LIB, _
                     igPROCESS_PD_AFTER_SYMBOL_UPDATE
    
    ' This is called when CornerCollar is updated on IJUACornerCollar interface
    sMsg = "Setting Thickness property."
    pPDs.AddProperty "Thickness", 2, IID_IJPlate, "CM_Thickness", imsCOOKIE_ID_USS_LIB, _
                     igPROCESS_PD_BEFORE_SYMBOL_UPDATE
    
    ' This is called when CornerCollar is updated on IJShpStrCollarPlacement interface
    ' ... changes in the System/User Attributes: "SideOfPart", "LongPointEdges"
    sMsg = "Listening to IJShpStrCollarPlacement Interfaces"
    pPDs.AddProperty "ShpStrCollarPlacement", 3, "IJShpStrCollarPlacement", "CM_CornerCollar", imsCOOKIE_ID_USS_LIB, _
                     igPROCESS_PD_BEFORE_SYMBOL_UPDATE
    
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

    iDispId = 1
    sDispId = "CornerCollarLapPC1_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing
    
    iDispId = 2
    sDispId = "CornerCollarLapPC2_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing
    
    iDispId = 3
    sDispId = "CornerCollarButtPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing

    iDispId = 4
    sDispId = "CornerCollarPC_WithSupport1_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing

    iDispId = 5
    sDispId = "CornerCollarPC_WithSupport2_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing

    iDispId = 6
    sDispId = "CornerCollarPC_WithSplitSupport1_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing

    iDispId = 7
    sDispId = "CornerCollarPC_WithSplitSupport2_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnNeeded"
    Set oMemDesc = Nothing
    

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructCC(ByVal pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = m_ItemProgid & "::CMFinalConstructCC"
    On Error GoTo ErrorHandler
''MsgBox METHOD

    Dim sMsg As String
    Dim lSideOfPart As Long
    Dim lProperties As Long
    Dim oCornerCollar As Object
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition
    
    Dim oBoundaries As Collection
    Dim oSDOCollar As StructDetailObjectsex.IJSDOCollar
    
    Set oCornerCollar = pAggregatorDescription.CAO
    If oCornerCollar Is Nothing Then
        Exit Sub
    End If
    
    ' Expect this to be called twice
    ' once for the igPROCESS_PD_BEFORE_SYMBOL_UPDATE (i.e: lProperties = 0)
    '       the Symbol USS needs to know SideOfPart and Collar Thickness
    ' once for the igPROCESS_PD_AFTER_SYMBOL_UPDATE  (i.e: lProperties = 1)
    '       the Place method needs SideOfPart, Collar Thickness and the Symbol Output Contour
    
    Set oSmartOccurrence = oCornerCollar
    Set oSmartItem = oSmartOccurrence.SmartItemObject
    Set oSymbolDefinition = oSmartItem.DefinitionDef
    lProperties = oSmartOccurrence.Properties

    If lProperties = 0 Then
        ' this being called before any of the Parameter Rule Outputs have been set or initialized
        ' since the important inputs are controlled / set by the Parameter Rule
        ' AND the parameter Rule has NOT been executed yet
        ' AND there is nothng that needs to be initialized outside of the Parameter Rule
        ' just return
        Exit Sub
    End If
    
    ' Check/Verify that the CornerCollar symbol has valid Representations and a valid Output Cotnour
    ' if no representations exits, CornerCollar is NOT placed (hooked to the CreatePlatePart semnatic)
    ' if no Symbol Output contour exits, CornerCollar is NOT placed (hooked to the CreatePlatePart semnatic)
    '
    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    Set oSDOCollar.object = oCornerCollar
    lSideOfPart = oSDOCollar.SideOfPart
    oSDOCollar.PlaceCollarAtCorner True, True, True, oBoundaries

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMFinalConstructCC").Number
End Sub

'***********************************************************************
Public Sub CM_IsPhysConnNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = m_ItemProgid & "::CM_IsPhysConnNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    Dim lDIspid As Long
    
    Dim dSideOfPart As Double
    Dim sSideOfPart As String
    Dim dUlength As Double
    Dim dRadius As Double
    
    'Get the Current Selection
    bIsNeeded = False
    
    lDIspid = pMemberDescription.Dispid
    
    Dim oCornerCollar As Object
    Dim oSDOCollar As StructDetailObjectsex.IJSDOCollar
    
    Set oCornerCollar = pMemberDescription.CAO

    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    Set oSDOCollar.object = oCornerCollar
    dSideOfPart = oSDOCollar.SideOfPart
    sSideOfPart = CStr(dSideOfPart)
    
    'Determine which regular and split PCs are needed
    Dim bCFSupport1RegularPC As Boolean
    Dim bCFSupport2RegularPC As Boolean
    Dim bCFSupport1SplitPC As Boolean
    Dim bCFSupport2SplitPC As Boolean
    
    Dim oSupp1Coll As JCmnShp_CollectionAlias
    Dim oSupp2Coll As JCmnShp_CollectionAlias
    Dim oSmartPltUtils As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    Dim bComputeSupp1 As Boolean
    Dim bComputeSupp2 As Boolean
    bComputeSupp1 = IIf(lDIspid = 4 Or lDIspid = 6, True, False)
    bComputeSupp2 = IIf(lDIspid = 5 Or lDIspid = 7, True, False)
    If lDIspid >= 4 And lDIspid <= 7 Then
        Set oSmartPltUtils = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        oSmartPltUtils.GetCornerCollarOverlappingPorts oCornerCollar, _
            bComputeSupp1, bComputeSupp2, Nothing, Nothing, oSupp1Coll, oSupp2Coll
        If Not oSupp1Coll Is Nothing Then
            bCFSupport1RegularPC = IIf(oSupp1Coll.Count >= 1, True, False)
            bCFSupport1SplitPC = IIf(oSupp1Coll.Count > 1, True, False)
        ElseIf Not oSupp2Coll Is Nothing Then
            bCFSupport2RegularPC = IIf(oSupp2Coll.Count >= 1, True, False)
            bCFSupport2SplitPC = IIf(oSupp2Coll.Count > 1, True, False)
        End If
    End If
    
    If InStr(LCase(sSideOfPart), "1") And (lDIspid = 1) Then
        bIsNeeded = True
    ElseIf InStr(LCase(sSideOfPart), "2") And (lDIspid = 2) Then
        bIsNeeded = True
    ElseIf InStr(LCase(sSideOfPart), "3") And (lDIspid = 3) Then
        bIsNeeded = True
    ElseIf lDIspid = 4 Then
       If bCFSupport1RegularPC Then bIsNeeded = True
    ElseIf lDIspid = 5 Then
       If bCFSupport2RegularPC Then bIsNeeded = True
    ElseIf lDIspid = 6 Then
       If bCFSupport1SplitPC Then bIsNeeded = True
    ElseIf lDIspid = 7 Then
       If bCFSupport2SplitPC Then bIsNeeded = True
    Else
        bIsNeeded = False
    End If
            
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_ConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                             ByVal pResourceManager As IUnknown, _
                             ByRef pObject As Object)
Const METHOD = m_ItemProgid & "::CM_ConstructPC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lDIspid As Long

    sMsg = "Creating Insert Plate Items " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.Dispid)
    lDIspid = pMemberDescription.Dispid

        ' Always:
        ' Physical Connection between Profile Edge Feature Port and Insert Plate Lateral Edge
        CreatePhysicalConns pMemberDescription, pResourceManager, _
                                       False, "ButtWeld", pObject

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
    
    Dim oCornerCollar As Object
    Dim oPartWithCornerFeature As Object
    Dim oSDOCollar As StructDetailObjectsex.IJSDOCollar
    
    ' Initialize/Set the Corner Collar Materail,grade, and Thickness
    ' if the Part with the Corner Feature is a Plate
    '... use the same Material, Grade, and (ratio) Thickness
    Set oCornerCollar = pPD.CAO
    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    Set oSDOCollar.object = oCornerCollar

    Set oPartWithCornerFeature = oSDOCollar.Penetrating
    GetCollarThicknessFromCornerFeature oPartWithCornerFeature, dThickness
    SetCollarThickness oCornerCollar, dThickness
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_MatAndGrade(pPD As IJDPropertyDescription, pObject As Object)
Const METHOD = m_ItemProgid & "::CM_MatAndGrade"
    On Error GoTo ErrorHandler
    Dim sMsg As String
''MsgBox METHOD
       
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_CornerCollar(pPD As IJDPropertyDescription, pObject As Object)
Const METHOD = m_ItemProgid & "::CM_CornerCollar"
    On Error GoTo ErrorHandler
    Dim sMsg As String
''MsgBox METHOD
    
    Dim sSideOfPart As String
    Dim lSideOfPart As Long
    Dim lLongShortPointEdges As Long


    Dim sSmartItem As String
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oInputCollection As JCmnShp_CollectionAlias
    
    Dim oCornerCollar As Object
    Dim oSDOCollar As StructDetailObjectsex.IJSDOCollar
    
    ' Need to Update (the Placement of) the Collar when the SideOfPart has been changed
    ' ... Want to Do this AFTER the Symbol has been Updated

    Set oCornerCollar = pPD.CAO

    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    Set oSDOCollar.object = oCornerCollar
    lSideOfPart = oSDOCollar.SideOfPart
    lLongShortPointEdges = oSDOCollar.LongPointEdges

    If lSideOfPart >= 0 Then
       oSDOCollar.PlaceCollarAtCorner True, True, True, oInputCollection
    End If
    
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



