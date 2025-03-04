VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFlangeNonPenetratedDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenFlangeNonPenetratedDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    19/Oct/2011 - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'   22/Nov/2011 - pnalugol
'           CR-CP-205624 - Removing functionality to defer corner feature from webcut to flange cut
'    03/August/2015 - MDT TR-274273 Change SM orientation from vertical to horizontal- deletes PCs under GenMemACs
'*********************************************************************************************
Const m_sClassName As String = "GenFlangeNonPenetratedDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Definition Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
 
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description :
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    pAD.UserTypeClsid = CA_FLANGECUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    pAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"

    Dim pPDs As IJDPropertyDescriptions
    Set pPDs = pAD
    pPDs.RemoveAll
    
    ' Set the FlangeCutting Behaviour
    sMsg = "Setting Flange Cutting Behaviour property."
    pPDs.AddProperty "CuttingBehavior", 1, "IJUAMbrEndCuts", "CM_SetFlangeCuttingBehaviour", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sDispId As String
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    'Physical Connection for the Flange Cut
    Set oMemDesc = pMDs.AddMember("FlangePC", 1, "CM_ConstructPhysConns", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CM_NeedToComputeNonPenPC", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    Set oMemDesc = pMDs.AddMember("InsideCorner", 2, "CM_ConstructInsideCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsInsideCornerFeatureNeeded"
    Set oMemDesc = Nothing
   
   'removed methods to create corner features deferred to flange as part of CR 205624
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
'*********************************************************************************************
' Method      : CM_ConstructPhysConns
' Description :
'
'*********************************************************************************************
Public Function CM_ConstructPhysConns(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructPhysConns"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)

    CreateModify_PhysConns pMemberDescription, pResourceManager, pObject
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Smart Definition as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
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
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Dim pAPDs As IJDPropertyDescriptions
    Set pAD = pDefinition
    Set pAPDs = pDefinition
    pAPDs.RemoveAll ' Removes all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Removes all the previous Member descriptions
    ItemMembers pMDs
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                            ByVal defParams As Variant, _
                                                            ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_DefinitionProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CreateModify_PhysConns
' Description :
'
'*********************************************************************************************
Private Function CreateModify_PhysConns(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
                                        
    Const METHOD = m_DefinitionProgid & "::CreateModify_PhysConns"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)

    ' --------------------------------------------
    ' Find out if this is the top or bottom flange
    ' --------------------------------------------
    Dim sBottomFlange As String
    Dim isBottom As Boolean
    isBottom = False
    
    GetSelectorAnswer pMemberDescription, "BottomFlange", sBottomFlange
    If sBottomFlange = "Yes" Then
        isBottom = True
    End If
    
    ' ----------------------------------------------------------------------
    ' If the bounding object is a member, find out which edge is intersected
    ' ----------------------------------------------------------------------
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_FlangeCut.object = pMemberDescription.CAO
    Set oSDO_WebCut.object = oSDO_FlangeCut.WebCut
    
    Dim boundingEdgeId As eBounding_Edge
    Dim mappedEdgeId As JXSEC_CODE
    
    boundingEdgeId = None
    
    If TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic Then
        GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, isBottom
    End If

    ' -----------------------------------------------------------
    ' Create a Tee or Butt weld, depending on bounding edge width
    ' -----------------------------------------------------------
    Dim oDefCM As New EndCutDefCM
    If boundingEdgeId = Bottom_Flange_Right Or boundingEdgeId = Top_Flange_Right Then
        Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "ButtWeld", mappedEdgeId, 8193)
    ElseIf Not (boundingEdgeId = Above Or boundingEdgeId = Below Or boundingEdgeId = None) Then
        Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", mappedEdgeId, 8193)
    ElseIf TypeOf oSDO_FlangeCut.Bounding Is IJPlatePart Then
            Dim oSDO_PlatePart As New StructDetailObjects.PlatePart
            Dim oBaseSurface As IJSurfaceBody
            Dim oFlangeSurface As IJSurfaceBody
            Dim oSGOModelBodyUtilities As SGOModelBodyUtilities
            Dim oClosestPosition1 As IJDPosition
            Dim oClosestPosition2 As IJDPosition
            Dim dMinDist As Double
            Dim FlangeNormal As IJDVector
            Dim oNormalToBase As IJDVector
            Dim dDotProduct As Double
            
            Set oClosestPosition1 = New DPosition
            Set oClosestPosition2 = New DPosition
            Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
            Set oSDO_PlatePart.object = oSDO_FlangeCut.Bounding
            Set oBaseSurface = oSDO_PlatePart.BasePort(BPT_Base).Geometry
            If HasTopFlange(oSDO_FlangeCut.Bounded) Then
               Set oFlangeSurface = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_TOP).Geometry
            ElseIf HasBottomFlange(oSDO_FlangeCut.Bounded) Then
               Set oFlangeSurface = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_BOTTOM).Geometry
            End If
            Set FlangeNormal = New dVector

           
            oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oBaseSurface, oFlangeSurface, oClosestPosition1, oClosestPosition2, dMinDist
            oBaseSurface.GetNormalFromPosition oClosestPosition1, oNormalToBase
            oFlangeSurface.GetNormalFromPosition oClosestPosition2, FlangeNormal
              
            dDotProduct = FlangeNormal.Dot(oNormalToBase)
            
             Dim oBoundingTopoPort As IJStructPort
             Set oBoundingTopoPort = oSDO_WebCut.BoundingPort
                 If oBoundingTopoPort.ContextID And CTX_LATERAL Then
                      If dDotProduct = 1 Or dDotProduct = -1 Then
                          Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "ButtWeld", mappedEdgeId, 8193)
                      Else
                          Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", mappedEdgeId, 8193)
                      End If
                 Else
                     Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", mappedEdgeId, 8193)
                 End If

    Else
        Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", -1, 8193)
    End If
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function


Public Sub CM_IsInsideCornerFeatureNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_DefinitionProgid & "::CM_IsInsideCornerFeatureNeeded"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sAnswer As String
    GetSelectorAnswer oMD, "CreateInsideCornerFeature", sAnswer

    If UCase(sAnswer) = "YES" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function CM_ConstructInsideCornerFeature(ByVal oMD As IJDMemberDescription, _
                                                 ByVal oResMgr As IUnknown, _
                                                 ByRef oObject As Object)
                                                 
    Const METHOD = m_DefinitionProgid & "::CM_ConstructInsideCornerFeature"
    
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' ---------------------------------------------------------------
    ' Create the feature between the flange cut and the input web cut
    ' ---------------------------------------------------------------
    Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
    Set oSDOFlangeCut.object = oMD.CAO

    Set oObject = CreateCornerFeatureBetweenTwoEndCuts(oSDOFlangeCut.WebCut, oSDOFlangeCut.object, oResMgr, oSDOFlangeCut.object)

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function


'******************************************************************************************
' Method:
' CM_SetFlangeCuttingBehaviour
'
' *******************************************************************************************
Public Sub CM_SetFlangeCuttingBehaviour(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CM_SetFlangeCuttingBehaviour"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a Flange Cutting Behaviour"
            
    Dim oMemberDescriptions As IJDMemberDescriptions
    Set oMemberDescriptions = pPropertyDescriptions.definition
    
    Dim oMemberDescription As IJDMemberDescription
    Dim oEndCutDefCM As New EndCutDefCM
    Dim bIsNeeded As Boolean
    bIsNeeded = True
    
    Dim IsProExists As Boolean
    
    'Check if the attribute is exists on flange cut
    IsProExists = Has_Attribute(pPropertyDescriptions.CAO, "CuttingBehavior")
    
    'Exit if attribute is not bulkloaded
    If Not IsProExists Then Exit Sub
        
    'Call PC Item member conditional and check weather it is True or False
    'Based on this Set FlangeCut Cutting Behaviour
    Set oMemberDescription = oMemberDescriptions.ItemByDispid(1)
    oMemberDescription.CAO = pPropertyDescriptions.CAO

    oEndCutDefCM.CM_IsPhysConnsNeeded oMemberDescription, bIsNeeded
    
    'Set Proper Flange Cutting Behaviour
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
    
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

