VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebNonPenetratedDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenWebNonPenetratedDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    03/August/2015 - MDT TR-274273 Change SM orientation from vertical to horizontal- deletes PCs under GenMemACs
'*********************************************************************************************
Const m_sClassName As String = "GenWebNonPenetratedDef"
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
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
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
    
    pAD.UserTypeClsid = CA_WEBCUT
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    
    pAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    pAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
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

    'Physical Connection for the Web Cut
    iDispId = 1
    sDispId = "GenericWeb_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPhysConns", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CM_NeedToComputeNonPenPC", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Top Flange Cut
    iDispId = 2
    sDispId = "GenericTopFlange_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructTopFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID + "MbrEndCut.EndCutDefCM", "CM_CopyQuestionsToTopFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oMemDesc = Nothing
    
    'Bottom Flange Cut
    iDispId = 3
    sDispId = "GenericBottomFlange_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructBottomFlangeCut", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomFlangeCutNeeded"
    oMemDesc.SetCMSetInputs CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_CopyQuestionsToBottomFlangeCut"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oMemDesc = Nothing
    
    'Top PC
    iDispId = 4
    sDispId = "GenericNonPenTopPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructNonPenTopPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopPCNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CM_NeedToComputeNonPenPC", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Bottom PC
    iDispId = 5
    sDispId = "GenericNonPenBottomPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructNonPenBottomPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomPCNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputePC", 1, "IJSmartOccurrence", "CM_NeedToComputeNonPenPC", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Top Corner
    iDispId = 6
    sDispId = "GenericNonPenTopCorner_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructNonPenTopCorner", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsTopCornerNeeded"
    Set oMemDesc = Nothing
    
    'Bottom Corner
    iDispId = 7
    sDispId = "GenericNonPenBottomCorner_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructNonPenBottomCorner", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsBottomCornerNeeded"
    Set oMemDesc = Nothing
    
    ' Added below four item members for creation of PCs between Gusset plate face and Bounded member flange
    'With Plate Base
    iDispId = 8
    sDispId = "GenericNonPenTopAndBtmPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndMbr", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsPCWithPlateNeeded"
    Set oMemDesc = Nothing
   
    'With offset plate
    iDispId = 9
    sDispId = "GenericNonPenTopAndBtmPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndMbr", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsPCWithPlateNeeded"
    Set oMemDesc = Nothing
      
    iDispId = 10
    sDispId = "GenericNonPenTopAndBtmPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndMbr", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsPCWithPlateNeeded"
    Set oMemDesc = Nothing

    iDispId = 11
    sDispId = "GenericNonPenTopAndBtmPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndMbr", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_IsPCWithPlateNeeded"
    Set oMemDesc = Nothing
      
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
    sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    CreateModify_PhysConns pMemberDescription, pResourceManager, pObject
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

'*********************************************************************************************
' Method      : CM_MigratePhysConn
' Description :
'
'*********************************************************************************************
Public Sub CM_MigratePhysConn(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = m_DefinitionProgid & "::CM_MigratePhysConn"
    On Error GoTo ErrorHandler
    
    ' The Physical Connection Inputs have been Migrated at same time with the EndCut Inputs
    ' Therefore, there should be nothing to do here
    Dim sMsg As String
    sMsg = "Migrating Physical Connection Inputs"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

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
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean

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
    Dim lEdgeId As Long
    Dim lDispId As Long
    
    Dim eBoundingPort As eUSER_CTX_FLAGS
    Dim eBoundedSubPort As JXSEC_CODE
    
    lDispId = pMemberDescription.dispid
    lEdgeId = 8193
    eBoundingPort = CTX_NOP
    eBoundedSubPort = lEdgeId

    sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(lDispId)
    
    If pResourceManager Is Nothing Then
        sMsg = "Modifying Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
        Modify_PhysConn pMemberDescription, eBoundingPort, eBoundedSubPort
        
    Else
        sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
                        
            
'         ----------------------------------------------------------------------
'         If the bounding object is a member, find out which edge is intersected
'         ----------------------------------------------------------------------
        Dim oSDO_WebCut As New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = pMemberDescription.CAO
        Dim oDefCM As New EndCutDefCM
        Dim boundingEdgeId As eBounding_Edge
        Dim mappedEdgeId As JXSEC_CODE
    
        boundingEdgeId = None
       If TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Then
            GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, False
       End If
       
        Dim bIsTubalar As Boolean
        bIsTubalar = False
        bIsTubalar = IsTubularMember(oSDO_WebCut.Bounded)
        
        ' -----------------------------------------------------------
        ' Create a Tee or Butt weld, depending on bounding edge width
        ' -----------------------------------------------------------
        If boundingEdgeId = eBounding_Edge.Bottom_Flange_Right Or boundingEdgeId = eBounding_Edge.Top_Flange_Right Then
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "ButtWeld", eBoundingPort, eBoundedSubPort)
        ElseIf Not (boundingEdgeId = eBounding_Edge.Above Or boundingEdgeId = eBounding_Edge.Below Or boundingEdgeId = eBounding_Edge.None) Then
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "TeeWeld", eBoundingPort, eBoundedSubPort)
                                      
        ElseIf TypeOf oSDO_WebCut.Bounding Is IJPlatePart And Not bIsTubalar Then
            Dim oSDO_PlatePart As StructDetailObjects.PlatePart
            Dim oBoundingPart As Object
            Dim oBaseSurface As IJSurfaceBody
            Dim oWebLeftSurface As IJSurfaceBody
            Dim oSGOModelBodyUtilities As SGOModelBodyUtilities
            Dim oClosestPosition1 As IJDPosition
            Dim oClosestPosition2 As IJDPosition
            Dim dMinDistToWebLeft As Double
            Dim WLNormal As IJDVector
            Dim oNormalToBase As IJDVector
            Dim dWLDotProduct As Double
            
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = pMemberDescription.CAO
            Set oClosestPosition1 = New DPosition
            Set oClosestPosition2 = New DPosition
            Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
            Set oBoundingPart = oSDO_WebCut.Bounding
            Set oSDO_PlatePart = New StructDetailObjects.PlatePart
            Set oSDO_PlatePart.object = oBoundingPart
            Set oBaseSurface = oSDO_PlatePart.BasePort(BPT_Base).Geometry
            Set oWebLeftSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_LEFT).Geometry
            Set WLNormal = New dVector

           
            oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oBaseSurface, oWebLeftSurface, oClosestPosition1, oClosestPosition2, dMinDistToWebLeft
            oBaseSurface.GetNormalFromPosition oClosestPosition1, oNormalToBase
            oWebLeftSurface.GetNormalFromPosition oClosestPosition2, WLNormal
              
            dWLDotProduct = WLNormal.Dot(oNormalToBase)
            
             Dim oBoundingTopoPort As IJStructPort
             Set oBoundingTopoPort = oSDO_WebCut.BoundingPort
                 If oBoundingTopoPort.ContextID And CTX_LATERAL Then
                      If dWLDotProduct = 1 Or dWLDotProduct = -1 Then
                          Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "ButtWeld", mappedEdgeId, 8193)
                      Else
                          Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", mappedEdgeId, 8193)
                      End If
                 Else
                     Set pObject = oDefCM.CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", mappedEdgeId, 8193)
                 End If
                
                                 
        Else
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "TeeWeld", eBoundingPort, eBoundedSubPort)
        End If
         End If
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

'*********************************************************************************************
' Method      : CM_ConstructNonPenTopPC
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructNonPenTopPC(ByVal oMD As IJDMemberDescription, ByVal oResMgr As IUnknown, ByRef oObject As Object)
    
    Dim oDefCM As New EndCutDefCM
    Set oObject = oDefCM.ConstructNonPenTopOrBottomPC(oMD, oResMgr, False)
                               
End Sub

'*********************************************************************************************
' Method      : CM_ConstructNonPenBottomPC
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructNonPenBottomPC(ByVal oMD As IJDMemberDescription, ByVal oResMgr As IUnknown, ByRef oObject As Object)
    
    Dim oDefCM As New EndCutDefCM
    Set oObject = oDefCM.ConstructNonPenTopOrBottomPC(oMD, oResMgr, True)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructNonPenTopCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructNonPenTopCorner(ByVal oMD As IJDMemberDescription, ByVal oResMgr As IUnknown, ByRef oObject As Object)
       
    Dim oDefCM As New EndCutDefCM
    Set oObject = oDefCM.CAConstruct_MbrCornerFeature(oMD, oResMgr, JXSEC_TOP, JXSEC_IDEALIZED_BOUNDARY)
                               
End Sub

'*********************************************************************************************
' Method      : CM_ConstructNonPenBottomCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructNonPenBottomCorner(ByVal oMD As IJDMemberDescription, ByVal oResMgr As IUnknown, ByRef oObject As Object)

    Dim oDefCM As New EndCutDefCM
    Set oObject = oDefCM.CAConstruct_MbrCornerFeature(oMD, oResMgr, JXSEC_BOTTOM, JXSEC_IDEALIZED_BOUNDARY)

End Sub

