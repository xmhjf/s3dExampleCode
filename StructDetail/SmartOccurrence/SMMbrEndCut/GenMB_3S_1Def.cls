VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenMB_3S_1Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2014-16, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenMB_3S_1Def.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'           23/May/2014 : Created
'           22/April/2015 - modified MDT TR-271041.Added properties to update the PCs when bounded member is rotated after AC placement
'           07/Jan/2016   - Modified GHM TR-279987 Added new properties to Corner feature to support rotation of bounded member
'*********************************************************************************************
Const m_sClassName As String = "GenMB_3S_1Def"
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
    
    pAD.UserTypeClsid = CA_WEBCUT
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
' Description : Three PC items and One Corner Feature
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    
    sError = "Adding Items"
    
    'Physical Connections With First Bounding Surface
    Set oMemDesc = pMDs.AddMember("FirstPC", 1, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute", 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Physical Connections With Second Bounding Surface
    Set oMemDesc = pMDs.AddMember("SecondPC", 2, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute", 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Physical Connections With Third Bounding Surface
    Set oMemDesc = pMDs.AddMember("ThirdPC", 3, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute", 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Create Corner Feature Between First and Second Bounding Surfaces
    Set oMemDesc = pMDs.AddMember("CornerFeature", 4, "CM_ConstructCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsCornerFeatureNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeCF", 1, "IJSmartOccurrence", "CMNeedtoComputeCF", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

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
' Method      : CM_IsPCNeeded
' Description :
'
'*********************************************************************************************

Public Sub CM_IsPCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_DefinitionProgid & "::CM_IsPCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sError As String
    bIsNeeded = False
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim eEndCutType As eEndCutTypes
    
    'Get bounding and bounded ports and also objects
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort, , , eEndCutType
    
    If oBoundedPort Is Nothing Then
        Exit Sub
    End If
    
    Dim oMappedPortsColl As Collection
    Dim oIntersectedTopInsidePort As IJPort
    Dim oIntersectedBtmInsidePort As IJPort
    Dim oTopInsidePos As IJDPosition
    Dim oBtmInsidePos As IJDPosition
    Dim PositionColl As Collection
    Dim oFeature As IJStructFeature
    Dim lDispId As Long
        
    Set oFeature = pMemberDescription.CAO
    lDispId = pMemberDescription.dispid
    
    If lDispId = 2 Then
        bIsNeeded = True
        Exit Sub
    ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
        bIsNeeded = True
        Exit Sub
    End If
    
    'Below checks are needed if the Flange Cut is covering the bounding Port then Web
    'should not create PC as Flange willl be creating it.

    GetMultiBoundingEdgeMap oFeature, eEndCutType, oMappedPortsColl, , PositionColl, _
                                oIntersectedTopInsidePort, oTopInsidePos, oIntersectedBtmInsidePort, oBtmInsidePos
    Dim eFeatureType As StructFeatureTypes
    eFeatureType = 0
    eFeatureType = oFeature.get_StructFeatureType
    
    If lDispId <= oMappedPortsColl.Count Then
        'For Disid 1 and 3 we have to consider the Flange Thickness as it Flange Cut can remove the valid PC
        If lDispId = 1 And (eFeatureType = SF_WebCut) Then
            If Not oIntersectedTopInsidePort Is Nothing Then
                If Not (oMappedPortsColl.Item(1) Is oIntersectedTopInsidePort) Then
                    bIsNeeded = False
                Else
                    Dim oRefPos1 As New DPosition
                    Set oRefPos1 = PositionColl.Item(2)
                    If GreaterThanOrEqualTo(oTopInsidePos.DistPt(oRefPos1), GetMinPCLength) Then
                        bIsNeeded = True
                    Else
                        bIsNeeded = False
                    End If
                End If
            Else
                bIsNeeded = True
            End If
        ElseIf lDispId = 3 And (eFeatureType = SF_WebCut) Then
            If Not oIntersectedBtmInsidePort Is Nothing Then
                If Not (oMappedPortsColl.Item(oMappedPortsColl.Count) Is oIntersectedBtmInsidePort) Then
                    bIsNeeded = False
                Else
                    Dim oRefPos2 As New DPosition
                    Set oRefPos2 = PositionColl.Item(PositionColl.Count - 1)
                    If GreaterThanOrEqualTo(oBtmInsidePos.DistPt(oRefPos2), GetMinPCLength) Then
                        bIsNeeded = True
                    Else
                        bIsNeeded = False
                    End If
                End If
            Else
                bIsNeeded = True
            End If
        Else
            bIsNeeded = True
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
End Sub

'*********************************************************************************************
' Method      : CM_ConstructPC
' Description :
'
'*********************************************************************************************

Public Function CM_ConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, ByRef pObject As Object)

    Const METHOD = m_DefinitionProgid & "::CM_ConstructPC"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lDispId As Long
    Dim lBoundingEdgeID As Long
    Dim lBoundedEdgeID As Long
    lDispId = pMemberDescription.DispId
    
    GetBoundingAndBoundedEdgeIds lDispId, lBoundingEdgeID, lBoundedEdgeID
   
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(lDispId)
               
    Dim oEndCutDefCM As New EndCutDefCM
    Set pObject = oEndCutDefCM.Construct_PhysConn(pMemberDescription, pResourceManager, "TeeWeld", lBoundingEdgeID, lBoundedEdgeID)
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

'*********************************************************************************************
' Method      : CM_IsCornerFeatureNeeded
' Description :
'
'*********************************************************************************************

Public Sub CM_IsCornerFeatureNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsCornerFeatureNeeded"
    On Error GoTo ErrorHandler
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded pMemberDescription.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sError As String
    
    'Out of all the corners list whichever forms first non reflexive angle
    'we will be creating CF there
    bIsNeeded = True

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructCornerFeature
' Description :
'
'*********************************************************************************************
Public Function CM_ConstructCornerFeature(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructCornerFeature"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim eEndCutType As eEndCutTypes
    
    'Get bounding and bounded ports and also objects
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort, , , eEndCutType
    
    Dim oMappedPortsColl As Collection
    Dim oAnglesColl As Collection
    Dim lEdgeID1 As Long
    Dim lEdgeID2 As Long
    
    GetMultiBoundingEdgeMap pMemberDescription.CAO, eEndCutType, oMappedPortsColl, oAnglesColl
                               
    'For Three Surfaces Bounding there is possibility of 2 Corners
    'Between 1st and 2nd and other one is between 2nd and 3rd.
    'Out of these two whichever corner forms angle < 180 deg
    'we take those two ports as inputs for CF
    
    If LessThan(oAnglesColl.Item(1), GetPI) Then
        lEdgeID1 = e_JXSEC_MultipleBounding_5001
        lEdgeID2 = e_JXSEC_MultipleBounding_5002
    Else
        lEdgeID1 = e_JXSEC_MultipleBounding_5003
        lEdgeID2 = e_JXSEC_MultipleBounding_5002
    End If
    
    sMsg = "Creating Corner Feature " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
                          
    Dim oEndCutDefCM As New EndCutDefCM
    Set pObject = oEndCutDefCM.Construct_CornerFeature(pMemberDescription, pResourceManager, "MbrEndCutCorner", lEdgeID1, lEdgeID2)
    
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
        
    'Call all PC Item members conditional and check whether it is True or False
    'Based on this Set FlangeCut Cutting Behaviour
    Dim iCount As Integer
    
    For iCount = 1 To 3
        Set oMemberDescription = oMemberDescriptions.ItemByDispid(1)
        oMemberDescription.CAO = pPropertyDescriptions.CAO
        
        CM_IsPCNeeded oMemberDescription, bIsNeeded
        
        If bIsNeeded Then Exit For
    Next
    
    'Set Proper Flange Cutting Behaviour
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
    
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
Private Sub GetBoundingAndBoundedEdgeIds(lDispId As Long, lBoundingEdgeID As Long, _
                                                                lBoundedEdgeID As Long)
    Const METHOD = "GetBoundingAndBoundedEdgeIds"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Getting the Bounded and bounding Edge IDs based on DispId"
    
    If lDispId = 0 Then
        GoTo ErrorHandler
    ElseIf lDispId = 1 Then
        lBoundingEdgeID = e_JXSEC_MultipleBounding_5001
        lBoundedEdgeID = lBoundingEdgeID
    ElseIf lDispId = 2 Then
        lBoundingEdgeID = e_JXSEC_MultipleBounding_5002
        lBoundedEdgeID = lBoundingEdgeID
    ElseIf lDispId = 3 Then
        lBoundingEdgeID = e_JXSEC_MultipleBounding_5003
        lBoundedEdgeID = lBoundingEdgeID
    Else
        Exit Sub
    End If
      
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'******************************************************************************************
' Method:
' CMNeedtoComputePC
'
' Description: Replaces the old ports with new ones
' *******************************************************************************************
Public Sub CMNeedtoComputePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
                                            
    Const METHOD = "CMNeedtoComputePC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the PC "
            
    If TypeOf pObject Is IJStructPhysicalConnection Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        Dim oMemObjs As IJDMemberObjects
        Dim oMemDescs As IJDMemberDescriptions
        Dim oMemDesc As IJDMemberDescription
                        
        If Not oSmartItem Is Nothing Then
            ' Getting the Member objects of endcut
            Set oMemObjs = pPropertyDescriptions.CAO
            Dim lDispId As Long
            Dim Reqdispid As Long
            Dim index As Long
            Dim MemDescName As String
            Dim oTempObj As Object
            ' Getting the member descriptions from member objects
            Set oMemDescs = oMemObjs.MemberDescriptions
            ' loop through each of member descriptions
            For index = 1 To oMemDescs.Count
                Set oMemDesc = oMemDescs.Item(index)
                lDispId = oMemDesc.DispId
                Set oMemDesc = oMemDescs.ItemByDispid(lDispId)
                If Not oMemDesc.object Is Nothing Then
                    Set oTempObj = oMemDesc.object
                Else
                    GoTo ErrorHandler
                End If
                If TypeOf oTempObj Is IJStructPhysicalConnection Then
                    If oTempObj Is pObject Then ' if the member description object and the pobject matches, that means the evaluation is currently on this object
                        Reqdispid = lDispId ' this is the required dispid
                        Exit For
                    End If
                End If
            Next
            Dim lBoundingEdgeID As eMultiBoundingEdgeIDs
            Dim lBoundedEdgeID As eMultiBoundingEdgeIDs
            GetBoundingAndBoundedEdgeIds Reqdispid, lBoundingEdgeID, lBoundedEdgeID ' get the bounding and bounded edge IDs based on the dispid
            
            Dim oEndCutDefCM As New EndCutDefCM
            oEndCutDefCM.CMNeedtoComputePhysicalConn pPropertyDescriptions, pObject, _
                                                            lBoundingEdgeID, lBoundedEdgeID
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'******************************************************************************************
' Method:
' CMNeedtoComputeCF
'
' Description: Replaces the old ports with new ones
' *******************************************************************************************

Public Sub CMNeedtoComputeCF(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
                                            
    Const METHOD = "CMNeedtoComputeCF"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the CF "
            
    If pObject Is Nothing Then Exit Sub

    'Update Corner Feature input ports
    Dim oEndCutDefCM As New EndCutDefCM
    oEndCutDefCM.CMNeedtoComputeCF_MB pPropertyDescriptions, pObject
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
