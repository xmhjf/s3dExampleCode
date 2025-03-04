VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenMB_PCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenMB_PCDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    23/MAY/2014 - Created
'    22/April/2015 - modified MDT TR-271041.Added properties to update the PCs when bounded member is rotated after AC placement
'*********************************************************************************************
Const m_sClassName As String = "GenMB_PCDef"
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
' Description : List all the Item members that this SmartClass can create
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
    
    'Physical Connections With Fourth Bounding Surface
    Set oMemDesc = pMDs.AddMember("ForthPC", 4, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute", 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Physical Connections With Fifth Bounding Surface
    Set oMemDesc = pMDs.AddMember("FifthPC", 5, "CM_ConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute", 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
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
    
    'Below checks are needed if the Flange Cut is covering the bounding Port then Web
    'should not create PC as Flange willl be creating it.
    
    GetMultiBoundingEdgeMap pMemberDescription.CAO, eEndCutType, oMappedPortsColl, , PositionColl, _
                                oIntersectedTopInsidePort, oTopInsidePos, oIntersectedBtmInsidePort, oBtmInsidePos
    Dim eFeatureType As StructFeatureTypes
    eFeatureType = 0
    eFeatureType = oFeature.get_StructFeatureType
    
    If lDispId <= oMappedPortsColl.Count Then
        'For Disid 1 and 5 we have to consider the Flange Thickness as it Flange Cut can remove the valid PC
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
        ElseIf lDispId = oMappedPortsColl.Count And (eFeatureType = SF_WebCut) Then
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
    
    For iCount = 1 To 5
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
    
    Select Case lDispId
        Case 0:
            GoTo ErrorHandler
        Case 1:
            lBoundingEdgeID = e_JXSEC_MultipleBounding_5001
            
        Case 2:
            lBoundingEdgeID = e_JXSEC_MultipleBounding_5002
            
        Case 3:
            lBoundingEdgeID = e_JXSEC_MultipleBounding_5003
            
        Case 4:
            lBoundingEdgeID = e_JXSEC_MultipleBounding_5004
            
        Case 5:
            lBoundingEdgeID = e_JXSEC_MultipleBounding_5005
        
        Case Else:
            Exit Sub
        
    End Select
    
    lBoundedEdgeID = lBoundingEdgeID
      
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
                    If oTempObj Is pObject Then  ' if the member description object and the pobject matches, that means the evaluation is currently on this object
                        Reqdispid = lDispId ' this is the required dispid
                        Exit For
                    End If
                End If
            Next
            Dim lBoundingEdgeID As eMultiBoundingEdgeIDs
            Dim lBoundedEdgeID As eMultiBoundingEdgeIDs
            GetBoundingAndBoundedEdgeIds Reqdispid, lBoundingEdgeID, lBoundedEdgeID ' get the bounding and bounded edge IDs based on the dispid
            Dim oEndCutDefCM As New EndCutDefCM
            oEndCutDefCM.CMNeedtoComputePhysicalConn pPropertyDescriptions, pObject, lBoundingEdgeID, lBoundedEdgeID
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub







