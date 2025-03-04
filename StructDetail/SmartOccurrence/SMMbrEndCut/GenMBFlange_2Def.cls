VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenMBFlange_2Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenMBFlange_2Def.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    29/MAY/2014 - Created
'    22/April/2015 - modified MDT TR-271041.Added properties to update the PCs when bounded member is rotated after AC placement
'*********************************************************************************************
Const m_sClassName As String = "GenMBFlange_2Def"
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
    Set oMemDesc = pMDs.AddMember("Flange_PC", 1, "CM_ConstructPhysConns", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoCompute", 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
        
    Set oMemDesc = pMDs.AddMember("InsideCorner", 2, "CM_ConstructInsideCornerFeature", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsInsideCornerFeatureNeeded"
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
' Method      : CM_IsPhysConnsNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsPhysConnsNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsPhysConnsNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = True

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub
    
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


    ' --------------------------------------------
    ' Find out if this is the top or bottom flange
    ' --------------------------------------------
    Dim strStartClass As String
    Dim sBottomFlange As String
    Dim isBottom As Boolean
    isBottom = False
    
    GetSelectorAnswer pMemberDescription, "BottomFlange", sBottomFlange
    If sBottomFlange = "Yes" Then
        isBottom = True
    End If
    
    Dim oWebCut As New StructDetailObjects.WebCut
    Dim oFeature As IJStructFeature
    Dim oBdedPart As Object
    Dim oBdingPort As IJPort
    Dim oBdingPart As Object
    Dim oFlangeCut As New StructDetailObjects.FlangeCut
    
    Set oFeature = pMemberDescription.CAO
    
    sMsg = "Getting Bounding object"
    
    If oFeature.get_StructFeatureType = SF_FlangeCut Then
        Set oFlangeCut.object = pMemberDescription.CAO
        Set oBdedPart = oFlangeCut.Bounded
        Set oBdingPort = oFlangeCut.BoundingPort
        Set oBdingPart = oFlangeCut.Bounding ' Get Bounding part from wrapper
    End If
    
    Dim oBoundingPort As IJPort
    Dim pSDOHelper As New StructDetailObjects.Helper
    Set oBoundingPort = oBdingPort
    Set oBoundingPort = pSDOHelper.GetEquivalentLastPort(oBoundingPort)
    
    
    ' Get bounded profile Port
        Dim oBoundedPort As IJPort
    If pSDOHelper.ObjectType(oBdedPart) = SDOBJECT_STIFFENER Then
        Dim oBoundedPart As StructDetailObjects.ProfilePart
        Set oBoundedPart = New StructDetailObjects.ProfilePart
        Set oBoundedPart.object = oBdedPart
        Set oBoundedPort = oBoundedPart.CutoutSubPort(oFeature, 8193)
    ElseIf pSDOHelper.ObjectType(oBdedPart) = SDOBJECT_MEMBER Then
        Dim oStructProfilePart As IJStructProfilePart
        Dim oStructEndCutUtil As IJStructEndCutUtil
        Set oStructProfilePart = oBdedPart
        Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
        oStructEndCutUtil.GetLatePortForFeatureSegment oFeature, 8193, oBoundedPort
    End If
    
     ' Get the web cut, make it a parent of the PC
    Dim pSystemParent As IJSystemChild '
    sMsg = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
    strStartClass = "TeeWeld"
    
    ' Create physical connection
    If Not oBoundedPort Is Nothing Then
    
        Dim oPhysicalConnection As New PhysicalConn
        sMsg = "Creating Physical Connection"
        oPhysicalConnection.Create pResourceManager, _
                                   oBoundedPort, _
                                   oBoundingPort, _
                                   strStartClass, _
                                   pSystemParent, ConnectionStandard
                                   
        sMsg = "Setting Physical Connection to private variable"
        Set pObject = oPhysicalConnection.object
        
    End If
    
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
    
    bIsNeeded = True
        
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
        
    'Call PC Item member conditional and check whether it is True or False
    'Based on this Set FlangeCut Cutting Behaviour
    Set oMemberDescription = oMemberDescriptions.ItemByDispid(1)
    oMemberDescription.CAO = pPropertyDescriptions.CAO

    CM_IsPhysConnsNeeded oMemberDescription, bIsNeeded
    
    'Set Proper Flange Cutting Behaviour
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
    
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
                                
        If Not oSmartItem Is Nothing Then
            Dim oSD_PhysicalConn As New StructDetailObjects.PhysicalConn
            Dim oStructFeature As IJStructFeature
            Set oSD_PhysicalConn.object = pObject
            Dim sBottomFlange As String
            Dim oSD_FlangeCut As New StructDetailObjects.FlangeCut
            If TypeOf pPropertyDescriptions.CAO Is IJStructFeature Then
                Dim oBoundingPort As IJPort
                Set oSD_FlangeCut.object = pPropertyDescriptions.CAO
                ' get the bounding port
                Set oBoundingPort = oSD_FlangeCut.BoundingPort
                Dim oBoundedPort As IJPort
                Set oBoundedPort = oSD_FlangeCut.BoundedPort
                Dim pSDOHelper As New StructDetailObjects.Helper
                ' get the bounded port
                If pSDOHelper.ObjectType(oBoundedPort.Connectable) = SDOBJECT_STIFFENER Then
                    Dim oBoundedPart As StructDetailObjects.ProfilePart
                    Set oBoundedPart = New StructDetailObjects.ProfilePart
                    Set oBoundedPart.object = oBoundedPort.Connectable
                    Set oBoundedPort = oBoundedPart.CutoutSubPort(oSD_FlangeCut.object, 8193)
                ElseIf pSDOHelper.ObjectType(oBoundedPort.Connectable) = SDOBJECT_MEMBER Then
                    Dim oStructProfilePart As IJStructProfilePart
                    Dim oStructEndCutUtil As IJStructEndCutUtil
                    Set oStructProfilePart = oBoundedPort.Connectable
                    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
                    oStructEndCutUtil.GetLatePortForFeatureSegment oSD_FlangeCut.object, 8193, oBoundedPort
                End If
                ' checking whether the objects are same or not
                If Not oBoundingPort Is oSD_PhysicalConn.Port2 Then
                    Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
                    Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
                    oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                                 oSD_PhysicalConn.Port2, oBoundingPort
                    oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                                 oSD_PhysicalConn.Port1, oBoundedPort
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


