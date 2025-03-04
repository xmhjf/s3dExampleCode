VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeNonPenDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014-15, Intergraph Corporation.  All rights reserved.
'
'  File        : FlangeNonPenDef.cls
'
'  Description :
'
'  Author      : Alligators

' 03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
' 27/Feb/15  MDT/GH CR-265236 CMNeedtoComputeFlangePC() method is added to update the flange PCs for box connections when flip question answer is changed.
' 12/May/15   GH    CR-260982 Updated CMNeedtoComputeFlangePC() to handle smart item changes
'*********************************************************************************************
Const m_sClassName As String = "FlangeNonPenDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

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
Public Sub ItemAggregator(oAD As IJDAggregatorDescription)
    
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    oAD.UserTypeClsid = CA_FLANGECUT       '
    oAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    oAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    oAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
    Dim oPDs As IJDPropertyDescriptions
    Set oPDs = oAD
    oPDs.RemoveAll
    
    ' Set the FlangeCutting Behaviour
    sMsg = "Setting Flange Cutting Behaviour property."
    oPDs.AddProperty "CuttingBehavior", 1, "IJUAMbrEndCuts", "SetFlangeCuttingBehaviour", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
        
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
    Set oMemDesc = pMDs.AddMember("FlangePC", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedtoComputeFlangePC", 1, "IJSmartOccurrence", "CMNeedtoComputeFlangePC", imsCOOKIE_ID_USS_LIB
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
 
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
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

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
'*********************************************************************************************
' Method      : CM_IsPhysConnsNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsPhysConnsNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsPhysConnsNeeded"
    
    On Error GoTo ErrorHandler
     Dim oBoundedPort As IJPort
    
    Dim oBoundngPort As IJPort
    Dim sMsg As String
    sMsg = m_DefinitionProgid & "::CM_IsPhysConnsNeeded"
        
    ' ---------------------------
    ' Get the assembly connection
    ' ---------------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName oMD.CAO, , oAppConnection
    
    ' -----------------------------
    ' Get the name of the attribute
    ' -----------------------------
    Dim strPCAttribute As String
    Dim oCM As New CustomMethods
    Dim bIsMiterOrBoxCut As Boolean
    Dim oACDef As New ACDef
    
    bIsMiterOrBoxCut = False
    
    bIsMiterOrBoxCut = oACDef.IsBoxCut(oAppConnection)
    strPCAttribute = oCM.GetFlangeNonPenPCAttribute(oMD, bIsMiterOrBoxCut)
    
    ' -----------------------
    ' Get the attribute value
    ' -----------------------
    Dim strPCSelection As String
    strPCSelection = oACDef.GetStandardACAttribute(oAppConnection, strPCAttribute)
    
    ' ------------------------------------
    ' If not null, then the item is needed
    ' ------------------------------------
    If strPCSelection <> vbNullString Then
     GetBoundingAndBounded oMD.CAO, oBoundngPort, oBoundedPort
        
        If (TypeOf oBoundngPort.Connectable Is SPSSlabEntity) Or (TypeOf oBoundngPort.Connectable Is SPSWallPart) Then
        
                bIsNeeded = False
        Else
                bIsNeeded = True
        End If
    Else
           bIsNeeded = False
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'******************************************************************************************
' Method:
' SetFlangeCuttingBehaviour
' *******************************************************************************************
Public Sub SetFlangeCuttingBehaviour(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    
    Const METHOD = "SetFlangeCuttingBehaviour"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Updating a Flange Cutting Behaviour"
            
    ' ----------------------------------------------
    ' Exit if the behavior attribute does not exists
    ' ----------------------------------------------
    Dim oMemberDescriptions As IJDMemberDescriptions
    Set oMemberDescriptions = pPropertyDescriptions.definition
    
    Dim IsProExists As Boolean
    IsProExists = Has_Attribute(pPropertyDescriptions.CAO, "CuttingBehavior")
    
    If Not IsProExists Then
        Exit Sub
    End If

    ' -------------------------------
    ' Determine if there will be a PC
    ' -------------------------------
    Dim oMemberDescription As IJDMemberDescription
    
    Set oMemberDescription = oMemberDescriptions.ItemByDispid(1)
    oMemberDescription.CAO = pPropertyDescriptions.CAO

    Dim bIsNeeded As Boolean
    bIsNeeded = True

    CM_IsPhysConnsNeeded oMemberDescription, bIsNeeded
    
    ' ----------------------------------------------------------------------------------
    ' Get proper flange cutting behaviour based on existence of PC and flange cut inputs
    ' ----------------------------------------------------------------------------------
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
    
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function CMConstructPC(ByVal oMD As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)

    Const METHOD = m_DefinitionProgid & "::CMConstructPC"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & "...oMD.dispid = " & Str(oMD.dispid)

    Dim oCM As New CustomMethods
    Set oObject = oCM.CMConstructNonPenPC(oMD, oResourceManager, False, False)

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' **********************************************************
' Method : CMNeedtoComputeFlangePC
' This method will replace the physical connection ports with updated ports based on the flipped ports
' **********************************************************
Public Sub CMNeedtoComputeFlangePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "::CMNeedtoComputeFlangePC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the FlangePC"
    
    Dim eBoundingSubPort As JXSEC_CODE
    Dim eBoundedSubPort As JXSEC_CODE
    
    If TypeOf pObject Is IJStructPhysicalConnection Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        
        Dim oAppConnection As IJAppConnection
        Dim bIsBoxCut As Boolean
        Dim oACDef As New ACDef
        AssemblyConnection_SmartItemName pPropertyDescriptions.CAO, , oAppConnection
        
        bIsBoxCut = oACDef.IsBoxCut(oAppConnection)
        
        If Not oSmartItem Is Nothing And bIsBoxCut Then
                     
            Dim oSD_PhysicalConn As New StructDetailObjects.PhysicalConn
            Set oSD_PhysicalConn.object = pObject
            
            Dim oSD_FlangeCut As New StructDetailObjects.FlangeCut
            Set oSD_FlangeCut.object = pPropertyDescriptions.CAO
            
            Dim oSD_Webcut As New StructDetailObjects.WebCut
            Set oSD_Webcut.object = oSD_FlangeCut.WebCut
            
            Dim oBoundedObject As Object
            Set oBoundedObject = oSD_FlangeCut.Bounded
            ' replacing of physical connection ports is not needed when the oboundedObject before and after flip are same.
            If oBoundedObject Is oSD_PhysicalConn.ConnectedObject1 Then
                Exit Sub
            End If
                    
            Dim eBoundingEdge As eBounding_Edge
            Dim eBoundingEdgeID As JXSEC_CODE
            Dim bisBottom As Boolean
            eBoundingEdge = None
            eBoundingEdgeID = -1
            
            If TypeOf oSD_Webcut.Bounding Is ISPSMemberPartPrismatic Then
                If Not IsTubularMember(oSD_Webcut.Bounding) Then
                    GetNonPenetratedIntersectedEdge oSD_Webcut.object, oSD_Webcut.BoundingPort, oSD_Webcut.BoundedPort, _
                                                    eBoundingEdge, eBoundingEdgeID, bisBottom
                End If
            End If
                                
            ' -------------------------------
            ' Get the PC selection and filter
            ' -------------------------------
            Dim oBoundingPort As IJPort
            Dim pHelper As New StructDetailObjects.Helper
    
            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSD_FlangeCut.Bounding, eBoundingEdgeID)
            Set oBoundingPort = pHelper.GetEquivalentLastPort(oBoundingPort)
    
            Dim oStructProfilePart As IJStructProfilePart
            Dim oStructEndCutUtil As IJStructEndCutUtil
            Dim oBoundedPort As IJPort
            
            Set oStructProfilePart = oSD_Webcut.Bounded
            Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
            oStructEndCutUtil.GetLatePortForFeatureSegment oSD_FlangeCut.object, 8193, oBoundedPort
    
            Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
            Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
            
            oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysicalConn.Port1, oBoundedPort
            oConnectionDefinition.ReplacePhysicalConnectionPort pObject, oSD_PhysicalConn.Port2, oBoundingPort
        
        ElseIf Not oSmartItem Is Nothing Then
            
            'Update PC if needed
            Dim oMemberDesc As IJDMemberDescription
            Set oMemberDesc = GetMemberDescriptionFromPropertyDescs(pPropertyDescriptions, pObject)
            
            If Not oMemberDesc Is Nothing Then
                                
                Dim strACAttribute As String
                Dim strPCSelection As String
                Dim oCM As New CustomMethods
                Dim strNewFilter As String
                            
                strACAttribute = oCM.GetFlangeNonPenPCAttribute(oMemberDesc, False)
        
                'Get PC Filter
                oCM.GetPCSelection oMemberDesc, strACAttribute, strPCSelection, strNewFilter
                
                'Update PC
                UpdatePCWithNewFilterProgID pObject, strNewFilter
            End If
        End If
    End If
    
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

