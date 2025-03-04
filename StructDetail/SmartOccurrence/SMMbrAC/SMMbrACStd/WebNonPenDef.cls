VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebNonPenDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  File        : WebNonPenDef.cls
'
'  Description :
'
'  Author      : Alligators
'
' 12/May/15   GH    CR-260982 Added Property description to the Physical connection Item members to handle smart item change.
'*********************************************************************************************
Const m_sClassName As String = "WebNonPenDef"
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
    sDispId = "WebPC"
    Set oMemDesc = pMDs.AddMember("WebPC", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigratePhysConn"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "Compute_" & sDispId, 1, "IJSmartOccurrence", "CMNeedtoComputePC", imsCOOKIE_ID_USS_LIB
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
    
    strPCAttribute = oCM.GetWebNonPenPCAttribute(oMD, False)
    
    ' -----------------------
    ' Get the attribute value
    ' -----------------------
    Dim oACDef As New ACDef
    Dim strPCSelection As String
    strPCSelection = oACDef.GetStandardACAttribute(oAppConnection, strPCAttribute)
    
    ' ------------------------------------
    ' If not null, then the item is needed
    ' ------------------------------------
    If strPCSelection <> vbNullString Then
     'slab& wall
        GetBoundingAndBounded oMD.CAO, oBoundngPort, oBoundedPort
        
        If (TypeOf oBoundngPort.Connectable Is SPSSlabEntity) Or (TypeOf oBoundngPort.Connectable Is SPSWallPart) Then
                bIsNeeded = False
        Else
          bIsNeeded = True
        End If
        
    Else
        bIsNeeded = False
    End If
    
    ' -------------------------------------------------------------------
    ' If the Web is completely outside the bounding member, no PC is needed
    ' -------------------------------------------------------------------
    If bIsNeeded Then
        Dim oTopORWL As ConnectedEdgeInfo
        Dim oBottomOrWR As ConnectedEdgeInfo
        Dim oInsideTopFlgOrFL As ConnectedEdgeInfo
        Dim oInsideBtmFlgOrFR As ConnectedEdgeInfo
        
        GetConnectedEdgeInfo oMD.CAO, oBoundedPort, oBoundngPort, oTopORWL, oBottomOrWR, oInsideTopFlgOrFL, oInsideBtmFlgOrFR
        If (oTopORWL.IntersectingEdge = Above And oBottomOrWR.IntersectingEdge = Above) Or _
            (oInsideTopFlgOrFL.IntersectingEdge = Below And oBottomOrWR.IntersectingEdge = Below) Then
            bIsNeeded = False
        End If
    End If
    
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
    Set oObject = oCM.CMConstructNonPenPC(oMD, oResourceManager, True, False)

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub CMNeedtoComputePC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)

    Const METHOD = "::CMNeedtoComputePC"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Modifying the PC "
    
    If pObject Is Nothing Then Exit Sub
    
    Dim oSmartOccurance As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    
    Set oSmartOccurance = pObject
    Set oSmartItem = oSmartOccurance.ItemObject
    
    'Exit while creating the PC
    If oSmartItem Is Nothing Then
        Exit Sub
    End If
    
    'Check object Type
    If TypeOf pObject Is IJStructPhysicalConnection Then
    
        Dim oMemberDesc As IJDMemberDescription
        Set oMemberDesc = GetMemberDescriptionFromPropertyDescs(pPropertyDescriptions, pObject)
        
        If Not oMemberDesc Is Nothing Then
            
            Dim eBoundingSubPort As JXSEC_CODE
            Dim eBoundedEdgeSubPort As JXSEC_CODE
            
            Dim strACAttribute As String
            Dim strPCSelection As String
            Dim oCM As New CustomMethods
            Dim strNewFilter As String
                        
            strACAttribute = oCM.GetWebNonPenPCAttribute(oMemberDesc, False)
    
            'Get PC Filter
            oCM.GetPCSelection oMemberDesc, strACAttribute, strPCSelection, strNewFilter
            
            'Update PC
            UpdatePCWithNewFilterProgID pObject, strNewFilter
        End If
    End If
    
 Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

