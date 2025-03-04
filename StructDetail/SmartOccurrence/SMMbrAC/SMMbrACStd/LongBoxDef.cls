VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongBoxDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrAxisWebDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'
'*********************************************************************************************
Const m_sClassName As String = "LongBoxDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
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
    
    pMDs.RemoveAll  ' Removes all the previous cached member descriptions(best practice)
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    iDispId = 1
    sDispId = "AxisWeb_PC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPhysConns", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
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
    
    bIsNeeded = False 'Initial value, make it True only for valid cases
    
    Dim oEndCutObject As Object
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim sSelection As String

    sSelection = ""
    Set oEndCutObject = pMemberDescription.CAO

    If TypeOf oEndCutObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oEndCutObject
        Set oSmartItem = oSmartOccurrence.SmartItemObject
        sSelection = oSmartItem.Name
    End If
    
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    Dim oSDO_WeCut As New StructDetailObjects.WebCut
    
    Set oSDO_WeCut.object = oEndCutObject
    Set oBoundedPort = oSDO_WeCut.BoundedPort
    Set oBoundingPort = oSDO_WeCut.BoundingPort
    
    If oSDO_WeCut.IsFreeEndCut Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
'    Dim bPenetratesWeb As Boolean
'
'    Dim eWL As ConnectedEdgeInfo
'    Dim eWR As ConnectedEdgeInfo
'    Dim eFL As ConnectedEdgeInfo
'    Dim eFR As ConnectedEdgeInfo
'
'    Dim oParentObj As Object
'    Dim sParentItemName As String
'    Parent_SmartItemName oEndCutObject, sParentItemName, oParentObj
'
'    Dim oAppConnection As IJAppConnection
'    Set oAppConnection = oParentObj
'
'    If sSelection = "Axis_Web_Straight" Or IsTubularMember(oSDO_WeCut.Bounded) = False Or sSelection = "Axis_Web_Cope_Ext" Then
'        If (TypeOf oSDO_WeCut.Bounded Is ISPSMemberPartCommon Or TypeOf oSDO_WeCut.Bounded Is IJProfile) And _
'           TypeOf oSDO_WeCut.Bounding Is ISPSMemberPartCommon Then
'
'           'If Tube is bounding member, determine if PC is needed and exit
'           If IsTubularMember(oSDO_WeCut.Bounding) Then
'                Dim dDistanceFromBounding1 As Double
'                Dim dDistanceFromBounding2 As Double
'                'Check distance from web-left/web-right value
'                dDistanceFromBounding1 = GetDistanceFromBounding(oSDO_WeCut.Bounding, oSDO_WeCut.Bounded, _
'                                            JXSEC_WEB_LEFT)
'                dDistanceFromBounding2 = GetDistanceFromBounding(oSDO_WeCut.Bounding, oSDO_WeCut.Bounded, _
'                                            JXSEC_WEB_RIGHT)
'                If dDistanceFromBounding1 < 0.0001 Or _
'                    dDistanceFromBounding2 < 0.0001 Then 'less than 0.1mm
'                    bIsNeeded = True
'                End If
'                Exit Sub '*** Exit ***
'           End If
'
'           bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
'
'           If Not bPenetratesWeb Then
'               GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eWL, eWR, eFL, eFR
'               If (eWL.IntersectingEdge = Above And eWR.IntersectingEdge = Above) Or _
'                  (eWL.IntersectingEdge = Below And eWR.IntersectingEdge = Below) Or _
'                  (eWL.CoplanarEdge = Bottom) Or (eWR.CoplanarEdge = Top) Then
'                     bIsNeeded = False
'                Else
'                    bIsNeeded = True
'                End If
'           Else
'               bIsNeeded = True
'           End If
'        Else
'          bIsNeeded = False
'        End If
'    ElseIf InStr(1, sSelection, "Tube", vbTextCompare) > 0 Then
'        bIsNeeded = True
'    End If
    
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
   
    ' Get Bounding Part's Port
    Dim eBoundingPort As eUSER_CTX_FLAGS
    Dim eBoundedSubPort As JXSEC_CODE
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    
    'TR 206472 - Change bounding port to use Base port of the bounding plate part when the Bounding is a built-up
    
    Dim oBUMember As ISPSDesignedMember
    Dim bBuiltup As Boolean
    
    IsFromBuiltUpMember oSDO_WebCut.Bounding, bBuiltup, oBUMember
    If bBuiltup Then
        eBoundingPort = CTX_BASE
    Else
        eBoundingPort = CTX_LATERAL
    End If
    
    eBoundedSubPort = 8193
    
    If pResourceManager Is Nothing Then
        sMsg = "Modifying Physical Connection " & _
               "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
        Set pObject = Modify_PhysConn(pMemberDescription, eBoundingPort, eBoundedSubPort)
        
    Else
        sMsg = "Creating Physical Connection " & _
               "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)

'         ----------------------------------------------------------------------
'         If the bounding object is a member, find out which edge is intersected
'         --------------------------------------------------------------------
        Dim boundingEdgeId As eBounding_Edge
        Dim mappedEdgeId As JXSEC_CODE
    
        boundingEdgeId = None
       If TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic And Not IsTubularMember(oSDO_WebCut.Bounding) Then
            GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, False
       End If
       
        ' -----------------------------------------------------------
        ' Create a Tee or Butt weld, depending on bounding edge width
        ' -----------------------------------------------------------
        
        If boundingEdgeId = eBounding_Edge.Bottom_Flange_Right Or boundingEdgeId = eBounding_Edge.Top_Flange_Right Then
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "ButtWeld", eBoundingPort, eBoundedSubPort)
        ElseIf Not (boundingEdgeId = eBounding_Edge.Above Or boundingEdgeId = eBounding_Edge.Below Or boundingEdgeId = eBounding_Edge.None) Then
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "TeeWeld", eBoundingPort, eBoundedSubPort)
        Else
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "Teeweld", eBoundingPort, eBoundedSubPort)
        End If
          
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


