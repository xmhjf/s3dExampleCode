VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrGenericWebDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'File : MbrGenericWebDef.cls
'
'Author : D.A. Trent
'
'Description :
'   Control creation of the Member Items for the MbrGenericWebSel Selector Class
'
'
'History:
'
'********************************************************************

Option Explicit

Const m_FamilyProgid As String = ""
Const m_DefinitionName As String = "Mbr_EndCuts.MbrGenericWebDef"
Const m_DefinitionProgid As String = "Mbr_EndCuts.MbrGenericWebDef"
Const MODULE = "\StructDetail\Data\SmartOccurrence\Mbr_EndCuts\MbrGenericWebDef.cls"

Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

Implements IJDUserSymbolServices
'

' List any graphic Inputs that the Definition has here
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

' List any graphic Inputs that the Definition has here
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    pAD.UserTypeClsid = CA_WEBCUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CM_FinalConstructEndCut"
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateAggregator"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' List all the Item members that this SmartClass can create
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sDispId As String
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    ' Generic WebCuts can consist of (2) Physical Connections
    '                                (1) Flange Cut
    ' Index(1) - Idealize Boundary Port (8193)
    ' Index(2) - Base/Offset Port when bounding is Plate Lateral Port (8192)
    ' Index(3) - Flange Cut based on Web Cut
    For iDispId = 1 To 3
        sDispId = "GenericWeb_" & Trim(Str(iDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPhysConns", imsCOOKIE_ID_USS_LIB)
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
        oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
        oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
        Set oMemDesc = Nothing
    Next iDispId

    
    Set oMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
Public Sub CM_PhysConnsCount(pMemberDescription As IJDMemberDescription, nCount As Long)
Const METHOD = m_DefinitionProgid & "::CM_PhysConnsCount"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lStatus As Long
    nCount = 1
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CM_IsPhysConnsNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = m_DefinitionProgid & "::CM_IsPhysConnsNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim sSelection As String
    Dim sEndCutType As String
    Dim sPortFaceType As String
    
    Dim lDispId As Long
    
    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object

    Dim oPort As IJPort
    Dim oConnectable As IJConnectable
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper

    ' Get the Current Selection
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oEndCutObject = pMemberDescription.CAO
    
    sSelection = ""
    If TypeOf oEndCutObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oEndCutObject
        Set oSmartItem = oSmartOccurrence.SmartItemObject
        sSelection = oSmartItem.Name
    End If
    
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    Set oBoundedObject = oSDO_WebCut.BoundedPort
    Set oBoundingObject = oSDO_WebCut.BoundingPort
    
    ' Generic WebCuts can consist of (2) Physical Connections
    '                                (1) Flange Cut
    ' Index(1) - Idealize Boundary Port (8193)
    ' Index(2) - Base/Offset Port when bounding is Plate Lateral Port (8192)
    ' Index(3) - Flange Cut based on Web Cut
    If Not TypeOf oBoundingObject Is IJPort Then
        ' Bounding Object is NOT IJPort: No Physical Connections allowed
        
    ElseIf lDispId = 1 Then
        If sSelection = "M_Web_Generic_01" Then
            bIsNeeded = True
        ElseIf sSelection = "M_Web_Generic_02" Then
            bIsNeeded = True
        ElseIf sSelection = "M_Web_Generic_03" Then
            bIsNeeded = True
        ElseIf sSelection = "M_Web_Generic_Tube01" Then
            bIsNeeded = True
        End If
    
    ElseIf lDispId = 2 Then
        If sSelection = "M_Web_Generic_03" Then
            Set oPort = oBoundingObject
            Set oConnectable = oPort.Connectable
            If TypeOf oConnectable Is IJPlatePart Then
                sPortFaceType = Get_PortFaceType(oBoundingObject)
                If sPortFaceType = C_Port_Lateral Then
                    bIsNeeded = True
                End If
            
            ElseIf TypeOf oConnectable Is IJProfilePart Then
                sPortFaceType = Get_PortFaceType(oBoundingObject)
                If sPortFaceType = C_Port_Lateral Then
                    bIsNeeded = True
                End If
            
            ElseIf TypeOf oConnectable Is ISPSMemberPartPrismatic Then
                sPortFaceType = Get_PortFaceType(oBoundingObject)
                If sPortFaceType = C_Port_Lateral Then
                    bIsNeeded = True
                End If
            End If
        End If
    
    ElseIf lDispId = 3 Then
        ' Currently Do Not Suuport placing Flange Cut
    Else
        '
        bIsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CM_CopyQuestionsToMember(pMemberDescription As IJDMemberDescription)
Const METHOD = m_DefinitionProgid & "::CM_CopyQuestionsToMember"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

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

Public Sub CM_FinalConstructEndCut(ByVal pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = m_DefinitionProgid & "::CM_FinalConstructEndCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    EndCut_FinalConstruct pAggregatorDescription
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_MigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, _
                                pMigrateHelper As IJMigrateHelper)
Const METHOD = m_DefinitionProgid & "::CM_MigrateAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Mirgrating EndCut Object"
    MigrateEndCutObject pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_MigratePhysConn(pMemberDesc As IJDMemberDescription, _
                              pMigrateHelper As IJMigrateHelper)
Const METHOD = m_DefinitionProgid & "::CM_MigratePhysConn"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' The Pyhsical Connection Inputs have been Mirgrated at same time with the EndCut Inputs
    ' Therefore, there should be nothing to do here
    sMsg = "Mirgrating Pyhsical Connection Inputs"

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
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

' ********************************************************************************************
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
    pIH.Definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    ItemMembers pMDs
End Sub

' ********************************************************************************************
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

' ********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
End Sub

' ********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
' ********************************************************************************************
' ********************************************************************************************
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
    
    ' Generic WebCuts can consist of (2) Physical Connections
    '                                (1) Flange Cut
    ' Index(1) - Idealize Boundary Port (8193)
    ' Index(2) - Base/Offset Port when bounding is Plate Lateral Port (8192)
    ' Index(3) - Flange Cut based on Web Cut
    lDispId = pMemberDescription.dispid
    If lDispId = 1 Then
        lEdgeId = 8193
        eBoundingPort = CTX_NOP
        eBoundedSubPort = lEdgeId
        
    ElseIf lDispId = 2 Then
        lEdgeId = 8192
        eBoundingPort = CTX_VIRTUAL
        eBoundedSubPort = lEdgeId
    
    ElseIf lDispId = 3 Then
        ' Currently Do Not Suuport placing Flange Cut
        sMsg = "Creating Flange Cut (Not Suuported)" & _
               "...pMemberDescription.dispid = " & Str(lDispId)
        GoTo ErrorHandler
    
    Else
    End If

    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(lDispId)
    
    If pResourceManager Is Nothing Then
        sMsg = "Modifying Physical Connection " & _
               "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
        Modify_PhysConn pMemberDescription, eBoundingPort, eBoundedSubPort
        
    Else
        sMsg = "Creating Physical Connection " & _
               "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
        Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "TeeWeld", eBoundingPort, eBoundedSubPort)
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function




