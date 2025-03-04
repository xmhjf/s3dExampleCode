VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateEdgeToPlateEdgeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.PlateEdgeToPlateEdgeDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.PlateEdgeToPlateEdgeDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\PlateEdgeToPlateEdgeDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs", strError).Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim pMemDesc As IJDMemberDescription

    strError = "Constructing physical connection."
    Set pMemDesc = pMDs.AddMember("PlateEdgeToPlateEdge1", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsPhysConn"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMChamferPhysConn"
    Set pMemDesc = Nothing
   
    'Create chamfer on object 1
    strError = "Constructing chamfer."
    Set pMemDesc = pMDs.AddMember("PlateEdgeToPlateEdgeChamfer1", 2, "CMConstructChamfer", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsChamfer"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateChamfer1"
    Set pMemDesc = Nothing
        
    'Create chamfer on object 2
    strError = "Constructing chamfer."
    Set pMemDesc = pMDs.AddMember("PlateEdgeToPlateEdgeChamfer2", 3, "CMConstructChamfer2", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetChamferNumber"
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMCreateChamfer2"
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMSetChamferNumber(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Dim strError As String
          
    CMCopyQuestionsChamfer pMemberDescription
    
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set weld number"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "ChamferWeld", "Second"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetChamferNumber", strError).Number
End Sub
Public Sub CMCopyQuestionsChamfer(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Calling Copy answer to copy assembly questions to collar"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "ChamferType", _
                                 CUSTOMERID + "AssyConRul.PlateEdgeByPlateEdgeSel", "ChamferType"
    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "Shipyard", _
                                 CUSTOMERID + "AssyConRul.PlateEdgeByPlateEdgeSel", "Shipyard"
    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "WeldingType", _
                                 CUSTOMERID + "AssyConRul.PlateEdgeByPlateEdgeSel", "WeldingType"
                                
    Set pCopyAnswerHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMemebers", strError).Number
End Sub

Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim sSmartClass As String
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim clsAssyConnCM As New AssyConnDefCM
    
    Set oAssyConn = New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pMemberDescription.CAO
    
    ' Check for Special case of Plate Edge to Plate Edge
    ' Builtup Members have connection cases where Plate Edge to Plate Edge
    ' are "Tee" type connections (not "Butt" connections)
    ' Therefore, need to create required Physical Connection
    sSmartClass = "ButtWeld"
    If oAssyConn.IsEdgeToEdgeTeeConnection Then
        sSmartClass = "TeeWeld"
    End If
    

    strError = "Constructing physical connection."
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnFromACPorts(pMemberDescription, pResourceManager, sSmartClass)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC", strError).Number
End Sub
Public Sub CMConstructChamfer(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Retrieving CAO and constructing chamfer."
    Set pObject = clsAssyConnCM.CMConstruct_Chamfer(pMemberDescription, pResourceManager, "RootChamfer")
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructChamfer", strError).Number
End Sub

Public Sub CMConstructChamfer2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Retrieving CAO and constructing chamfer."
    Set pObject = clsAssyConnCM.CMConstruct_Chamfer2(pMemberDescription, pResourceManager, "RootChamfer")
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructChamfer2", strError).Number
End Sub

Public Sub CMCopyQuestionsPhysConn(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    Dim oAssyConn As New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pMemberDescription.CAO
    
    strError = "Calling Copy answer to copy assembly questions to collar"
    
    'Check for Special case of Plate Edge to Plate Edge
    'Builtup Members have connection cases where Plate Edge to Plate Edge
    'are "Tee" type connections (not "Butt" connections)
    '     if it is Tee Weld we need not copy any answer to Phys Rules
    If Not oAssyConn.IsEdgeToEdgeTeeConnection Then
        Dim pCopyAnswerHelper As New CopyAnswerHelper
        Set pCopyAnswerHelper.MemberDescription = pMemberDescription
        pCopyAnswerHelper.CopyAnswer CUSTOMERID + "PhysConnRules.ButtWeldSel", "Shipyard_FromAssyConn", _
                                     CUSTOMERID + "AssyConRul.PlateEdgeByPlateEdgeSel", "Shipyard"
        pCopyAnswerHelper.CopyAnswer CUSTOMERID + "PhysConnRules.ButtWeldSel", "WeldingType_FromAssyConn", _
                                     CUSTOMERID + "AssyConRul.PlateEdgeByPlateEdgeSel", "WeldingType"
                                
        Set pCopyAnswerHelper = Nothing
    End If
    
    Set oAssyConn = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMemebers", strError).Number
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
    pMDs.RemoveAll ' Remove all the previous Member descriptions
    ItemMembers pMDs
  
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
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
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

