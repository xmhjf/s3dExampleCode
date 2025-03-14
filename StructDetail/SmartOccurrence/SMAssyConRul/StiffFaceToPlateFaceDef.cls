VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffFaceToPlateFaceDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2013 Intergraph Corporation. All rights reserved.
'
'File : BracketToPlateFaceDef.cls
'
'Author : Alligators
'
'Description :
'
'
'History :
'    5/Nov/2013 - svsmylav
'          DI-CP-240506 'UpdateCFWhenSeamIsDeleted' method is called conditinally (e.g. not-seam movement case).
'
'*****************************************************************************
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.StiffFaceToPlateFaceDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.StiffFaceToPlateFaceDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffFaceToPlateEdgeDef.cls"

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
    
    Set pMemDesc = Nothing
    
    strError = "Constructing physical connection."
    Set pMemDesc = pMDs.AddMember("StiffFaceToPlateFace1", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional CMLIBRARY_ASSYCONNRULES, "CMConstructPCCon"
    
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMemberDescription.CAO

    strError = "Getting Mounting face"
    Dim oProfile As New StructDetailObjects.ProfilePart
    Dim oAC As New StructDetailObjects.AssemblyConn
    Set oAC.object = pMemberDescription.CAO
    Set oProfile.object = oAC.ConnectedObject1
    Dim eMountingFace As ProfileFaceName
    eMountingFace = oProfile.MountingFaceName
    Dim ePortFace As JXSEC_CODE
    ePortFace = clsAssyConnCM.ConvertMountingFaceToPort(eMountingFace)

    strError = "Constructing physical connection."
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnFromSubToACPort(pMemberDescription, pResourceManager, _
                                                                    ePortFace, "TeeWeld")
    
    'CF update when seam is deleted
    
    If Not IsSeamMovement(pAssyConn.Port2) Then
        UpdateCFWhenSeamIsDeleted pAssyConn.object
    End If
        
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC", strError).Number
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
    pMDs.RemoveAll ' Remove all the previous member descriptions
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

