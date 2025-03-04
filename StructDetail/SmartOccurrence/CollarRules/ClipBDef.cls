VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipBDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "CollarRules.ClipBDef"
Const m_ItemName As String = "CollarRules.ClipBDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\ClipBDef.cls"

Dim m_oPhysConn1 As Object
Dim m_oPhysConn2 As Object

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_BOUNDINGPLATE
  pIH.SetInput INPUT_SLOT
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
  On Error GoTo ErrorHandler
  
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructCollar"
  
    
    Dim pPDs As IJDPropertyDescriptions
    Set pPDs = pAD
              
    Dim strError As String
    
    ' Set the thickness
    strError = "Setting Thickness property."
    pPDs.AddProperty "Thickness", 1, IID_IJPlate, "CMCopyParentThickness", "CollarRules.CollarDefCM"
    
    ' Set the material and grade
    strError = "Setting MaterialAndGrade property."
    pPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CMCopyParentMatAndGrade", "CollarRules.CollarDefCM"
    
   ' Set the Side Of Plate
    strError = "Setting Side Of Plate property."
    pPDs.AddProperty "SideOfPlate", 3, IID_IJCollarPart, "CMSetSideOfPlate", "CollarRules.CollarDefCM"
    
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
  On Error GoTo ErrorHandler
  
  pMDs.AddMember "ClipBPC1", 1, "CMConstructPC1", imsCOOKIE_ID_USS_LIB
  
  Dim pMemDesc As IJDMemberDescription
  Set pMemDesc = pMDs.AddMember("ClipAPC2", 2, "CMConstructPC2", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide"
  
  Set pMemDesc = Nothing
  Set pMemDesc = pMDs.AddMember("ClipAPC2b", 3, "CMConstructPC2b", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructCollar(ByVal pAggregatorDescription As IJDAggregatorDescription)
  Dim clsCollarCM As New CollarDefCM
  clsCollarCM.Collar_FinalConstruct pAggregatorDescription
  Set clsCollarCM = Nothing
End Sub
Public Sub CMConstructPC1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
 On Error GoTo ErrorHandler
 
    Dim sERROR As String
 
    ' Get Class
    Dim pCollar As New StructDetailObjects.Collar ' create new collar wrapper class
    Set pCollar.object = pMemberDescription.CAO ' Intialize the wrapper with slot
    
    ' Get Penetrating Part web LEFT port
    Dim oPenetratingPart As Object ' This is the penetrating part
    sERROR = "Getting Penetrating object"
    Set oPenetratingPart = pCollar.Penetrating ' Get penetrating part from wrapper
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    sERROR = "Setting Profile Part wrapper object to penetrating object"
    Set oProfilePart.object = oPenetratingPart
    Dim oPenetratingPort As IJPort
    sERROR = "Getting Penetrating object's port"
    Set oPenetratingPort = oProfilePart.SubPort(JXSEC_WEB_LEFT)

    ' Get the correct sub port on Collar for this PC
    Dim pCollarPort As IJPort
    Set pCollarPort = pCollar.SubPort(JXSEC_WEB_LEFT)
    
    ' Construct PC
    Dim clsCollarCM As New CollarDefCM
    Set m_oPhysConn1 = clsCollarCM.CAConstruct_PhysConn(pMemberDescription, pResourceManager, _
                                                         "TeeWeld", _
                                                         pCollarPort, oPenetratingPort)
    Set pObject = m_oPhysConn1
    Set clsCollarCM = Nothing
  
  Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMConstructPC1", sERROR).Number
End Sub
Public Sub CMConstructPC2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim clsCollarCM As New CollarDefCM
  Set m_oPhysConn1 = clsCollarCM.CreateLapPCforCollar(pMemberDescription, pResourceManager)
  Set pObject = m_oPhysConn1
  Set clsCollarCM = Nothing
End Sub
Public Sub CMConstructPC2b(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim clsCollarCM As New CollarDefCM
  Set m_oPhysConn1 = clsCollarCM.CreateLapPCforCollarOnOtherSide(pMemberDescription, pResourceManager)
  Set pObject = m_oPhysConn1
  Set clsCollarCM = Nothing
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
  
  ' define the aggregator
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  ItemAggregator pAD
   
  ' define the members
  Dim pMDs As IJDMemberDescriptions
  Set pMDs = pDefinition
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
Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
  MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
    & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub



