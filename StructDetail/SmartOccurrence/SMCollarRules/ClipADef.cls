VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipADef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "CollarRules.ClipADef"
Const m_ItemName As String = CUSTOMERID + "CollarRules.ClipADef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CollarRules\ClipADef.cls"

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
              
    pPDs.RemoveAll 'Removes cached Property description (best practice)
    
    Dim strError As String
    
    ' Set the thickness
    strError = "Setting Thickness property."
    pPDs.AddProperty "Thickness", 1, IID_IJPlate, "CMCopyParentThickness", CUSTOMERID + "CollarRules.CollarDefCM"
    
    ' Set the material and grade
    strError = "Setting MaterialAndGrade property."
    pPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CMCopyParentMatAndGrade", CUSTOMERID + "CollarRules.CollarDefCM"
    
   ' Set the Side Of Plate
    strError = "Setting Side Of Plate property."
    pPDs.AddProperty "SideOfPlate", 3, IID_IJCollarPart, "CMSetSideOfPlate", CUSTOMERID + "CollarRules.CollarDefCM"
    
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
  On Error GoTo ErrorHandler
  pMDs.RemoveAll  'Removes cached Member Descriptions(best practice)
  
  Dim pMemDesc As IJDMemberDescription
  Set pMemDesc = pMDs.AddMember("ClipAPC1", 1, "CMConstructPC1", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional m_ItemProgid, "CMIsPCNeeded"
  
  Set pMemDesc = Nothing
  Set pMemDesc = pMDs.AddMember("ClipAPC2", 2, "CMConstructPC2", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide"
  
  Set pMemDesc = Nothing
  Set pMemDesc = pMDs.AddMember("ClipAPC2b", 3, "CMConstructPC2b", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide"
 
  ' Note:ClipATopRightCornerSnipe,ClipABottomRightCornerSnipe are not related to "AddCornerSnipe" question
  Set pMemDesc = Nothing
  Set pMemDesc = pMDs.AddMember("ClipATopRightCornerSnipe", 4, "CMConstructTopRightCornerSnipe", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional m_ItemProgid, "CMIsTopRightCornerSnipeNeeded"
  
  Set pMemDesc = Nothing
  Set pMemDesc = pMDs.AddMember("ClipABottomRightCornerSnipe", 5, "CMConstructBottomRightCornerSnipe", imsCOOKIE_ID_USS_LIB)
  pMemDesc.SetCMConditional m_ItemProgid, "CMIsBottomRightCornerSnipeNeeded"
  
  Set pMemDesc = Nothing
  'additional lap PCs when collar overlaps with more than one penetrated part
  Set pMemDesc = pMDs.AddMember("ClipAPC6", 6, "CMConstructPC2Lap2", CMLIBRARY_COLLARRULES)
  pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide2"
    
  Set pMemDesc = Nothing
  Set pMemDesc = pMDs.AddMember("ClipAPC7", 7, "CMConstructPC2bLap2b", CMLIBRARY_COLLARRULES)
  pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide2"
  
  Set pMemDesc = Nothing
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
    
    ' Get Penetrating Part web Right port
    Dim oPenetratingPart As Object ' This is the penetrating part
    sERROR = "Getting Penetrating object"
    Dim oPenetratingPort As IJPort
    Set oPenetratingPart = pCollar.Penetrating ' Get penetrating part from wrapper
    
    If TypeOf oPenetratingPart Is IJProfile Then
      Dim oProfilePart As New StructDetailObjects.ProfilePart
      sERROR = "Setting Profile Part wrapper object to penetrating object"
      Set oProfilePart.object = oPenetratingPart
      sERROR = "Getting Penetrating object's port"
      Set oPenetratingPort = oProfilePart.SubPort(JXSEC_WEB_RIGHT)
      
    ElseIf TypeOf oPenetratingPart Is IJPlate Then
      Dim oBasePort As Object
      Dim oMappedPorts As JCmnShp_CollectionAlias
      Set oMappedPorts = New Collection

      Dim pSDOHelper As New StructDetailObjects.Helper
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      oSlotMappingRule.GetEmulatedPorts oPenetratingPart, pCollar.Penetrated, oBasePort, oMappedPorts
      
      Set oPenetratingPort = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item(CStr(JXSEC_WEB_RIGHT)))
    Else
      Exit Sub
    End If
    Set oPenetratingPart = Nothing
    ' Get the correct sub port on Collar for this PC
    Dim pCollarPort As IJPort
    Set pCollarPort = pCollar.SubPort(JXSEC_WEB_RIGHT)
    
    ' Construct PC
    Dim clsCollarCM As New CollarDefCM
    Set m_oPhysConn1 = clsCollarCM.CAConstruct_PhysConn(pMemberDescription, pResourceManager, _
                                                         "TeeWeld", _
                                                         pCollarPort, oPenetratingPort)
    Set pObject = m_oPhysConn1
    Set clsCollarCM = Nothing
    Set oPenetratingPort = Nothing
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
  pIH.definition = pDefinition
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

Public Sub CMIsTopRightCornerSnipeNeeded(ByVal oMD As IJDMemberDescription, _
                                         ByRef bIsNeeded As Boolean)
   On Error GoTo ErrorHandler
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_CollarCornerFeature) Then
       bIsNeeded = False
       Exit Sub
   End If
    
   bIsNeeded = True
   
   ' Top Right Corner snipe will not be applied if followings are true:
   ' Penetrated is a LBH
   ' BasePlate is HULL
   ' and collar right side is > 0.065 and < 0.090
   
   Dim oCollar As Object
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim bIsProfileOnHullPenetratingLBH As Boolean
   
   Set oCollar = oMD.CAO
   Set oCollarWrapper.object = oCollar
         
   bIsProfileOnHullPenetratingLBH = CheckProfileOnHullPenetratingLBH(oMD.CAO)
   If bIsProfileOnHullPenetratingLBH Then
      Dim dAngle As Double
      Dim oSlot As Object
      
      Set oSlot = oCollarWrapper.Slot
      dAngle = 0
      GetSlotAngles oSlot, JXSEC_WEB_RIGHT, dAngle
      If dAngle > HALF_PI Then
         Dim dClipRightLength As Double
         
         dClipRightLength = EstimateClipRightSideLength(oSlot, dAngle)
         If dClipRightLength > 0.065 And dClipRightLength < 0.09 Then
            bIsNeeded = False
         End If
      End If
   End If
    
Exit Sub

ErrorHandler:
   Err.Raise LogError(Err, MODULE, "CMIsTopRightCornerSnipeNeeded").Number
End Sub
Public Sub CMConstructTopRightCornerSnipe(ByVal oMD As IJDMemberDescription, _
                                          ByVal oResourceManager As IUnknown, _
                                          ByRef oObject As Object)
   On Error GoTo ErrorHandler
   Dim eWebXid As JXSEC_CODE
    
   eWebXid = JXSEC_RIGHT ' Must match edge name in collar symbol
   CreateCornerSnipe oMD.CAO, _
                     oResourceManager, _
                     CTX_BASE, _
                     JXSEC_TOP, _
                     eWebXid, _
                     oObject
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "CMConstructTopRightCornerSnipe").Number

End Sub

Public Sub CMIsBottomRightCornerSnipeNeeded(ByVal oMD As IJDMemberDescription, _
                                         ByRef bIsNeeded As Boolean)
   On Error GoTo ErrorHandler
   'Always needed
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_CollarCornerFeature) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   bIsNeeded = True
    
Exit Sub

ErrorHandler:
   Err.Raise LogError(Err, MODULE, "CMIsTopRightCornerSnipeNeeded").Number
End Sub
Public Sub CMConstructBottomRightCornerSnipe(ByVal oMD As IJDMemberDescription, _
                                             ByVal oResourceManager As IUnknown, _
                                             ByRef oObject As Object)
   On Error GoTo ErrorHandler
   Dim eWebXid As JXSEC_CODE
    
   eWebXid = JXSEC_RIGHT ' Must match edge name in collar symbol
   CreateCornerSnipe oMD.CAO, _
                     oResourceManager, _
                     CTX_BASE, _
                     JXSEC_BOTTOM, _
                     eWebXid, _
                     oObject
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "CMConstructBottomRightCornerSnipe").Number

End Sub

'********************************************************************************************
' Method      : CMIsPCNeeded
' Description :
'********************************************************************************************
Public Sub CMIsPCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
  Const METHOD = "::CMIsPCNeeded"
  On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
