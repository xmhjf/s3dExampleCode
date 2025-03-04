VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "CollarRules.ClipCDef"
Const m_ItemName As String = "CollarRules.ClipCDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\ClipCDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(oIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
  
    oIH.SetInput INPUT_PENETRATING
    oIH.SetInput INPUT_BOUNDINGPLATE
    oIH.SetInput INPUT_SLOT
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub
Public Sub ItemAggregator(oAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
  
    oAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructCollar"
  
    Dim oPDs As IJDPropertyDescriptions
    Dim strError As String
    
    Set oPDs = oAD

    ' Set the thickness
    strError = "Setting Thickness property."
    oPDs.AddProperty "Thickness", 1, IID_IJPlate, "CMCopyParentThickness", "CollarRules.CollarDefCM"
    
    ' Set the material and grade
    strError = "Setting MaterialAndGrade property."
    oPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CMCopyParentMatAndGrade", "CollarRules.CollarDefCM"
    
   ' Set the Side Of Plate
    strError = "Setting Side Of Plate property."
    oPDs.AddProperty "SideOfPlate", 3, IID_IJCollarPart, "CMSetSideOfPlate", "CollarRules.CollarDefCM"
  
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub
Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
  
    ' PC between clip and profile web left or right
    oMDs.AddMember "ClipCWebPC", 1, "CMConstructWebPC", imsCOOKIE_ID_USS_LIB
  
    ' PC between clip and base plate
    oMDs.AddMember "ClipCBasePC", 2, "CMConstructBasePC", imsCOOKIE_ID_USS_LIB
  
    Dim oMemDesc As IJDMemberDescription
  
    'PC between clip and penerated on either normal side or opposite side
    Set oMemDesc = oMDs.AddMember("ClipCLapPCN", 3, "CMConstructLapPCN", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide"
    Set oMemDesc = Nothing
  
    Set oMemDesc = oMDs.AddMember("ClipCLapPCO", 4, "CMConstructLapPCO", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide"
    Set oMemDesc = Nothing
  
    Set oMemDesc = oMDs.AddMember("ClipCCornerSnipe", 5, "CMConstructCornerSnipe", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_ItemProgid, "CMIsCornerSnipeNeeded"
    Set oMemDesc = Nothing
  
    Set oMemDesc = oMDs.AddMember("ClipCDrainHole", 6, "CMConstructDrainHole", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_ItemProgid, "CMIsDrainHoleNeeded"
    Set oMemDesc = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructCollar(ByVal oAggregatorDescription As IJDAggregatorDescription)
    Dim oCollarCM As New CollarDefCM
  
    oCollarCM.Collar_FinalConstruct oAggregatorDescription
    Set oCollarCM = Nothing
End Sub
Public Sub CMConstructWebPC(ByVal oMD As IJDMemberDescription, ByVal oResourceManager As IUnknown, ByRef oObject As Object)
 On Error GoTo ErrorHandler
 
    Dim oCollar As New StructDetailObjects.Collar
    Dim oPenetratingPart As Object
    Dim oPenetratingPort As IJPort
    Dim oCollarPort As IJPort
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim strAnswerToCreationOrder As String
    
    Set oCollar.object = oMD.CAO
    Set oPenetratingPart = oCollar.Penetrating ' Get penetrating part from wrapper
    Set oProfilePart.object = oPenetratingPart
                   
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToCreationOrder = oCommonHelper.GetSelectorAnswer(oMD.CAO, "CollarRules.RootClipSel", gsCollarCreationOrder)
    Set oCommonHelper = Nothing
    
    If strAnswerToCreationOrder = "Primary" Then
        Set oCollarPort = oCollar.SubPort(JXSEC_WEB_RIGHT)
        Set oPenetratingPort = oProfilePart.SubPort(JXSEC_WEB_RIGHT)
    Else
        Set oCollarPort = oCollar.SubPort(JXSEC_WEB_LEFT)
        Set oPenetratingPort = oProfilePart.SubPort(JXSEC_WEB_LEFT)
    End If
    Set oProfilePart = Nothing
    Set oCollar = Nothing
    
    ' Construct PC
    Dim oPCWrapper As New StructDetailObjects.PhysicalConn
    
    oPCWrapper.Create oResourceManager, _
                                  oCollarPort, _
                                  oPenetratingPort, _
                                  "TeeWeld", _
                                   oMD.CAO, _
                                   ConnectionStandard
    Set oObject = oPCWrapper.object
    Set oPCWrapper = Nothing

    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructWebPC").Number
End Sub
Public Sub CMConstructBasePC(ByVal oMD As IJDMemberDescription, _
                                                 ByVal oResourceManager As IUnknown, _
                                                 ByRef oObject As Object)
    On Error GoTo ErrorHandler
     
    Dim oCollarWrapper As New StructDetailObjects.Collar
    
    Dim oCollarBottomPort As IJPort
    Dim oBasePlatePort As IJPort
    
    Set oCollarWrapper.object = oMD.CAO
    Set oCollarBottomPort = oCollarWrapper.SubPort(JXSEC_BOTTOM)
    Set oBasePlatePort = oCollarWrapper.BasePlatePort
    Set oCollarWrapper = Nothing
    
    Dim oSDOHelper As New StructDetailObjects.Helper
    Dim oLastBasePlatePort As IJPort
    
    Set oLastBasePlatePort = oSDOHelper.GetEquivalentLastPort(oBasePlatePort)
    Set oSDOHelper = Nothing
    
    ' Construct PC
    Dim oPCWrapper As New StructDetailObjects.PhysicalConn
    
    oPCWrapper.Create oResourceManager, _
                                  oCollarBottomPort, _
                                  oLastBasePlatePort, _
                                  "TeeWeld", _
                                  oMD.CAO, _
                                  ConnectionStandard
    Set oObject = oPCWrapper.object
    Set oPCWrapper = Nothing
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructBasePC").Number
End Sub
Public Sub CMConstructLapPCN(ByVal oMD As IJDMemberDescription, _
                                                 ByVal oResourceManager As IUnknown, _
                                                 ByRef oObject As Object)
    On Error GoTo ErrorHandler
    
    Dim oCollarCM As New CollarDefCM
  
    Set oObject = oCollarCM.CreateLapPCforCollar(oMD, oResourceManager)
    Set oCollarCM = Nothing
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructLapPCN").Number
End Sub
Public Sub CMConstructLapPCO(ByVal oMD As IJDMemberDescription, _
                                            ByVal oResourceManager As IUnknown, _
                                            ByRef oObject As Object)
    On Error GoTo ErrorHandler
    
    Dim oCollarCM As New CollarDefCM
  
    Set oObject = oCollarCM.CreateLapPCforCollarOnOtherSide(oMD, oResourceManager)
    Set oCollarCM = Nothing
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructLapPCO").Number
End Sub
Public Sub CMIsCornerSnipeNeeded(ByVal oMD As IJDMemberDescription, _
                                                      ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    bIsNeeded = False

    Dim strAnswerToCreateSnipe As String
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToCreateSnipe = oCommonHelper.GetSelectorAnswer(oMD.CAO, "CollarRules.ClipCSel", gsAddCornerSnipe)
    Set oCommonHelper = Nothing

    If strAnswerToCreateSnipe = "Yes" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPCO").Number
End Sub
Public Sub CMConstructCornerSnipe(ByVal oMD As IJDMemberDescription, _
                                                       ByVal oResourceManager As IUnknown, _
                                                       ByRef oObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strAnswerToCollarCreationOrder As String
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToCollarCreationOrder = oCommonHelper.GetSelectorAnswer(oMD.CAO, _
                                                                                                            "CollarRules.RootClipSel", _
                                                                                                             gsCollarCreationOrder)
    Set oCommonHelper = Nothing
        
    Dim eWebXid As JXSEC_CODE
    
    If strAnswerToCollarCreationOrder = "Primary" Then
        eWebXid = JXSEC_WEB_RIGHT ' Must match edge name in collar symbol
    Else
        eWebXid = JXSEC_WEB_LEFT
    End If
    
    CreateCornerSnipe oMD.CAO, _
                                oResourceManager, _
                                CTX_BASE, _
                                JXSEC_BOTTOM, _
                                eWebXid, _
                                oObject
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructCornerSnipe").Number

End Sub
Public Sub CMIsDrainHoleNeeded(ByVal oMD As IJDMemberDescription, _
                                                      ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    bIsNeeded = False
    
    Dim strAnswerToAddDrainHole As String
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToAddDrainHole = oCommonHelper.GetSelectorAnswer(oMD.CAO, _
                                                                                                   "CollarRules.ClipCSel", _
                                                                                                    gsAddDrainHole)
    Set oCommonHelper = Nothing
    
    If strAnswerToAddDrainHole = "Yes" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMIsDrainHoleNeeded").Number
End Sub
Public Sub CMConstructDrainHole(ByVal oMD As IJDMemberDescription, _
                                                   ByVal oResourceManager As IUnknown, _
                                                   ByRef oObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strAnswerToCollarCreationOrder As String
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToCollarCreationOrder = oCommonHelper.GetSelectorAnswer(oMD.CAO, _
                                                                                                            "CollarRules.RootClipSel", _
                                                                                                             gsCollarCreationOrder)
    Set oCommonHelper = Nothing
                
    Dim eWebXid As JXSEC_CODE

    If strAnswerToCollarCreationOrder = "Primary" Then
        eWebXid = JXSEC_WEB_RIGHT ' Must match edge name in collar symbol
    Else
        eWebXid = JXSEC_WEB_LEFT
    End If
        
    CreateDrainHole oMD.CAO, _
                             oResourceManager, _
                             CTX_BASE, _
                             JXSEC_BOTTOM, _
                             eWebXid, _
                             oObject
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructDrainHole").Number
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
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(oDefinition As IJDSymbolDefinition)
  
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oDefinition.IJDInputs.RemoveAllInput
    oDefinition.IJDRepresentations.RemoveAllRepresentation

    oDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    oDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
    ' define the inputs
    Dim oIH As IJDInputsHelper
    
    Set oIH = New InputHelper
    oIH.Definition = oDefinition
    oIH.InitAs m_FamilyProgid
    ItemInputs oIH
    Set oIH = Nothing
    
    ' define the aggregator
    Dim oAD As IJDAggregatorDescription
  
    Set oAD = oDefinition
    ItemAggregator oAD
    Set oAD = Nothing
    
    ' define the members
    Dim oMDs As IJDMemberDescriptions
  
    Set oMDs = oDefinition
    ItemMembers oMDs
    Set oMDs = Nothing
End Sub
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                                                          ByVal defParams As Variant, _
                                                                                          ByVal oActiveConnection As Object) As Object
    Dim oDefinition As IJDSymbolDefinition
    Dim oCAFactory As New CAFactory
  
    Set oDefinition = oCAFactory.CreateCAD(oActiveConnection)
    Set oCAFactory = Nothing

    ' Set definition progId and codebase
    oDefinition.ProgId = m_ItemProgid
    oDefinition.CodeBase = CodeBase
    oDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
    
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition oDefinition
    Set IJDUserSymbolServices_InstanciateDefinition = oDefinition
    Set oDefinition = Nothing
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef oSymbolOccurence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
  MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
               & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub



