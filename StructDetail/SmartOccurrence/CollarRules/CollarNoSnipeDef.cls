VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarNoSnipeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "CollarRules.CollarNoSnipeDef"
Const m_ItemName As String = "CollarRules.CollarNoSnipeDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\CollarNoSnipeDef.cls"

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
    
    Set oPDs = oAD
    ' Set the thickness
    oPDs.AddProperty "Thickness", 1, IID_IJPlate, "CMCopyParentThickness", "CollarRules.CollarDefCM"
    
    ' Set the material and grade
    oPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CMCopyParentMatAndGrade", "CollarRules.CollarDefCM"
    
   ' Set the Side Of Plate
    oPDs.AddProperty "SideOfPlate", 3, IID_IJCollarPart, "CMSetSideOfPlate", "CollarRules.CollarDefCM"
      
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub
Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
   
    '  Following members are Tee welds and are always required and therefore no conditional is necessary
    oMDs.AddMember "CLTPC1", 1, "CMWebLeftPC", CMLIBRARY_COLLARRULES
    oMDs.AddMember "CLTPC2", 2, "CMWebRightPC", CMLIBRARY_COLLARRULES
    oMDs.AddMember "CLTPC3", 3, "CMTopPC", CMLIBRARY_COLLARRULES

    ' Following members are the lap weld for the collar.  There is a conditional based on the side of part.
    Dim oMemDesc As IJDMemberDescription
    
    Set oMemDesc = oMDs.AddMember("CLTPC4", 4, "CMConstructPC2", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide"

    Set oMemDesc = Nothing
    Set oMemDesc = oMDs.AddMember("CLTPC5", 5, "CMConstructPC2b", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide"

    
    '  Following members are Tee welds and are conditional base on profile type.
    Set oMemDesc = oMDs.AddMember("CLTPC6", 6, "CMTopFlangeLeftBottomPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeLeftBottomCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC7", 7, "CMTopFlangeRightBottomPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightBottomCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC8", 8, "CMTopFlangeLeftPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeLeftCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC9", 9, "CMTopFlangeRightPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC10", 10, "CMWebRightTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMWebRightTopCornerCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC11", 11, "CMWebLeftTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMWebLeftTopCornerCon"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("CLTPC12", 12, "CMTopFlangeRightBottomCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightBottomCornerCon"
    Set oMemDesc = Nothing

'    Set oMemDesc = oMDs.AddMember("CLTPC13", 13, "CMTopFlangeLeftBottomCornerPC", CMLIBRARY_COLLARRULES)
'    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeLeftBottomCornerCon"
'    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("CLTPC14", 14, "CMTopFlangeRightTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightTopCornerCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC15", 15, "CMTopFlangeRightTopPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightTopCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC16", 16, "CMWebRightTopPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMWebRightTopCon"
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
Public Sub CMConstructPC2(ByVal oMD As IJDMemberDescription, _
                                           ByVal oResourceManager As IUnknown, _
                                           ByRef oObject As Object)
    Dim oCollarCM As New CollarDefCM
  
    Set oObject = oCollarCM.CreateLapPCforCollar(oMD, oResourceManager)
    Set oCollarCM = Nothing
End Sub
Public Sub CMConstructPC2b(ByVal oMD As IJDMemberDescription, _
                                             ByVal oResourceManager As IUnknown, _
                                             ByRef oObject As Object)
    Dim oCollarCM As New CollarDefCM
  
    Set oObject = oCollarCM.CreateLapPCforCollarOnOtherSide(oMD, oResourceManager)
    Set oCollarCM = Nothing
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
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal oActiveConnection As Object) As Object
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
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal oOutputColl As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef oSymbolOccurence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                 & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

