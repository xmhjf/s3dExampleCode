VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarTwoSnipesDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "CollarRules.CollarTwoSnipesDef"
Const m_ItemName As String = "CollarRules.CollarTwoSnipesDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\CollarTwoSnipesDef.cls"

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
    Set oMemDesc = Nothing
    
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

'    Set pMemDesc = pMDs.AddMember("CLTPC13", 13, "CMTopFlangeLeftBottomCornerPC", CMLIBRARY_COLLARRULES)
'    pMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeLeftBottomCornerCon"
'    Set pMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("CLTPC14", 14, "CMTopFlangeRightTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightTopCornerCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC15", 15, "CMTopFlangeRightTopPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMTopFlangeRightTopCon"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("CLTPC16", 16, "CMWebRightTopPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMWebRightTopCon"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("CornerSnipeWeb_Left", 17, "CMConstructCornerSnipeWeb_Left", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_ItemProgid, "CMIsCornerSnipeWeb_LeftNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("CornerSnipeWeb_Right", 18, "CMConstructCornerSnipeWeb_Right", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_ItemProgid, "CMIsCornerSnipeWeb_RightNeeded"
    Set oMemDesc = Nothing
        
    ' PC between clip and base plate: always created ... No conditional required
    Set oMemDesc = oMDs.AddMember("BaseRightPC", 19, "CMConstructBaseRightPC", imsCOOKIE_ID_USS_LIB)
    Set oMemDesc = Nothing
        
    ' PC between clip and base plate: always created ... No conditional required
    Set oMemDesc = oMDs.AddMember("BaseLeftPC", 20, "CMConstructBaseLeftPC", imsCOOKIE_ID_USS_LIB)
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

Public Sub CMIsCornerSnipeWeb_LeftNeeded(ByVal oMD As IJDMemberDescription, _
                                                                    ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    Dim oSL As IJDSelectorLogic
    Dim strAnswerToCreateSnipes As String
        
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToCreateSnipes = oCommonHelper.GetSelectorAnswer(oMD.CAO, _
                                                                                                    "CollarRules.CollarTwoSnipesSel", _
                                                                                                    gsAddCornerSnipes)
    Set oCommonHelper = Nothing
    If strAnswerToCreateSnipes = gsWebLeft Or strAnswerToCreateSnipes = gsBoth Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMIsCornerSnipeNeededWeb_Left").Number
End Sub

Public Sub CMConstructCornerSnipeWeb_Left(ByVal oMD As IJDMemberDescription, _
                                                                    ByVal oResourceManager As IUnknown, _
                                                                    ByRef oObject As Object)
    On Error GoTo ErrorHandler
        
    CreateCornerSnipe oMD.CAO, _
                                oResourceManager, _
                                CTX_BASE, _
                                JXSEC_BOTTOM_LEFT, _
                                JXSEC_WEB_LEFT, _
                                oObject
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructCornerSnipeWeb_Left").Number

End Sub

Public Sub CMIsCornerSnipeWeb_RightNeeded(ByVal oMD As IJDMemberDescription, _
                                                                      ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    Dim oSL As IJDSelectorLogic
    Dim strAnswerToCreateSnipes As String
        
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

    strAnswerToCreateSnipes = oCommonHelper.GetSelectorAnswer(oMD.CAO, _
                                                                                                    "CollarRules.CollarTwoSnipesSel", _
                                                                                                    gsAddCornerSnipes)
    Set oCommonHelper = Nothing
    
    If strAnswerToCreateSnipes = gsWebRight Or strAnswerToCreateSnipes = gsBoth Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMIsCornerSnipeWeb_RightNeeded").Number
End Sub

Public Sub CMConstructCornerSnipeWeb_Right(ByVal oMD As IJDMemberDescription, _
                                                                      ByVal oResourceManager As IUnknown, _
                                                                      ByRef oObject As Object)
    On Error GoTo ErrorHandler
    
    CreateCornerSnipe oMD.CAO, _
                                oResourceManager, _
                                CTX_BASE, _
                                JXSEC_BOTTOM_RIGHT, _
                                JXSEC_WEB_RIGHT, _
                                oObject
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructCornerSnipeWeb_Right").Number
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
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSblOcc As Object, ByVal repName As String, ByVal oOutputColl As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef oSymbolOccurence As Object, ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                 & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CMConstructBaseRightPC(ByVal oMD As IJDMemberDescription, _
                                  ByVal oResourceManager As IUnknown, _
                                  ByRef oObject As Object)
    '
    ' Create Physical Connection between Collar Bottom Edge and Base Plate
    ConstructBaseEdgePC JXSEC_BOTTOM_RIGHT, oMD, oResourceManager, oObject
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructBaseRightPC").Number
End Sub

Public Sub CMConstructBaseLeftPC(ByVal oMD As IJDMemberDescription, _
                                 ByVal oResourceManager As IUnknown, _
                                 ByRef oObject As Object)
    '
    ' Create Physical Connection between Collar Bottom Left Edge and Base Plate
    ConstructBaseEdgePC JXSEC_BOTTOM_LEFT, oMD, oResourceManager, oObject
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructBaseLeftPC").Number
End Sub





