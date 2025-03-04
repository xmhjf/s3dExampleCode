VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarOneSnipeAdef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'
Private Const m_sClassName As String = "CollarOneSnipeADef"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

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
    oPDs.RemoveAll ' Removes the cached Peropety Description
    ' Set the thickness
    oPDs.AddProperty "Thickness", 1, IID_IJPlate, "CMCopyParentThickness", CUSTOMERID + "CollarRules.CollarDefCM"
    
    ' Set the material and grade
    oPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CMCopyParentMatAndGrade", CUSTOMERID + "CollarRules.CollarDefCM"
    
   ' Set the Side Of Plate
    oPDs.AddProperty "SideOfPlate", 3, IID_IJCollarPart, "CMSetSideOfPlate", CUSTOMERID + "CollarRules.CollarDefCM"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub
Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    oMDs.RemoveAll ' Removes the Cached Member description
    Dim oMemDesc As IJDMemberDescription

    '  Following members are Tee welds
    Set oMemDesc = oMDs.AddMember("BaseRightPC", 1, "CMConstructBaseRightPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBaseRightPCNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("WebRightPC", 2, "CMWebRightPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsWebRightPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("WebRightTopCornerPC", 3, "CMWebRightTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsWebRightTopCornerPCNeeded"
    Set oMemDesc = Nothing

    Set oMemDesc = oMDs.AddMember("TopFlangeRightBottomPC", 4, "CMTopFlangeRightBottomPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeRightBottomPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("TopFlangeRightBottomCornerPC", 5, "CMTopFlangeRightBottomCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeRightBottomCornerPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("TopFlangeRightPC", 6, "CMTopFlangeRightPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeRightPCNeeded"
    Set oMemDesc = Nothing
                        
    ' Following members are the lap weld for the collar
    Set oMemDesc = oMDs.AddMember("NormalSideLapPC", 7, "CMConstructNormalSideLapPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("OppositeSideLapPC", 8, "CMConstructOppositeSideLapPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide"
    Set oMemDesc = Nothing
                
    ' Corner snipe
    Set oMemDesc = oMDs.AddMember("CornerSnipe", 9, "CMConstructCornerSnipe", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsCornerSnipeNeeded"
    Set oMemDesc = Nothing

    ' Top Right PC, between collar and penetrated plate, not between collar and penetrating profile
    Set oMemDesc = oMDs.AddMember("TopRightPC", 10, "CMConstructTopRightPCBetweenCollarAndPenetrated", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopRightPCBetweenCollarAndPenetratedNeeded"
    Set oMemDesc = Nothing
    
    ' Top Right Corner PC, between collar and penetrated plate, not between collar and penetrating profile
    Set oMemDesc = oMDs.AddMember("TopRightCornerPC", 11, "CMConstructTopRightCornerPCBetweenCollarAndPenetrated", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopRightCornerPCBetweenCollarAndPenetratedNeeded"
    Set oMemDesc = Nothing
    
    ' Right PC, between collar and penetrated plate, not between collar and penetrating profile
    Set oMemDesc = oMDs.AddMember("RightPC", 12, "CMConstructRightPCBetweenCollarAndPenetrated", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsRightPCBetweenCollarAndPenetratedNeeded"
    Set oMemDesc = Nothing
    
    ' Free edge treatment at bottom flange right bottom
    Set oMemDesc = oMDs.AddMember("BottomFlangeRightBottomFET", 13, "CMConstructBottomFlangeRightBottomFET", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomFlangeRightBottomFETNeeded"
    Set oMemDesc = Nothing
    
    ' Free edge treatment at bottom flange right bottom corner
    Set oMemDesc = oMDs.AddMember("BottomFlangeRightBottomCornerFET", 14, "CMConstructBottomFlangeRightBottomCornerFET", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomFlangeRightBottomCornerFETNeeded"
    Set oMemDesc = Nothing
    
    ' Top flange right top corner PC between collar and profile for CollarBCT_A3
    Set oMemDesc = oMDs.AddMember("TopFlangeRightTopCornerPC", 15, "CMTopFlangeRightTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeRightTopCornerPCNeeded"
    Set oMemDesc = Nothing
    
    ' Top Connection for CollarBCT_A3, CollarTCT_A3 collars
    Set oMemDesc = oMDs.AddMember("TopPC", 16, "CMTopPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopPCNeeded"
    Set oMemDesc = Nothing
    
    ' Following members are the additional lap weld for the collar
    Set oMemDesc = oMDs.AddMember("NormalSideLapPC17", 17, "CMConstructNormalSideLapPC2", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide2"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("OppositeSideLapPC18", 18, "CMConstructOppositeSideLapPC2", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide2"
    Set oMemDesc = Nothing
    
    'additional base pc when collar crosses seam
    Set oMemDesc = oMDs.AddMember("BaseRight2PC", 19, "CMConstructBaseRightPC2", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateBaseRightPC2"
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
Public Sub CMIsCornerSnipeNeeded(ByVal oMD As IJDMemberDescription, _
                                                      ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_CollarCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oSL As IJDSelectorLogic
    Dim strAnswerToCreateSnipe As String
        
    Dim oCommonHelper As New DefinitionHlprs.CommonHelper

     'TR-169311 Trying to retrieve the question's answer
    'without harcoding the path
    GetSelectorAnswer oMD.CAO, gsAddCornerSnipe, strAnswerToCreateSnipe
    
    Set oCommonHelper = Nothing
    
    If strAnswerToCreateSnipe = "Yes" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMIsCornerSnipeNeeded").Number
End Sub
Public Sub CMConstructCornerSnipe(ByVal oMD As IJDMemberDescription, _
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
    Err.Raise LogError(Err, MODULE, "CMConstructCornerSnipe").Number

End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_sRuleName
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
    oIH.definition = oDefinition
    oIH.InitAs m_sFamilyProgID
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
    oDefinition.ProgId = m_sRuleProgID
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

Public Sub CMIsBaseRightPCNeeded(ByVal oMD As IJDMemberDescription, _
                                 ByRef bIsNeeded As Boolean)
    bIsNeeded = True
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
   
End Sub

Public Sub CMIsWebRightPCNeeded(ByVal oMD As IJDMemberDescription, _
                                 ByRef bIsNeeded As Boolean)
    bIsNeeded = True
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
   
End Sub

Public Sub CMIsWebRightTopCornerPCNeeded(ByVal oMD As IJDMemberDescription, _
                                         ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
    
   ' Dependent on cross section
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oProfileWrapper As New StructDetailObjects.ProfilePart
   Dim sSectionType As String
    
   Set oCollarWrapper.object = oMD.CAO
   Dim bIsPenetratingObjPlate As Boolean
   
   If TypeOf oCollarWrapper.Penetrating Is IJProfile Then
    Set oProfileWrapper.object = oCollarWrapper.Penetrating
    sSectionType = oProfileWrapper.sectionType
    bIsPenetratingObjPlate = False
   ElseIf TypeOf oCollarWrapper.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oCollarWrapper.Penetrating, oCollarWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Set oSlotMappingRule = Nothing
    bIsPenetratingObjPlate = True
   Else
   
   End If
        
   Select Case sSectionType
      Case "B", "EA", "UA"
        'Not needed if Penetrating Object is a plate, because there is no radius on the section alias
        If Not bIsPenetratingObjPlate Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
      Case Else
         bIsNeeded = False
   End Select
                
End Sub

Public Sub CMIsTopFlangeRightBottomPCNeeded(ByVal oMD As IJDMemberDescription, _
                                            ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   ' Dependent on cross section
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oProfileWrapper As New StructDetailObjects.ProfilePart
   Dim sSectionType As String
    
   Set oCollarWrapper.object = oMD.CAO
   
   If TypeOf oCollarWrapper.Penetrating Is IJProfile Then
    Set oProfileWrapper.object = oCollarWrapper.Penetrating
    sSectionType = oProfileWrapper.sectionType
   ElseIf TypeOf oCollarWrapper.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oCollarWrapper.Penetrating, oCollarWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Set oSlotMappingRule = Nothing
   Else
   
   End If
      
   Select Case sSectionType
      Case "B", "EA", "UA", "BUT", "BUTL2"
         bIsNeeded = True
      Case Else
         bIsNeeded = False
   End Select
   
End Sub

Public Sub CMIsTopFlangeRightBottomCornerPCNeeded(ByVal oMD As IJDMemberDescription, _
                                                  ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   ' Dependent on cross section
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oProfileWrapper As New StructDetailObjects.ProfilePart
   Dim sSectionType As String
    
   Set oCollarWrapper.object = oMD.CAO
   Dim bIsPenetratingObjPlate As Boolean
   If TypeOf oCollarWrapper.Penetrating Is IJProfile Then
    Set oProfileWrapper.object = oCollarWrapper.Penetrating
    sSectionType = oProfileWrapper.sectionType
    bIsPenetratingObjPlate = False
   ElseIf TypeOf oCollarWrapper.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oCollarWrapper.Penetrating, oCollarWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Set oSlotMappingRule = Nothing
    bIsPenetratingObjPlate = True
   Else
   
   End If
      
   Select Case sSectionType
      Case "B", "EA", "UA"
        'Not needed if Penetrating Object is a plate, because there is no radius on the section alias
        If Not bIsPenetratingObjPlate Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
      Case Else
         bIsNeeded = False
   End Select

End Sub

Public Sub CMIsTopFlangeRightPCNeeded(ByVal oMD As IJDMemberDescription, _
                                 ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   ' Dependent on cross section
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oProfileWrapper As New StructDetailObjects.ProfilePart
   Dim sSectionType As String
    
   Set oCollarWrapper.object = oMD.CAO
   
   If TypeOf oCollarWrapper.Penetrating Is IJProfile Then
    Set oProfileWrapper.object = oCollarWrapper.Penetrating
    sSectionType = oProfileWrapper.sectionType
   ElseIf TypeOf oCollarWrapper.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oCollarWrapper.Penetrating, oCollarWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Set oSlotMappingRule = Nothing
   Else
   
   End If
      
   Select Case sSectionType
      Case "EA", "UA", "BUT", "BUTL2"
         bIsNeeded = True
      Case Else
         bIsNeeded = False
   End Select
End Sub

Public Sub CMIsTopRightPCBetweenCollarAndPenetratedNeeded( _
                  ByVal oMD As IJDMemberDescription, _
                  ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   bIsNeeded = False
   
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
   
   If InStr(sItemName, "_A2") > 0 Then
      bIsNeeded = True
   End If
      
End Sub

Public Sub CMIsTopRightCornerPCBetweenCollarAndPenetratedNeeded( _
               ByVal oMD As IJDMemberDescription, _
               ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   bIsNeeded = False
    
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
    
   If InStr(sItemName, "_A2") > 0 Then
      bIsNeeded = True
   End If

End Sub
Public Sub CMIsRightPCBetweenCollarAndPenetratedNeeded( _
                         ByVal oMD As IJDMemberDescription, _
                         ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   bIsNeeded = False
    
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
    
   If InStr(sItemName, "_A2") > 0 Then
      bIsNeeded = True
   End If

End Sub

Public Sub CMIsBottomFlangeRightBottomFETNeeded( _
                              ByVal oMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_FreeEdgeTreatment) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
   
   If InStr(sItemName, "_A") > 0 Then
      If InStr(sItemName, "_A2") > 0 Or _
         InStr(sItemName, "_A3") > 0 Then
         bIsNeeded = False
      Else
         bIsNeeded = True
      End If
   Else
      bIsNeeded = False
   End If
End Sub

Public Sub CMIsBottomFlangeRightBottomCornerFETNeeded( _
                              ByVal oMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_FreeEdgeTreatment) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
   
   If InStr(sItemName, "_A") > 0 Then
      If InStr(sItemName, "_A2") > 0 Or _
         InStr(sItemName, "_A3") > 0 Then
         bIsNeeded = False
      Else
         bIsNeeded = True
      End If
   Else
      bIsNeeded = False
   End If
End Sub
Public Sub CMIsTopFlangeRightTopCornerPCNeeded( _
                           ByVal oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   bIsNeeded = False
   
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
   
   If InStr(sItemName, "CollarBCT_A3") > 0 Then
      bIsNeeded = True
   End If

End Sub

Public Sub CMIsTopPCNeeded( _
                           ByVal oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
       
   bIsNeeded = False
   
   Dim oSmartOccurrence As IJSmartOccurrence
   Dim oSmartItem As IJSmartItem
   Dim sItemName As String

   Set oSmartOccurrence = oMD.CAO
   Set oSmartItem = oSmartOccurrence.ItemObject
   sItemName = oSmartItem.Name
   
   If InStr(sItemName, "_A3") > 0 Then
      bIsNeeded = True
   End If
End Sub


