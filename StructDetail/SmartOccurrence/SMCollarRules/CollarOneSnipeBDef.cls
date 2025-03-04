VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarOneSnipeBDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'
Private Const m_sClassName As String = "CollarOneSnipeBDef"

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
    oPDs.RemoveAll ' Removes the cached Property description
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
    oMDs.RemoveAll ' Removes the cached member description
    Dim oMemDesc As IJDMemberDescription

    '  Following members are Tee welds
    Set oMemDesc = oMDs.AddMember("BaseLeftPC", 1, "CMConstructBaseLeftPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBaseLeftPCNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("WebLeftPC", 2, "CMWebLeftPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsWebLeftPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("WebLeftTopCornerPC", 3, "CMWebLeftTopCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsWebLeftTopCornerPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("TopFlangeLeftBottomPC", 4, "CMTopFlangeLeftBottomPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeLeftBottomPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("TopFlangeLeftBottomCornerPC", 5, "CMTopFlangeLeftBottomCornerPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeLeftBottomCornerPCNeeded"
    Set oMemDesc = Nothing
            
    Set oMemDesc = oMDs.AddMember("TopFlangeLeftPC", 6, "CMTopFlangeLeftPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopFlangeLeftPCNeeded"
    Set oMemDesc = Nothing
            
    ' Following members are the lap weld for the collar.  There is a conditional based on the side of part.
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
    
    ' Left PC, between collar and penetrated plate, not between collar and penetrating profile
    Set oMemDesc = oMDs.AddMember("LeftPC", 10, "CMConstructLeftPCBetweenCollarAndPenetrated", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsLeftPCBetweenCollarAndPenetratedNeeded"
    Set oMemDesc = Nothing
    
    ' Top Left Corner PC, between collar and penetrated plate, not between collar and penetrating profile
    Set oMemDesc = oMDs.AddMember("TopLeftCornerPC", 11, "CMConstructTopLeftCornerPCBetweenCollarAndPenetrated", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopLeftCornerPCBetweenCollarAndPenetratedNeeded"
    Set oMemDesc = Nothing
    
    ' Top Left PC, between collar and penetrated plate, not between collar and penetrating profile
    Set oMemDesc = oMDs.AddMember("TopLeftPC", 12, "CMConstructTopLeftPCBetweenCollarAndPenetrated", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopLeftPCBetweenCollarAndPenetratedNeeded"
    Set oMemDesc = Nothing

    ' Free edge treatment at bottom flange left bottom
    Set oMemDesc = oMDs.AddMember("BottomFlangeLeftBottomFET", 13, "CMConstructBottomFlangeLeftBottomFET", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomFlangeLeftBottomFETNeeded"
    Set oMemDesc = Nothing
    
    ' Free edge treatment at bottom flange left bottom
    Set oMemDesc = oMDs.AddMember("BottomFlangeLeftBottomCornerFET", 14, "CMConstructBottomFlangeLeftBottomCornerFET", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomFlangeLeftBottomCornerFETNeeded"
    Set oMemDesc = Nothing

    ' Top Connection for CollarTCT_B3 collars
    Set oMemDesc = oMDs.AddMember("TopPC", 15, "CMTopPC", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopPCNeeded"
    Set oMemDesc = Nothing
    
    'additional lap Pcs when collar laps with more than one penetrated part
    Set oMemDesc = oMDs.AddMember("NormalSideLapPC2", 16, "CMConstructNormalSideLapPC2", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipNormalSide2"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("OppositeSideLapPC2", 17, "CMConstructOppositeSideLapPC2", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateClipOppositeSide2"
    Set oMemDesc = Nothing
    
    'additional base pc when collar crosses seam
    Set oMemDesc = oMDs.AddMember("BaseLeft2PC", 18, "CMConstructBaseLeftPC2", CMLIBRARY_COLLARRULES)
    oMemDesc.SetCMConditional CMLIBRARY_COLLARRULES, "CMCreateBaseLeftPC2"
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
                      JXSEC_BOTTOM_LEFT, _
                      JXSEC_WEB_LEFT, _
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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CMIsBaseLeftPCNeeded(ByVal oMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
                                       
   bIsNeeded = True
End Sub

Public Sub CMIsWebLeftPCNeeded(ByVal oMD As IJDMemberDescription, _
                               ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
       bIsNeeded = False
       Exit Sub
   End If
          
   bIsNeeded = True
End Sub

Public Sub CMIsWebLeftTopCornerPCNeeded(ByVal oMD As IJDMemberDescription, _
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
      Case "B", "EA", "UA"
         bIsNeeded = True
      Case Else
         bIsNeeded = False
   End Select
                
End Sub

Public Sub CMIsTopFlangeLeftBottomPCNeeded(ByVal oMD As IJDMemberDescription, _
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

Public Sub CMIsTopFlangeLeftBottomCornerPCNeeded(ByVal oMD As IJDMemberDescription, _
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
      Case "B", "EA", "UA"
         bIsNeeded = True
      Case Else
         bIsNeeded = False
   End Select

End Sub

Public Sub CMIsTopFlangeLeftPCNeeded(ByVal oMD As IJDMemberDescription, _
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

Public Sub CMIsLeftPCBetweenCollarAndPenetratedNeeded( _
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
    
   If InStr(sItemName, "_B2") > 0 Then
      bIsNeeded = True
   End If

End Sub
Public Sub CMIsTopLeftCornerPCBetweenCollarAndPenetratedNeeded( _
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
    
   If InStr(sItemName, "_B2") > 0 Then
      bIsNeeded = True
   End If

End Sub
Public Sub CMIsTopLeftPCBetweenCollarAndPenetratedNeeded( _
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
   
   If InStr(sItemName, "_B2") > 0 Then
      bIsNeeded = True
   End If
      
End Sub

Public Sub CMIsBottomFlangeLeftBottomFETNeeded( _
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
   
   If InStr(sItemName, "_B") > 0 Then
      If InStr(sItemName, "_B2") > 0 Or _
         InStr(sItemName, "_B3") > 0 Then
         bIsNeeded = False
      Else
         bIsNeeded = True
      End If
   Else
      bIsNeeded = False
   End If

End Sub

Public Sub CMIsBottomFlangeLeftBottomCornerFETNeeded( _
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
   
   If InStr(sItemName, "_B") > 0 Then
      If InStr(sItemName, "_B2") > 0 Or _
         InStr(sItemName, "_B3") > 0 Then
         bIsNeeded = False
      Else
         bIsNeeded = True
      End If
   Else
      bIsNeeded = False
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
   
   If InStr(sItemName, "_B3") > 0 Then
      bIsNeeded = True
   End If
End Sub


