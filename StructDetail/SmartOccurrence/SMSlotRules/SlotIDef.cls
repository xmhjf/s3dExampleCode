VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotIDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'pnalugol - Nov1 2011 - Changes done to support 202656, added
'       new item members to create Corner features as child of Slots

Option Explicit

Implements IJDUserSymbolServices

' Modify class name to particular rule
Private Const m_sClassName As String = "SlotIDef"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"
Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
Dim SGOSurfaceBodyUtilities As GSCADShipGeomOps.SGOSurfaceBodyUtilities
'Newly created member objects
Private m_oPhysConn As Object

'Error management
Private sERROR As String
Private sMETHOD As String

Public Sub ItemInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PENETRATING
  pIH.SetInput INPUT_PENETRATED
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
  On Error GoTo ErrorHandler
  
  pAD.SetCMFinalConstruct CUSTOMERID + "SlotRules.SlotIDef", "CMFinalConstructSlot"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
  On Error GoTo ErrorHandler

    Dim oMemDesc As IJDMemberDescription
    Dim nIndex As Integer
    Dim nCount As Integer
    Dim sEdgeName As String
    
    nIndex = 0
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotIPC1", 1, "CMConstructLeftPC", CUSTOMERID + "SlotRules.SlotIDef")
    oMemDesc.SetCMConditional CUSTOMERID + "SlotRules.SlotIDef", "CMIsLeftPCNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotIPC2", 2, "CMConstructRightPC", CUSTOMERID + "SlotRules.SlotIDef")
    oMemDesc.SetCMConditional CUSTOMERID + "SlotRules.SlotIDef", "CMIsRightPCNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotIFET1", 3, "CMConstructFET1", CMLIBRARY_SLOTDEFCM)
    oMemDesc.SetCMConditional CMLIBRARY_SLOTDEFCM, "CMGetEdgeTreatment"
    Set oMemDesc = Nothing

    Dim sEdgeNames(1 To 20) As String
    
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "Top"

    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "TopFlangeRightTopCorner"
    
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "TopFlangeRight"
        
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "TopFlangeRightBottomCorner"
    
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "TopFlangeRightBottom"
    
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "WebRightTopCorner"
        
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "TopFlangeLeftBottom"
        
    nIndex = nIndex + 1
    sEdgeNames(nIndex) = "TopFlangeLeft"
    
        
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "WebRightBottomCorner"
        
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "BottomFlangeRightTop"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "BottomFlangeRight"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "BottomFlangeLeft"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "BottomFlangeLeftTop"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "WebLeftBottomCorner"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "WebLeftTopCorner"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "TopFlangeLeftBottomCorner"
    
'    nIndex = nIndex + 1
'    sEdgeNames(nIndex) = "TopFlangeLeftTopCorner"
    
    Dim sMemberName As String
    Dim sConstructMethodName As String
    Dim sConditionMethodName As String
    
    nCount = nIndex
    For nIndex = 4 To nCount
        sEdgeName = sEdgeNames(nIndex)
        sMemberName = "SlotIPC_" & sEdgeName
        sConstructMethodName = "CMConstruct" & sEdgeName & "PC"
        sConditionMethodName = "CMIs" & sEdgeName & "PCNeeded"

        Set oMemDesc = oMDs.AddMember(sMemberName, nIndex, sConstructMethodName, m_sRuleProgID)
        oMemDesc.SetCMConditional m_sRuleProgID, sConditionMethodName
        Set oMemDesc = Nothing
    Next
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_BottomLeftCF", nIndex, "CMConstructBottomLeftCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_BottomRightCF", nIndex, "CMConstructBottomRightCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_InsideLeftCF", nIndex, "CMConstructInsideLeftCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsInsideLeftCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_InsideRightCF", nIndex, "CMConstructInsideRightCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsInsideRightCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_OutsideLeftTopCF", nIndex, "CMConstructOutsideLeftTopCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsOutsideLeftTopCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_OutsideRightTopCF", nIndex, "CMConstructOutsideRightTopCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsOutsideRightTopCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_OutsideLeftBtmCF", nIndex, "CMConstructOutsideLeftBtmCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsOutsideLeftBtmCFNeeded"
    Set oMemDesc = Nothing
    
    nIndex = nIndex + 1
    Set oMemDesc = oMDs.AddMember("SlotI_OutsideRightBtmCF", nIndex, "CMConstructOutsideRightBtmCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsOutsideRightBtmCFNeeded"
    Set oMemDesc = Nothing
        
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructSlot(ByVal pAggregatorDescription As IJDAggregatorDescription)
  Dim clsSlotCM As New SlotDefCM
  clsSlotCM.Slot_FinalConstruct pAggregatorDescription
  Set clsSlotCM = Nothing
  
  Set m_oPhysConn = Nothing
End Sub
Public Sub CMConstructLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim clsSlotCM As New SlotDefCM
  Set m_oPhysConn = clsSlotCM.CAConstruct_PhysConn(pMemberDescription, pResourceManager, "TeeWeld", _
                                                    JXSEC_WEB_LEFT, JXSEC_WEB_LEFT)
  Set pObject = m_oPhysConn
  Set clsSlotCM = Nothing
End Sub
Public Sub CMConstructRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim clsSlotCM As New SlotDefCM
  Set m_oPhysConn = clsSlotCM.CAConstruct_PhysConn(pMemberDescription, pResourceManager, "TeeWeld", _
                                                   JXSEC_WEB_RIGHT, JXSEC_WEB_RIGHT)
  Set pObject = m_oPhysConn
  Set clsSlotCM = Nothing
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_sRuleProgID
End Function
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler

  Dim pDefinition As IJDSymbolDefinition
  Dim pCAFactory As New CAFactory
  
  Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_sRuleProgID
  pDefinition.CodeBase = CodeBase
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
    
  ' Initialize the definition
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "IJDUserSymbolServices_InstanciateDefinition").Number
End Function
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
  On Error GoTo ErrorHandler
  sMETHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
  sERROR = ""
  
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
  pIH.InitAs m_sFamilyProgID
  ItemInputs pIH
  
  Dim pPds As IJDPropertyDescriptions
  
  ' define the aggregator
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  Set pPds = pAD
  pPds.RemoveAll ' Remove all the previous property descriptions
  ItemAggregator pAD
   
  ' define the members
  Dim pMDs As IJDMemberDescriptions
  Set pMDs = pDefinition
  pMDs.RemoveAll ' Remove all the previous member descriptions
  ItemMembers pMDs
  
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

Public Sub CMIsTopPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
       If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) Then
                bIsNeeded = True
       End If
   End If

   Exit Sub
End Sub

Public Sub CMConstructTopPC(ByVal oMD As IJDMemberDescription, _
                            ByVal oResMgr As IUnknown, _
                            ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                    JXSEC_TOP, _
                                                    JXSEC_TOP)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
   
   Exit Sub
End Sub

Public Sub CMIsTopFlangeRightTopCornerPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If
      
      Set oSlotWrapper = Nothing
      
      If sSectionType = "B" Then
        If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER) Then
            bIsNeeded = True
        End If
      End If
   End If
   
   Exit Sub
End Sub

Public Sub CMConstructTopFlangeRightTopCornerPC(ByVal oMD As IJDMemberDescription, _
                                                ByVal oResMgr As IUnknown, _
                                                ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER, _
                                                     JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
   
   Exit Sub
End Sub

Public Sub CMIsTopFlangeRightPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If

      Set oSlotWrapper = Nothing
      
      If sSectionType = "EA" Or sSectionType = "UA" Or _
         sSectionType = "BUTL2" Or sSectionType = "BUT" Then
          If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT) Then
               bIsNeeded = True
          End If
      End If
   End If
   
   Exit Sub
End Sub
Public Sub CMConstructTopFlangeRightPC(ByVal oMD As IJDMemberDescription, _
                                       ByVal oResMgr As IUnknown, _
                                       ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_TOP_FLANGE_RIGHT, _
                                                     JXSEC_TOP_FLANGE_RIGHT)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
  
   Exit Sub
End Sub

Public Sub CMIsTopFlangeRightBottomCornerPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If

      Set oSlotWrapper = Nothing
      
      If sSectionType = "EA" Or sSectionType = "UA" Or _
         sSectionType = "B" Then
         If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER) Then
            bIsNeeded = True
         End If
      End If
   End If
   
   Exit Sub
End Sub

Public Sub CMConstructTopFlangeRightBottomCornerPC(ByVal oMD As IJDMemberDescription, _
                                                   ByVal oResMgr As IUnknown, _
                                                   ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER, _
                                                     JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
   
   Exit Sub
End Sub


Public Sub CMIsTopFlangeRightBottomPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If

      Set oSlotWrapper = Nothing
      
      If sSectionType = "EA" Or sSectionType = "UA" Or _
         sSectionType = "B" Or _
         sSectionType = "BUTL2" Or _
         sSectionType = "BUT" Then
        If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Then
                bIsNeeded = True
        End If
      End If
   End If
   
   Exit Sub
End Sub
Public Sub CMConstructTopFlangeRightBottomPC(ByVal oMD As IJDMemberDescription, _
                                                   ByVal oResMgr As IUnknown, _
                                                   ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_TOP_FLANGE_RIGHT_BOTTOM, _
                                                     JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
   
   Exit Sub
End Sub

Public Sub CMIsWebRightTopCornerPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If

      Set oSlotWrapper = Nothing
      
      If sSectionType = "EA" Or sSectionType = "UA" Or _
         sSectionType = "B" Then
        If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_RIGHT_TOP_CORNER) Then
                bIsNeeded = True
        End If
      End If
   End If
   
   Exit Sub
End Sub
Public Sub CMConstructWebRightTopCornerPC(ByVal oMD As IJDMemberDescription, _
                                          ByVal oResMgr As IUnknown, _
                                          ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_WEB_RIGHT_TOP_CORNER, _
                                                     JXSEC_WEB_RIGHT_TOP_CORNER)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
  
   Exit Sub
End Sub

Public Sub CMIsTopFlangeLeftBottomPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If

      Set oSlotWrapper = Nothing
      
      If sSectionType = "BUTL2" Or _
         sSectionType = "BUT" Then
        If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT_BOTTOM) Then
            bIsNeeded = True
        End If
      End If
   End If

   Exit Sub
End Sub
Public Sub CMConstructTopFlangeLeftBottomPC(ByVal oMD As IJDMemberDescription, _
                                            ByVal oResMgr As IUnknown, _
                                            ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_TOP_FLANGE_LEFT_BOTTOM, _
                                                     JXSEC_TOP_FLANGE_LEFT_BOTTOM)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
   
   Exit Sub
End Sub
Public Sub CMIsTopFlangeLeftPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   Set oSmartOccurrence = Nothing
   sItemName = oSmartItem.Name
   Set oSmartItem = Nothing
   
   If InStr(sItemName, "_LTR_") > 0 Then
      Dim oSlotWrapper As New StructDetailObjects.Slot
      Set oSlotWrapper.object = oMD.CAO
      
      'Get the section type
      Dim sSectionType As String
      
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        sSectionType = oProfile.sectionType
        Set oProfile = Nothing
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
      Else
        'Unsupported Penetrating Object
        Exit Sub
      End If

      Set oSlotWrapper = Nothing
      
      If sSectionType = "BUTL2" Or _
         sSectionType = "BUT" Then
         If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT) Then
            bIsNeeded = True
         End If
      End If
   End If

   Exit Sub
End Sub

Public Sub CMConstructTopFlangeLeftPC(ByVal oMD As IJDMemberDescription, _
                                            ByVal oResMgr As IUnknown, _
                                            ByRef oObject As Object)
   Dim oSlotDefCM As New SlotDefCM
  
   Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMD, oResMgr, "TeeWeld", _
                                                     JXSEC_TOP_FLANGE_LEFT, _
                                                     JXSEC_TOP_FLANGE_LEFT)
   Set oObject = m_oPhysConn
   Set oSlotDefCM = Nothing
   
   Exit Sub
End Sub

Public Sub CMIsLeftPCNeeded(ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
                           
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
                           
    bIsNeeded = True
  
    If Not DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) Then
             bIsNeeded = False
    End If


   Exit Sub
End Sub

Public Sub CMIsRightPCNeeded(ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
   
    bIsNeeded = True
    
    If Not DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_RIGHT) Then
             bIsNeeded = False
    End If
   

   Exit Sub
End Sub

Public Sub CMIsBottomCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
   If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
       bIsNeeded = False
       Exit Sub
   End If
                                                      
   bIsNeeded = False
   Dim strAnswer As String
   
   Dim edgeID As JXSEC_CODE
   
   If oMD.Dispid = 13 Then
    'Bottom Left Corner
    edgeID = JXSEC_WEB_LEFT
   ElseIf oMD.Dispid = 14 Then
    'Bottom Right Corner
    edgeID = JXSEC_WEB_RIGHT
   Else
    'Unknown Edge
    bIsNeeded = False
    Exit Sub
   End If
   
   GetSelectorAnswer oMD.CAO, "BaseCorners", strAnswer
   If LCase(strAnswer) = LCase("Yes") Then
        If DoesEdgeOverlapPart(oMD.CAO, edgeID) Then
            bIsNeeded = True
        End If
   End If
          
   Exit Sub
End Sub

Public Sub CMIsInsideLeftCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    Dim strAnswer As String
    
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMD.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If

    Set oSlotWrapper = Nothing
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
    
    If InStr(1, oSmItem.Name, "_LTR_", vbTextCompare) > 0 Then
        If sSectionType = "BUT" Or sSectionType = "BUTL2" Or sSectionType = "BUTL3" Then
             GetSelectorAnswer oMD.CAO, "InsideCorners", strAnswer
             If LCase(strAnswer) = LCase("Yes") Then
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT_BOTTOM) Then
                    bIsNeeded = True
                End If
             End If
        End If
    End If

   Exit Sub
End Sub

Public Sub CMIsInsideRightCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    Dim strAnswer As String
    
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMD.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
   
    If InStr(1, oSmItem.Name, "_LTR_", vbTextCompare) > 0 Then
        If sSectionType = "BUT" Or sSectionType = "BUTL2" Or sSectionType = "BUTL3" Then
             GetSelectorAnswer oMD.CAO, "InsideCorners", strAnswer
             If LCase(strAnswer) = LCase("Yes") Then
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_RIGHT) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Then
                    bIsNeeded = True
                End If
             End If
        End If
    End If
          
   Exit Sub
End Sub

Public Sub CMIsOutsideLeftTopCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
    
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMD.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim strAnswer As String
    If InStr(1, oSmItem.Name, "_LTR_", vbTextCompare) > 0 Then
        GetSelectorAnswer oMD.CAO, "OutsideCorners", strAnswer
        If LCase(strAnswer) = LCase("Yes") Then
            Select Case sSectionType
            Case "FB", "B", "EA", "UA", "BUTL3"
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) Then
                    bIsNeeded = True
                End If
            Case "BUT", "BUTL2"
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT) Then
                    bIsNeeded = True
                End If
            End Select
        End If
    End If
    
   Exit Sub
End Sub

Public Sub CMIsOutsideRightTopCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    Dim strAnswer As String
    
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMD.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
   
    If InStr(1, oSmItem.Name, "_LTR_", vbTextCompare) > 0 Then
        If sSectionType = "BUT" Or sSectionType = "BUTL2" Or sSectionType = "BUTL3" Or sSectionType = "FB" Then
             GetSelectorAnswer oMD.CAO, "OutsideCorners", strAnswer
             If LCase(strAnswer) = LCase("Yes") Then
                If sSectionType = "FB" Then
                    If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_RIGHT) Then
                        bIsNeeded = True
                    End If
                Else
                    If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT) Then
                        bIsNeeded = True
                    End If
                End If
             End If
        End If
    End If
          
   Exit Sub
End Sub

Public Sub CMIsOutsideLeftBtmCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    Dim strAnswer As String
    
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMD.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
    
    If InStr(1, oSmItem.Name, "_LTR_", vbTextCompare) > 0 Then
        If sSectionType = "BUT" Or sSectionType = "BUTL2" Or sSectionType = "BUTL3" Then
             GetSelectorAnswer oMD.CAO, "OutsideCorners", strAnswer
             If LCase(strAnswer) = LCase("Yes") Then
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT_BOTTOM) Then
                    bIsNeeded = True
                End If
             End If
        End If
    End If
     
   Exit Sub
End Sub

Public Sub CMIsOutsideRightBtmCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_SlotCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    Dim strAnswer As String
    
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMD.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
   
    If InStr(1, oSmItem.Name, "_LTR_", vbTextCompare) > 0 Then
        If sSectionType = "BUT" Or sSectionType = "BUTL2" Or sSectionType = "BUTL3" Then
             GetSelectorAnswer oMD.CAO, "OutsideCorners", strAnswer
             If LCase(strAnswer) = LCase("Yes") Then
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Then
                    bIsNeeded = True
                End If
             End If
        End If
    End If
    
   Exit Sub
End Sub

Public Sub CMConstructBottomLeftCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
    CreateBottomCF oMemberDescription, oResourceManager, JXSEC_WEB_LEFT, oObject
     
End Sub

Public Sub CMConstructBottomRightCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
    CreateBottomCF oMemberDescription, oResourceManager, JXSEC_WEB_RIGHT, oObject
End Sub

Public Sub CreateBottomCF(ByVal oMemberDescription As IJDMemberDescription, ByVal oResourceManager As IUnknown, eXID As JXSEC_CODE, ByRef oObject As Object)

    Dim oSlotWrapper As StructDetailObjects.Slot
    Dim oPenetratedStructConnectable As IJStructConnectable
    Dim oPenetratedPorts As IJElements
    Dim oPenetrationLoc As IJDPosition
    
    Set oSlotWrapper = New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMemberDescription.CAO
    
    Set oPenetratedStructConnectable = oSlotWrapper.Penetrated
    oPenetratedStructConnectable.enumConnectablePortsByOperation _
                            oPenetratedPorts, _
                            "CreatePlatePart.GeneratePlatePart_AE.1", _
                            PortFace, False, False
    Set oPenetrationLoc = oSlotWrapper.PenetrationLocation
    
    Dim oModelHelper As IJSGOModelBodyUtilities
    Dim oClosestPoint As IJDPosition
    Dim dDistance As Double
    Dim dMinDistance As Double
    Dim dDistToFace As Double
    Dim dMinDistToFace As Double
    
    Set oModelHelper = New SGOModelBodyUtilities
    dMinDistance = 100
    dMinDistToFace = 100
    
    Dim nPortIndex As Long
    Dim oStructPort As IJStructPort
    Dim oFacePort As IJPort
    Dim oEdgePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    Dim oPortMB As IJModelBody
    Dim nCtx As Long
    Dim oTempPort As IJPort
    
    Dim oBasePortColl As New Collection
    For nPortIndex = 1 To oPenetratedPorts.Count
       Set oStructPort = oPenetratedPorts.Item(nPortIndex)
       nCtx = oStructPort.ContextID
       
       If (nCtx And CTX_BASE) > 0 Then
            Set oTempPort = oPenetratedPorts.Item(nPortIndex)
            Set oPortMB = oTempPort.Geometry

            oModelHelper.GetClosestPointOnBody _
                                        oPortMB, _
                                        oPenetrationLoc, _
                                        oClosestPoint, _
                                        dDistToFace

            If dDistToFace < dMinDistToFace Then
                dMinDistToFace = dDistToFace
                Set oFacePort = oTempPort
            ElseIf oFacePort Is Nothing Then
                Set oFacePort = oTempPort
            End If
       ElseIf (nCtx And CTX_OFFSET) > 0 Then

       ElseIf (nCtx And CTX_LATERAL_LFACE) > 0 And (nCtx And CTX_3DCUTOUT) = 0 Then
          Set oEdgePort = oPenetratedPorts.Item(nPortIndex)
          Set oPortMB = oEdgePort.Geometry

          oModelHelper.GetClosestPointOnBody _
                                     oPortMB, _
                                     oPenetrationLoc, _
                                     oClosestPoint, _
                                     dDistance

          If dDistance <= dMinDistance Then
             If dDistance = dMinDistance Then
                oBasePortColl.Add oEdgePort
             Else
                Set oBasePortColl = New Collection
                oBasePortColl.Add oEdgePort
             End If

             dMinDistance = dDistance
             'Set oEdgePort2 = oEdgePort
          End If
       End If
    Next

    'If Penetrating Part Is a Plate there may be multiple ports at 0 distance
    Dim oPort As IJPort
    Dim oSurface As IJSurfaceBody
    Dim oPortNormal As New DVector
    Dim dDotProduct As Double
    Dim dMaxDotProduct As Double
    Dim nIndex As Integer
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate

    
    If TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        If oBasePortColl.Count = 1 Then
            Set oEdgePort2 = oBasePortColl.Item(1)
        ElseIf oBasePortColl.Count > 1 Then
            Dim oSlotMappingRule As IJSlotMappingRule
            Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance

            Dim U As IJDVector
            Dim V As IJDVector
            oSlotMappingRule.GetSketchingPlane oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, oPenetrationLoc, U, V
            
            dMaxDotProduct = 0
            V.Length = 1

            For nIndex = 1 To oBasePortColl.Count
                Set oPort = oBasePortColl.Item(nIndex)
                Set oSurface = oPort.Geometry
                oSurface.GetNormalFromPosition oPenetrationLoc, oPortNormal
                oPortNormal.Length = 1
                dDotProduct = Abs(V.Dot(oPortNormal))
                If dDotProduct > dMaxDotProduct Then
                    dMaxDotProduct = dDotProduct
                    Set oEdgePort2 = oBasePortColl.Item(nIndex)
                End If
            Next
        End If
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        If oBasePortColl.Count = 1 Then
            Set oEdgePort2 = oBasePortColl.Item(1)
        ElseIf oBasePortColl.Count > 1 Then
            Dim oProfile As New StructDetailObjects.ProfilePart
            Set oProfile.object = oSlotWrapper.Penetrating
            
            Dim oMountingPort As IJPort
            Dim oMountingPortSurface As IJSurfaceBody
            Dim oTempSurface As IJSurface
            Set oMountingPort = oProfile.MountingFacePort()
            Set oMountingPortSurface = oMountingPort.Geometry
            
            Dim oProjectedPoint As IJDPosition
            Dim oOutputNormal As IJDVector
            Dim oMountingNormal As New DVector
            oTopologyLocate.GetProjectedPointOnModelBody oMountingPortSurface, oPenetrationLoc, oProjectedPoint, oOutputNormal
            
            oMountingPortSurface.GetNormalFromPosition oProjectedPoint, oMountingNormal
            oMountingNormal.Length = 1
            dMaxDotProduct = 0
            
            For nIndex = 1 To oBasePortColl.Count
                Set oPort = oBasePortColl.Item(nIndex)
                Set oSurface = oPort.Geometry
                
                oTopologyLocate.GetProjectedPointOnModelBody oSurface, oPenetrationLoc, oProjectedPoint, oOutputNormal
                
                oSurface.GetNormalFromPosition oProjectedPoint, oPortNormal

                oPortNormal.Length = 1

                dDotProduct = Abs(oMountingNormal.Dot(oPortNormal))

                If dDotProduct > dMaxDotProduct Then
                    dMaxDotProduct = dDotProduct
                    Set oEdgePort2 = oBasePortColl.Item(nIndex)
                End If
            Next
        End If
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim nXid As Integer
    Dim oSegment As Object
    Dim eFeatureType As StructFeatureTypes
    
    nXid = eXID
    Set oStructFeature = oMemberDescription.CAO
    If Not oStructFeature Is Nothing Then
       eFeatureType = oStructFeature.get_StructFeatureType
       oStructFeature.GetContourSegment nXid, oSegment
    End If
    
    Dim sSelString As String
    sSelString = "RootCorner"
    
    Dim oSDOCornerFeature As IJSDOCornerFeature
    Set oSDOCornerFeature = New StructDetailObjectsEx.CornerFeature
    oSDOCornerFeature.Create oResourceManager, _
                             oFacePort, _
                              oSegment, _
                             oEdgePort2, _
                             sSelString, _
                             oMemberDescription.CAO
                             
    Set oObject = oSDOCornerFeature.object
  Exit Sub
  
End Sub

Public Sub CMConstructInsideLeftCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
    CreateCFFromSlotEdges oMemberDescription, oResourceManager, JXSEC_TOP_FLANGE_LEFT_BOTTOM, JXSEC_WEB_LEFT, "RootCorner", oObject
     
End Sub

Public Sub CMConstructInsideRightCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
    CreateCFFromSlotEdges oMemberDescription, oResourceManager, JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_WEB_RIGHT, "RootCorner", oObject
End Sub
Public Sub CMConstructOutsideLeftTopCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
                  
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMemberDescription.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim nXId2 As Long
    
    If sSectionType = "B" Or sSectionType = "EA" Or sSectionType = "UA" Or sSectionType = "FB" Then
        nXId2 = JXSEC_WEB_LEFT
    ElseIf sSectionType = "BUT" Or sSectionType = "BUTL2" Then
        nXId2 = JXSEC_TOP_FLANGE_LEFT
    End If

    CreateCFFromSlotEdges oMemberDescription, oResourceManager, JXSEC_TOP, nXId2, "VariableEdgeCorner", oObject
     
End Sub

Public Sub CMConstructOutsideRightTopCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                           ByRef oObject As Object)
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMemberDescription.CAO
    
    'Get the section type
    Dim sSectionType As String
    
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
      Dim oProfile As New StructDetailObjects.ProfilePart
      Set oProfile.object = oSlotWrapper.Penetrating
      sSectionType = oProfile.sectionType
      Set oProfile = Nothing
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
      Dim oSlotMappingRule As IJSlotMappingRule
      Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
      
      Dim oWeb As Object
      Dim oFlange As Object
      Dim o2ndWeb As Object
      Dim o2ndFlange As Object
      oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
      'Unsupported Penetrating Object
      Exit Sub
    End If
    
    Dim nXId2 As JXSEC_CODE
    
    If sSectionType = "FB" Then
        nXId2 = JXSEC_WEB_RIGHT
    ElseIf sSectionType = "BUT" Or sSectionType = "BUTL2" Then
        nXId2 = JXSEC_TOP_FLANGE_RIGHT
    End If
    
    CreateCFFromSlotEdges oMemberDescription, oResourceManager, JXSEC_TOP, nXId2, "VariableEdgeCorner", oObject
End Sub

Public Sub CMConstructOutsideLeftBtmCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
    CreateCFFromSlotEdges oMemberDescription, oResourceManager, JXSEC_TOP_FLANGE_LEFT_BOTTOM, JXSEC_TOP_FLANGE_LEFT, "VariableEdgeCorner", oObject
     
End Sub

Public Sub CMConstructOutsideRightBtmCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
    CreateCFFromSlotEdges oMemberDescription, oResourceManager, JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_TOP_FLANGE_RIGHT, "VariableEdgeCorner", oObject
End Sub


Public Sub CreateCFFromSlotEdges(oMemberDescription As IJDMemberDescription, ByVal oResourceManager As IUnknown, nXID1 As JXSEC_CODE, nXId2 As JXSEC_CODE, sSelString As String, ByRef oObject As Object)

    
    Dim oSlotWrapper As StructDetailObjects.Slot
    Dim oPenetratedStructConnectable As IJStructConnectable
    Dim oPenetratedPorts As IJElements
    Dim oPenetrationLoc As IJDPosition
    
    Set oSlotWrapper = New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMemberDescription.CAO
    
    Set oPenetratedStructConnectable = oSlotWrapper.Penetrated
    oPenetratedStructConnectable.enumConnectablePortsByOperation _
                            oPenetratedPorts, _
                            "CreatePlatePart.GeneratePlatePart_AE.1", _
                            PortFace, False, False
    Set oPenetrationLoc = oSlotWrapper.PenetrationLocation
    
    Dim oModelHelper As IJSGOModelBodyUtilities
    Dim oClosestPoint As IJDPosition
    Dim dDistance As Double
    Dim dMinDistance As Double
    Dim dDistToFace As Double
    Dim dMinDistToFace As Double
    
    Set oModelHelper = New SGOModelBodyUtilities
    dMinDistance = 100
    dMinDistToFace = 100
    
    Dim nPortIndex As Long
    Dim oStructPort As IJStructPort
    Dim oFacePort As IJPort
    Dim oEdgePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    Dim oPortMB As IJModelBody
    Dim nCtx As Long
    Dim oTempPort As IJPort
    
    For nPortIndex = 1 To oPenetratedPorts.Count
       Set oStructPort = oPenetratedPorts.Item(nPortIndex)
       nCtx = oStructPort.ContextID
       
       If (nCtx And CTX_BASE) > 0 Then
            Set oTempPort = oPenetratedPorts.Item(nPortIndex)
            Set oPortMB = oTempPort.Geometry

            oModelHelper.GetClosestPointOnBody _
                                        oPortMB, _
                                        oPenetrationLoc, _
                                        oClosestPoint, _
                                        dDistToFace
            If dDistToFace < dMinDistToFace Then
                dMinDistToFace = dDistToFace
                Set oFacePort = oTempPort
            ElseIf oFacePort Is Nothing Then
                Set oFacePort = oTempPort
            End If
       ElseIf (nCtx And CTX_OFFSET) > 0 Then
          
       ElseIf (nCtx And CTX_LATERAL_LFACE) > 0 Then
       
       End If
    Next
    
    Dim oStructFeature As IJStructFeature
    Dim nXid As Integer
    Dim oSupport1 As Object
    Dim oSupport2 As Object
    Dim eFeatureType As StructFeatureTypes
    
    Set oStructFeature = oMemberDescription.CAO
    
    If Not oStructFeature Is Nothing Then
        oStructFeature.GetContourSegment nXId2, oSupport2
        oStructFeature.GetContourSegment nXID1, oSupport1
    End If
    
    Dim oSDOCornerFeature As IJSDOCornerFeature
    
    Set oSDOCornerFeature = New StructDetailObjectsEx.CornerFeature
    oSDOCornerFeature.Create _
                      oResourceManager, _
                      oFacePort, _
                      oSupport1, _
                      oSupport2, _
                       sSelString, _
                      oMemberDescription.CAO
    Set oObject = oSDOCornerFeature.object
    

Exit Sub
End Sub
