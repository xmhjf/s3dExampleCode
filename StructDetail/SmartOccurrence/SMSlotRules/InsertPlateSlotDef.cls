VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsertPlateSlotDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'pnalugol - Nov1 2011 - Changes done to support 202656, added
'       new item members to create Corner features as child of Slots

Option Explicit

Implements IJDUserSymbolServices

' Modify class name to particular rule
Private Const m_sClassName As String = "InsertPlateSlotDef"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"

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
  
  pAD.SetCMFinalConstruct CUSTOMERID + "SlotRules.InsertPlateSlotDef", "CMFinalConstructSlot"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
  On Error GoTo ErrorHandler

    Dim oMemDesc As IJDMemberDescription
    
    'Insert Plate Web Left connection
    Set oMemDesc = oMDs.AddMember("InsertPlate_Left_PC", 1, "CMConstructLeftPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsLeftPCNeeded"
    Set oMemDesc = Nothing
    
    'Insert Plate Web Right connection
    Set oMemDesc = oMDs.AddMember("InsertPlate_Right_PC", 2, "CMConstructRightPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsRightPCNeeded"
    Set oMemDesc = Nothing

    'Insert Plate Top connection
    Set oMemDesc = oMDs.AddMember("InsertPlate_Top_PC", 3, "CMConstructTopPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopPCNeeded"
    Set oMemDesc = Nothing

    'Insert Plate Bottom connection
    Set oMemDesc = oMDs.AddMember("InsertPlate_Bottom_PC", 4, "CMConstructBottomPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomPCNeeded"
    Set oMemDesc = Nothing
    
    'Insert Plate Left Chamfer connection
    Set oMemDesc = oMDs.AddMember("InsertPlate_LeftChamfer_PC", 5, "CMConstructLeftChamferPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsLeftChamferPCNeeded"
    Set oMemDesc = Nothing
    
    'Insert Plate Right Chamfer connection
    Set oMemDesc = oMDs.AddMember("InsertPlate_RightChamfer_PC", 6, "CMConstructRightChamferPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsRightChamferPCNeeded"
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
On Error GoTo ErrorHandler

    sMETHOD = "Slot_FinalConstruct"
    sERROR = "Constructing Insert Plate Slot"
    
    Dim clsSlotCM As New SlotDefCM
    clsSlotCM.Slot_FinalConstruct pAggregatorDescription
    Set clsSlotCM = Nothing
    Set m_oPhysConn = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Numbert
    
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

Public Sub CMConstructLeftPC(ByVal oMemberDescription As IJDMemberDescription, _
                         ByVal oResourceManager As IUnknown, _
                         ByRef oObject As Object)
                         
On Error GoTo ErrorHandler

    sMETHOD = "CMConstructLeftPC"
    sERROR = "Creating PC on the Left side of Insert Plate"
    Dim oSlotDefCM As New SlotDefCM
    
    'pMemberDescription.
    Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMemberDescription, _
                                                    oResourceManager, _
                                                    "ButtWeld", _
                                                    JXSEC_WEB_LEFT, _
                                                    JXSEC_WEB_LEFT)
    Set oObject = m_oPhysConn
    Set oSlotDefCM = Nothing
    Set m_oPhysConn = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMIsLeftPCNeeded(ByRef oMD As IJDMemberDescription, _
                            ByRef bIsNeeded As Boolean)
On Error GoTo ErrorHandler

    sMETHOD = "CMIsLeftPCNeeded"
    sERROR = "Conditional for InsertPlate Left PC"
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
   
    bIsNeeded = True
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMIsRightPCNeeded(ByRef oMD As IJDMemberDescription, _
                            ByRef bIsNeeded As Boolean)
On Error GoTo ErrorHandler

    sMETHOD = "CMIsRightPCNeeded"
    sERROR = "Conditional for InsertPlate Right PC"
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMIsTopPCNeeded(ByRef oMD As IJDMemberDescription, _
                            ByRef bIsNeeded As Boolean)
  On Error GoTo ErrorHandler

    sMETHOD = "CMIsTopPCNeeded"
    sERROR = "Conditional for InsertPlate Top PC"
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    
    Dim oSD_Slot As StructDetailObjects.Slot
    Set oSD_Slot = New StructDetailObjects.Slot
    Set oSD_Slot.object = oMD.CAO
    
    Dim bIsPlateFlangePenetrated As Boolean
    'Check if insert plate is Flange Penetrated
    bIsPlateFlangePenetrated = IsPlateThroughMbrFlange(oSD_Slot.Penetrating, oSD_Slot.Penetrated)
    
    'For Flange penetrated check the plate position and C/S
    If bIsPlateFlangePenetrated Then
        
        Dim oMappedPorts As JCmnShp_CollectionAlias
        Dim oGeomOpr As New GSCADShipGeomOps.SGOModelBodyUtilities
        Dim oSlotMappingRule As IJSlotMappingRule
        Dim pSDOHelper As New StructDetailObjects.Helper
        Dim oBasePort As IJPort
        Dim oPlateport As IJPort
        Dim bIntersect As Boolean
        Dim strSectionType As String
        Dim oSDMemberPart As New StructDetailObjects.MemberPart
        Set oSDMemberPart.object = oSD_Slot.Penetrated
        
        Set oMappedPorts = New Collection
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
         
        'Get Mapped Ports and Section Type
        oSlotMappingRule.GetEmulatedPorts oSD_Slot.Penetrating, oSD_Slot.Penetrated, oBasePort, oMappedPorts
        Set oPlateport = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item(CStr(JXSEC_TOP)))
        strSectionType = oSDMemberPart.sectionType
        
        'For C and L C/S Need to check more to created PC
        If UCase(strSectionType) = "C" Or UCase(strSectionType) = "L" Then
        
            Dim oMbrPort As IJPort
            Set oMbrPort = GetLateralSubPortBeforeTrim(oSDMemberPart.object, JXSEC_WEB_LEFT)
            'If there is Intersection b/w Mapped platePort and Mbr WebLeft then no need to created PC.
            If oGeomOpr.HasIntersectingGeometry(oMbrPort.Geometry, oPlateport.Geometry) Then
                bIsNeeded = False
                Exit Sub
            End If
        End If
 
        'Need to Create PC if Mapped Port is intersecting Mbr Lateral Port
        If oGeomOpr.HasIntersectingGeometry(oSDMemberPart.BasePortBeforeTrim(BPT_Lateral).Geometry, oPlateport.Geometry) Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    Else ' IF not Flange Penertrated, then PC is not needed
        bIsNeeded = True
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMIsBottomPCNeeded(ByRef oMD As IJDMemberDescription, _
                            ByRef bIsNeeded As Boolean)
On Error GoTo ErrorHandler

    sMETHOD = "CMIsBottomPCNeeded"
    sERROR = "Conditional for InsertPlate Bottom PC"
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oSD_Slot As New StructDetailObjects.Slot
    Set oSD_Slot.object = oMD.CAO
    
    bIsNeeded = False
    
    Dim bIsPlateFlangePenetrated As Boolean
    'Check if insert plate is Flange Penetrated
    bIsPlateFlangePenetrated = IsPlateThroughMbrFlange(oSD_Slot.Penetrating, oSD_Slot.Penetrated)
    
    'For Flange penetrated check the plate position and C/S
    If bIsPlateFlangePenetrated Then
        
        Dim oSDMemberPart As New StructDetailObjects.MemberPart
        Set oSDMemberPart.object = oSD_Slot.Penetrated
            
        Dim oMappedPorts As JCmnShp_CollectionAlias
        Dim oGeomOpr As New GSCADShipGeomOps.SGOModelBodyUtilities
        Dim pSDOHelper As New StructDetailObjects.Helper
        Dim oSlotMappingRule As IJSlotMappingRule
        Dim oBasePort As IJPort
        Dim oPlateport As IJPort
        Dim bIntersect As Boolean
        Dim strSectionType As String
        
        Set oMappedPorts = New Collection
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        'Get Mapped Ports
        oSlotMappingRule.GetEmulatedPorts oSD_Slot.Penetrating, oSD_Slot.Penetrated, oBasePort, oMappedPorts
        Set oPlateport = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item(CStr(JXSEC_BOTTOM)))
 
        'Get Member Section Type
        strSectionType = oSDMemberPart.sectionType
        
        'For C and L C/S Need more checks to created PC
        If UCase(strSectionType) = "C" Or UCase(strSectionType) = "L" Then
        
            Dim oMbrPort As IJPort
            Set oMbrPort = GetLateralSubPortBeforeTrim(oSDMemberPart.object, JXSEC_WEB_LEFT)
            
            'If there is Intersection b/w Mapped platePort and Mbr WebLeft then no need to created PC.
            If oGeomOpr.HasIntersectingGeometry(oMbrPort.Geometry, oPlateport.Geometry) Then
                bIsNeeded = False
                Exit Sub
            End If
        End If
        
        'Need to Create PC if Mapped Port is intersecting Mbr Lateral Port
        If oGeomOpr.HasIntersectingGeometry(oSDMemberPart.BasePortBeforeTrim(BPT_Lateral).Geometry, oPlateport.Geometry) Then
            bIsNeeded = True
        End If
        
    Else ' IF not Flange Penertrated, then PC is not needed
        bIsNeeded = False
    End If
    
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
Public Sub CMConstructRightPC(ByVal oMemberDescription As IJDMemberDescription, _
                         ByVal oResourceManager As IUnknown, _
                         ByRef oObject As Object)
On Error GoTo ErrorHandler

    sMETHOD = "CMConstructRightPC"
    sERROR = "Constructs InsertPlate Right PC"
    
    Dim oSlotDefCM As New SlotDefCM
    
    'pMemberDescription.
    Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMemberDescription, _
                                                      oResourceManager, _
                                                      "ButtWeld", _
                                                      JXSEC_WEB_RIGHT, _
                                                      JXSEC_WEB_RIGHT)
    Set oObject = m_oPhysConn
    Set oSlotDefCM = Nothing
    Set m_oPhysConn = Nothing
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
Public Sub CMConstructTopPC(ByVal oMemberDescription As IJDMemberDescription, _
                         ByVal oResourceManager As IUnknown, _
                         ByRef oObject As Object)
On Error GoTo ErrorHandler

    sMETHOD = "CMConstructTopPC"
    sERROR = "Constructs InsertPlate Top PC"
    
    Dim oSlotDefCM As New SlotDefCM
    
    'pMemberDescription.
    Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMemberDescription, _
                                                      oResourceManager, _
                                                      "TeeWeld", _
                                                       JXSEC_TOP, _
                                                      JXSEC_TOP)
    Set oObject = m_oPhysConn
    Set oSlotDefCM = Nothing
    Set m_oPhysConn = Nothing
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
Public Sub CMConstructBottomPC(ByVal oMemberDescription As IJDMemberDescription, _
                         ByVal oResourceManager As IUnknown, _
                         ByRef oObject As Object)
On Error GoTo ErrorHandler

    sMETHOD = "CMConstructBottomPC"
    sERROR = "Constructs InsertPlate Bottom PC"
    Dim oSlotDefCM As New SlotDefCM
    
    'pMemberDescription.
    Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMemberDescription, _
                                                      oResourceManager, _
                                                      "TeeWeld", _
                                                      JXSEC_BOTTOM, _
                                                      JXSEC_BOTTOM)
    Set oObject = m_oPhysConn
    Set oSlotDefCM = Nothing
    Set m_oPhysConn = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMIsLeftChamferPCNeeded(ByRef oMD As IJDMemberDescription, _
                                    ByRef bIsNeeded As Boolean)
On Error GoTo ErrorHandler

    sMETHOD = "CMIsLeftChamferPCNeeded"
    sERROR = "Conditional for InsertPlate Left Chamfer PC"
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
   
    Dim oSmartItemName As IJSmartItem
    Dim oSmartOcc As IJSmartOccurrence
    
    bIsNeeded = False
    
    'Get Parent Smart Occurance
    GetSmartOccurrenceParent oMD.CAO, oSmartOcc
    Set oSmartItemName = oSmartOcc.SmartItemObject
    
    'Check if Parent Smart Item is Chamfer AC
    If oSmartItemName.Name = "Chamfered_InsertPlate_AC" Then
        bIsNeeded = True
    End If
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMIsRightChamferPCNeeded(ByRef oMD As IJDMemberDescription, _
                                    ByRef bIsNeeded As Boolean)
On Error GoTo ErrorHandler

    sMETHOD = "CMIsRightChamferPCNeeded"
    sERROR = "Conditional for InsertPlate Right Chamfer PC"
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
   
    Dim oSmartItemName As IJSmartItem
    Dim oSmartOcc As IJSmartOccurrence
    
    bIsNeeded = False
    
    'Get Parent Smart Occrance
    GetSmartOccurrenceParent oMD.CAO, oSmartOcc
    Set oSmartItemName = oSmartOcc.SmartItemObject
    
    'Check if Parent Smart Item is Chamfer AC
    If oSmartItemName.Name = "Chamfered_InsertPlate_AC" Then
        bIsNeeded = True
    End If
   
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
Public Sub CMConstructLeftChamferPC(ByVal oMemberDescription As IJDMemberDescription, _
                                    ByVal oResourceManager As IUnknown, _
                                    ByRef oObject As Object)
On Error GoTo ErrorHandler

    sMETHOD = "CMConstructLeftChamferPC"
    sERROR = "Constructs InsertPlate Left Chamfer PC"
    
    Dim oSlotDefCM As New SlotDefCM
    
    Dim oSD_Slot As New StructDetailObjects.Slot
    Set oSD_Slot.object = oMemberDescription.CAO

    Dim oPenetratedPort As IJPort
    Dim oPenetratingPort As IJPort
    Dim oTopology As New TopologyLocate
    Dim oMbrPartWrapper As New StructDetailObjectsEx.SDOMemberPart
    
    Set oMbrPartWrapper.object = oSD_Slot.Penetrated
    Set oPenetratedPort = oMbrPartWrapper.CutoutPort(oSD_Slot.object, JXSEC_WEB_LEFT_TOP_CORNER)
    
    Dim oSD_Platepart As New StructDetailObjects.PlatePart
    Set oSD_Platepart.object = oSD_Slot.Penetrating
    
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance()
    
    Dim oBasePort As IJPort
    Dim oMappedPorts As JCmnShp_CollectionAlias
    Set oMappedPorts = New Collection
    
    oSlotMappingRule.GetEmulatedPorts oSD_Slot.Penetrating, oSD_Slot.Penetrated, oBasePort, oMappedPorts
    
    Dim oStructPort As IJStructPort
    Dim ePortType As Base_Port_Types
    
    Set oStructPort = oMappedPorts.Item(CStr(JXSEC_TOP))
    
    If oStructPort.ContextID And CTX_BASE Then
        ePortType = BPT_Base
    ElseIf oStructPort.ContextID And CTX_OFFSET Then
        ePortType = BPT_Offset
    Else
        'Error case!!! unknown case
        'currently setting it to LATERAL
        ePortType = BPT_Lateral
    End If
    
    Set oPenetratingPort = oTopology.GetEquivalentLastPort(oSD_Platepart.BasePort(ePortType))

    Dim pSystemParent As IJSystemChild
    Set pSystemParent = oMemberDescription.CAO
    
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    oPhysConn.Create oResourceManager, oPenetratedPort, oPenetratingPort, "TeeWeld", pSystemParent, ConnectionStandard
    
    Set oObject = oPhysConn.object
    Set oSlotDefCM = Nothing
    Set oPhysConn = Nothing
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMConstructRightChamferPC(ByVal oMemberDescription As IJDMemberDescription, _
                                    ByVal oResourceManager As IUnknown, _
                                    ByRef oObject As Object)
On Error GoTo ErrorHandler

    sMETHOD = "CMConstructRightChamferPC"
    sERROR = "Constructs InsertPlate Right Chamfer PC"
    
    Dim oSlotDefCM As New SlotDefCM
    
    Dim oSD_Slot As New StructDetailObjects.Slot
    Set oSD_Slot.object = oMemberDescription.CAO

    Dim oPenetratedPort As IJPort
    Dim oPenetratingPort As IJPort
    Dim oTopology As New TopologyLocate
    Dim oMbrPartWrapper As New StructDetailObjectsEx.SDOMemberPart
    
    Set oMbrPartWrapper.object = oSD_Slot.Penetrated
    Set oPenetratedPort = oMbrPartWrapper.CutoutPort(oSD_Slot.object, JXSEC_WEB_RIGHT_TOP_CORNER)
    
    Dim oSD_Platepart As New StructDetailObjects.PlatePart
    Set oSD_Platepart.object = oSD_Slot.Penetrating
    
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance()
    
    Dim oBasePort As IJPort
    Dim oMappedPorts As JCmnShp_CollectionAlias
    Set oMappedPorts = New Collection
    
    oSlotMappingRule.GetEmulatedPorts oSD_Slot.Penetrating, oSD_Slot.Penetrated, oBasePort, oMappedPorts
    
    Dim oStructPort As IJStructPort
    Dim ePortType As Base_Port_Types
    
    Set oStructPort = oMappedPorts.Item(CStr(JXSEC_TOP))
    
    If oStructPort.ContextID And CTX_BASE Then
        ePortType = BPT_Base
    ElseIf oStructPort.ContextID And CTX_OFFSET Then
        ePortType = BPT_Offset
    Else
        'Error case!!! unknown case
        'currently setting it to LATERAL
        ePortType = BPT_Lateral
    End If
    
    Set oPenetratingPort = oTopology.GetEquivalentLastPort(oSD_Platepart.BasePort(ePortType))

    Dim pSystemParent As IJSystemChild
    Set pSystemParent = oMemberDescription.CAO
    
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    oPhysConn.Create oResourceManager, oPenetratedPort, oPenetratingPort, "TeeWeld", pSystemParent, ConnectionStandard
    
    Set oObject = oPhysConn.object
    Set oSlotDefCM = Nothing
    Set oPhysConn = Nothing
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

'------------------------------------------------------------------------------------------------------------
' METHOD:  IsPlateThroughMbrFlange
'
' DESCRIPTION:  Checks if Plate is penetrating through the Mbr and returns True if Plate is pentrating From top to Bottom
'               of Member Part
'
' Inputs : Plate Part and Member Part
'
' Output :Boolean (T/F)
'
'------------------------------------------------------------------------------------------------------------

Private Function IsPlateThroughMbrFlange(ByVal oPlatePart As Object, ByVal oMbrPart As Object) As Boolean

  On Error GoTo ErrorHandler

  sMETHOD = "IsPlateThroughMbrFlange"

    Dim oMbrTopPort As IJPort
    Dim oMbrBottomPort As IJPort
    Dim bHasTopPortIntersection As Boolean
    Dim bHasBtmPortIntersection As Boolean
    Dim oSGOModelBodyUtils As New GSCADShipGeomOps.SGOModelBodyUtilities
    
    IsPlateThroughMbrFlange = False

    'Get Mbr Top and Bottom ports
    Set oMbrTopPort = GetLateralSubPortBeforeTrim(oMbrPart, JXSEC_TOP)
    Set oMbrBottomPort = GetLateralSubPortBeforeTrim(oMbrPart, JXSEC_BOTTOM)
    
    'Check if Plate intersects Member Top
    bHasTopPortIntersection = oSGOModelBodyUtils.HasIntersectingGeometry(oMbrTopPort.Geometry, oPlatePart)

    'If Plate is not intersecting Mbr Top, then Exit function as not flange penetrated
    If Not bHasTopPortIntersection Then Exit Function
    
    'Check if Plate intersects Member Bottom
    bHasBtmPortIntersection = oSGOModelBodyUtils.HasIntersectingGeometry(oMbrBottomPort.Geometry, oPlatePart)

    If bHasBtmPortIntersection Then
        IsPlateThroughMbrFlange = True
    End If

  Exit Function
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

