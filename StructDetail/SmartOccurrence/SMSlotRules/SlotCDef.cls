VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'pnalugol - Nov1 2011 - Changes done to support 202656, added
'       new item members to create Corner features as child of Slots

Option Explicit

Implements IJDUserSymbolServices

' Modify class name to particular rule
Private Const m_sClassName As String = "SlotCDef"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"
Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
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
  
  pAD.SetCMFinalConstruct CUSTOMERID + "SlotRules.SlotCDef", "CMFinalConstructSlot"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(oMDs As IJDMemberDescriptions)
  On Error GoTo ErrorHandler
    Dim oMemDesc As IJDMemberDescription

    ' By default,it's Web Left connection
    Set oMemDesc = oMDs.AddMember("SlotCPC1", 1, "CMConstructPC", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsLeftPCNeeded"
    
    Set oMemDesc = oMDs.AddMember("SlotCFET1", 2, "CMConstructFET1", CMLIBRARY_SLOTDEFCM)
    oMemDesc.SetCMConditional CMLIBRARY_SLOTDEFCM, "CMGetEdgeTreatment"

    Set oMemDesc = oMDs.AddMember("SlotC_PC_Top", 3, "CMConstructPC_Top", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopPCNeeded"
    
    Set oMemDesc = oMDs.AddMember("SlotC_BottomLeftCF", 4, "CMConstructBottomLeftCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsBottomLeftCFNeeded"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("SlotC_TopLeftCF", 5, "CMConstructTopLeftCF", m_sRuleProgID)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsTopLeftCFNeeded"
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
Public Sub CMConstructPC(ByVal oMemberDescription As IJDMemberDescription, _
                         ByVal oResourceManager As IUnknown, _
                         ByRef oObject As Object)
  Dim oSlotDefCM As New SlotDefCM
  
  'pMemberDescription.
  Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMemberDescription, _
                                                    oResourceManager, _
                                                    "TeeWeld", _
                                                    JXSEC_WEB_LEFT, _
                                                    JXSEC_WEB_LEFT)
  Set oObject = m_oPhysConn
  Set oSlotDefCM = Nothing
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

Public Sub CMIsLeftPCNeeded(ByRef oMD As IJDMemberDescription, _
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
   
   If InStr(sItemName, "_L_") > 0 Or _
      InStr(sItemName, "_LT_") Then
  
       If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) Then
                bIsNeeded = True
       End If
       
   End If
   
   Exit Sub
End Sub

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
   
   If InStr(sItemName, "_T_") > 0 Or _
      InStr(sItemName, "_LT_") Then
        If DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) Then
                bIsNeeded = True
        End If
   End If
   
   Exit Sub
End Sub

Public Sub CMConstructPC_Top(ByVal oMemberDescription As IJDMemberDescription, _
                             ByVal oResourceManager As IUnknown, _
                             ByRef oObject As Object)
  'pMemberDescription.
  Dim oSlotDefCM As New SlotDefCM
  
  Set m_oPhysConn = oSlotDefCM.CAConstruct_PhysConn(oMemberDescription, _
                                                    oResourceManager, _
                                                    "TeeWeld", _
                                                    JXSEC_TOP, _
                                                    JXSEC_TOP)
  Set oObject = m_oPhysConn
  Set oSlotDefCM = Nothing
End Sub

Public Sub CMIsBottomLeftCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
                           
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
                               
    bIsNeeded = False
    Dim strAnswer As String
    
    Dim oSMOcc As IJSmartOccurrence
    Set oSMOcc = oMD.CAO
    
    Dim oSmItem As IJSmartItem
    Set oSmItem = oSMOcc.SmartItemObject
 
    If InStr(1, oSmItem.Name, "_L_", vbTextCompare) > 0 Or InStr(1, oSmItem.Name, "_LT_", vbTextCompare) > 0 Then
        GetSelectorAnswer oMD.CAO, "BaseCorners", strAnswer

        If LCase(strAnswer) = LCase("Yes") Then
            If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) Then
                bIsNeeded = True
            End If
        End If
    End If
                              
   Exit Sub
End Sub

Public Sub CMIsTopLeftCFNeeded( _
                           ByRef oMD As IJDMemberDescription, _
                           ByRef bIsNeeded As Boolean)
                           
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
                                                           
    bIsNeeded = False
    Dim strAnswer As String
   
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
    
    If InStr(1, oSmItem.Name, "LT", vbTextCompare) > 0 Then
        GetSelectorAnswer oMD.CAO, "OutsideCorners", strAnswer
        If LCase(strAnswer) = LCase("Yes") Then
            Select Case sSectionType
            Case "FB", "EA", "UA", "BUTL3", "B"
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP) Then
                    bIsNeeded = True
                End If
            Case "BUT", "BUTL2"
                If DoesEdgeOverlapPart(oMD.CAO, JXSEC_WEB_LEFT) And DoesEdgeOverlapPart(oMD.CAO, JXSEC_TOP_FLANGE_LEFT_BOTTOM) Then
                    bIsNeeded = True
                End If
            End Select
        End If
    End If
    
   Exit Sub
End Sub

Public Sub CMConstructBottomLeftCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)

    Dim oSlotWrapper As StructDetailObjects.Slot
    Dim oPenetratedStructConnectable As IJStructConnectable
    Dim oPenetratedPorts As IJElements
    Dim oPenetrationLoc As IJDPosition
    
    Set oSlotWrapper = New StructDetailObjects.Slot
    Set oSlotWrapper.object = oMemberDescription.CAO
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    
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
    Dim oPortMB As IJModelBody
    Dim nCtx As Long
    Dim oSupport1 As Object
    Dim oSupport2 As Object
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
             'Set oSupport1 = oEdgePort
          End If
       End If
    Next

    'If Penetrating Part Is a Plate there may be multiple ports at 0 distance
    Dim nIndex As Integer
    Dim oPort As IJPort
    Dim oSurface As IJSurfaceBody
    Dim oPortNormal As New DVector
    Dim dDotProduct As Double
    Dim dMaxDotProduct As Double
    If TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        If oBasePortColl.Count = 1 Then
            Set oSupport1 = oBasePortColl.Item(1)
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
                    Set oSupport1 = oBasePortColl.Item(nIndex)
                End If
            Next
        End If
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        If oBasePortColl.Count = 1 Then
            Set oSupport1 = oBasePortColl.Item(1)
        ElseIf oBasePortColl.Count > 1 Then
            Dim oProfile As New StructDetailObjects.ProfilePart
            Set oProfile.object = oSlotWrapper.Penetrating
            
            Dim oMountingPort As IJPort
            Dim oMountingPortSurface As IJSurfaceBody
            Set oMountingPort = oProfile.MountingFacePort()
            Set oMountingPortSurface = oMountingPort.Geometry
           
            Dim oProjectedPoint As IJDPosition
            Dim oOutputNormal As IJDVector
            
            oTopologyLocate.GetProjectedPointOnModelBody oMountingPortSurface, oPenetrationLoc, oProjectedPoint, oOutputNormal

            
            Dim oMountingNormal As New DVector
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
                    Set oSupport1 = oBasePortColl.Item(nIndex)
                End If
            Next
        End If
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim nXid As Integer
    Dim eFeatureType As StructFeatureTypes
    
    nXid = JXSEC_WEB_LEFT
    Set oStructFeature = oMemberDescription.CAO
    If Not oStructFeature Is Nothing Then
       eFeatureType = oStructFeature.get_StructFeatureType
       oStructFeature.GetContourSegment nXid, oSupport2
    End If
    
    Dim sSelString As String
    sSelString = "RootCorner"
    
    Dim oSDOCornerFeature As IJSDOCornerFeature
    
    Set oSDOCornerFeature = New StructDetailObjectsEx.CornerFeature
    oSDOCornerFeature.Create _
                    oResourceManager, _
                    oFacePort, _
                    oSupport2, _
                    oSupport1, _
                    sSelString, _
                    oMemberDescription.CAO
    Set oObject = oSDOCornerFeature.object

End Sub

Public Sub CMConstructTopLeftCF( _
                            ByVal oMemberDescription As IJDMemberDescription, _
                            ByVal oResourceManager As IUnknown, _
                            ByRef oObject As Object)
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
    
    Set oModelHelper = New SGOModelBodyUtilities
    dMinDistance = 100
        
    Dim nPortIndex As Long
    Dim oStructPort As IJStructPort
    Dim oFacePort As IJPort
    Dim oEdgePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    Dim oPortMB As IJModelBody
    Dim nCtx As Long
    
    For nPortIndex = 1 To oPenetratedPorts.Count
       Set oStructPort = oPenetratedPorts.Item(nPortIndex)
       nCtx = oStructPort.ContextID
       
       If (nCtx And CTX_BASE) > 0 Then
          Set oFacePort = oStructPort
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
    
    Dim sXSectionType As String
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = oSlotWrapper.Penetrating
        
        sXSectionType = oProfile.sectionType
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
    
    
    If Not oStructFeature Is Nothing Then
        If sXSectionType = "FB" Or sXSectionType = "B" Or sXSectionType = "EA" Or sXSectionType = "UA" Then
           nXid = JXSEC_TOP
           oStructFeature.GetContourSegment nXid, oSupport2
        Else
           nXid = JXSEC_TOP_FLANGE_LEFT_BOTTOM
           oStructFeature.GetContourSegment nXid, oSupport2
        End If
          
       nXid = JXSEC_WEB_LEFT
       oStructFeature.GetContourSegment nXid, oSupport1
    End If
    
    Dim sSelString As String
    sSelString = "VariableEdgeCorner"
    
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
    
End Sub


