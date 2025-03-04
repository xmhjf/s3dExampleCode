Attribute VB_Name = "CollarRulesUtils"
Option Explicit
Private Const MODULE = CUSTOMERID + "CollarRules.CollarRulesUtils"
'

'  This method is private to this project and creates a corner snipe or a drain hole.(Essentially both are corner features).
'  Inputs:
'    oCollar                     ---- Collar object on which to place snipe or drain hole
'    eCtx                        ---- Base or Offset face (CTX_BASE,CTX_OFFSET)
'    eBottomXid              ---- Support1 port xid  ( JXSEC_BOTTOM, JXSEC_BOTTOM_LEFT,JXSEC_BOTTOM_RIGHT)
'    eWebXid                  ---- Support2 port xid (JXSEC_LEFT, JXSEC_RIGHT)
'    strSnipeOrDrainHole ---- Type of object to create ("SnipeOrCollar", "DrainHole")
'    oObjectCreated        ---- Snipe or drain created

Private Sub CreateSnipeOrDrainHole(ByVal oCollar As Object, _
                                                      ByVal oResourceManager As IUnknown, _
                                                      ByVal eCtx As eUSER_CTX_FLAGS, _
                                                      ByVal eBottomXid As JXSEC_CODE, _
                                                      ByVal eWebXid As JXSEC_CODE, _
                                                      ByVal strSnipeOrDrainHole As String, _
                                                      ByRef oObjectCreated As Object)
    On Error GoTo ErrorHandler
        
    Dim oFacePort As IJPort
    Dim oWebPort As IJPort
    Dim oBottomPort As IJPort

    Dim oCollarWrapper As New StructDetailObjects.Collar

    Set oCollarWrapper.object = oCollar
    oCollarWrapper.GetPortsForCornerFeature eCtx, eWebXid, eBottomXid, oFacePort, oWebPort, oBottomPort
    Set oCollarWrapper = Nothing
        
    Dim oCornerFeatureWrapper As New StructDetailObjects.CornerFeature
    
    ' Bottom is Support1, Web is Support2
    oCornerFeatureWrapper.Create oResourceManager, _
                                                  oFacePort, _
                                                  oBottomPort, _
                                                  oWebPort, _
                                                  strSnipeOrDrainHole, _
                                                  oCollar
    Set oFacePort = Nothing
    Set oWebPort = Nothing
    Set oBottomPort = Nothing
    
    Set oObjectCreated = oCornerFeatureWrapper.object
    Set oCornerFeatureWrapper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CreateSnipeOrDrainHole").Number
End Sub
                                                               
'  This method creates a snipe
'  Inputs:
'    oCollar                     ---- Collar object on which to place snipe or drain hole
'    eCtx                        ---- Base or Offset face (CTX_BASE,CTX_OFFSET)
'    eBottomXid              ---- Support1 port xid  ( JXSEC_BOTTOM, JXSEC_BOTTOM_LEFT,JXSEC_BOTTOM_RIGHT)
'    eWebXid                  ---- Support2 port xid (JXSEC_LEFT, JXSEC_RIGHT)
'    oObjectCreated        ---- Snipe created
Public Sub CreateCornerSnipe(ByVal oCollar As Object, _
                                             ByVal oResourceManager As IUnknown, _
                                             ByVal eCtx As eUSER_CTX_FLAGS, _
                                             ByVal eBottomXid As JXSEC_CODE, _
                                             ByVal eWebXid As JXSEC_CODE, _
                                             ByRef oObjectCreated As Object)

    On Error GoTo ErrorHandler
    
    CreateSnipeOrDrainHole oCollar, _
                                        oResourceManager, _
                                        eCtx, _
                                        eBottomXid, _
                                        eWebXid, _
                                        "SnipeOnCollar", _
                                        oObjectCreated
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CreateCornerSnipe").Number
End Sub

'  This method creates a drain hole
'  Inputs:
'    oCollar                     ---- Collar object on which to place snipe or drain hole
'    eCtx                        ---- Base or Offset face (CTX_BASE or CTX_OFFSET)
'    eBottomXid              ---- Support1 port xid  ( JXSEC_BOTTOM, JXSEC_BOTTOM_LEFT,JXSEC_BOTTOM_RIGHT)
'    eWebXid                  ---- Support2 port xid (JXSEC_LEFT, JXSEC_RIGHT)
'    oObjectCreated        ---- Drain hole created
 
Public Sub CreateDrainHole(ByVal oCollar As Object, _
                                         ByVal oResourceManager As IUnknown, _
                                         ByVal eCtx As eUSER_CTX_FLAGS, _
                                         ByVal eBottomXid As JXSEC_CODE, _
                                         ByVal eWebXid As JXSEC_CODE, _
                                         ByRef oObjectCreated As Object)
    On Error GoTo ErrorHandler
    
    CreateSnipeOrDrainHole oCollar, _
                                        oResourceManager, _
                                        eCtx, _
                                        eBottomXid, _
                                        eWebXid, _
                                        "DrainHole", _
                                        oObjectCreated
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CreateDrainHole").Number
End Sub
 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub ConstructBaseEdgePC(ByVal eBaseEdgeId As GSCADSDCreateModifyUtilities.JXSEC_CODE, _
                               ByVal oMD As IJDMemberDescription, _
                               ByVal oResourceManager As IUnknown, _
                               ByRef oObject As Object)
    '
    ' Create Physical Connection between Collar Edge and Base Plate
    Dim oBasePlatePort As IJPort
    Dim oCollarBottomPort As IJPort
    Dim oLastBasePlatePort As IJPort
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDO_Collar As StructDetailObjects.Collar
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    
    On Error GoTo ErrorHandler
    
    Set oSDO_Collar = New StructDetailObjects.Collar
    Set oSDO_Collar.object = oMD.CAO
    Set oCollarBottomPort = oSDO_Collar.SubPort(eBaseEdgeId)
    Set oBasePlatePort = oSDO_Collar.BasePlatePort
    Set oSDO_Collar = Nothing
    
    
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oLastBasePlatePort = oSDO_Helper.GetEquivalentLastPort(oBasePlatePort)
    Set oSDO_Helper = Nothing
    
    ' Construct PC
    Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
    oSDO_PhysicalConn.Create oResourceManager, oCollarBottomPort, oLastBasePlatePort, _
                             "TeeWeld", oMD.CAO, ConnectionStandard
    Set oObject = oSDO_PhysicalConn.object
    Set oSDO_PhysicalConn = Nothing
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ConstructBaseEdgePC").Number
End Sub
' This method returns following angles:
' 1. Angle between web and base plate
'    This angle can be used to calculate length of clip right side
' 2. Angle between profile and penetrated entity
'    This angle can be used to calculate adjusted collar width
' 3. Angle between profile top and web (left or right,dependent on eWebLeftorWebRight value)
'    This angle can be used to calculate adjusted collar width
Public Sub GetSlotAngles(ByVal oSlot As Object, _
                         ByVal eWebLeftOrWebRight As JXSEC_CODE, _
                         Optional ByRef dWebBaseAngle As Double, _
                         Optional ByRef dProfilePenetratedAngle As Double = 0, _
                         Optional ByRef dTopWebAngle As Double = 0, _
                         Optional ByVal nCollarSide As Integer = 0, _
                         Optional ByVal bGetWebBaseAngle As Boolean = True, _
                         Optional ByVal bGetPenetratingAngle As Boolean = False, _
                         Optional ByVal bGetTopWebAngle As Boolean = False)
                         
   Dim oSlotWrapper As New StructDetailObjects.Slot
   Dim oTopologyLocate As TopologyLocate
   Dim oPenetrationPoint As IJDPosition
      
   On Error Resume Next

   Set oSlotWrapper.object = oSlot
   Set oPenetrationPoint = oSlotWrapper.PenetrationLocation
   Set oTopologyLocate = New TopologyLocate

   ' Get Base Plate Normal & Projected Point on Base Plate
   Dim oBasePort As IJPort
   Dim oPointOnBasePlate As IJDPosition
   Dim oBasePortGeom As Object
   Dim oBasePlateNormal As IJDVector
   
   Dim oSlotMappingRule As IJSlotMappingRule

    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Set oBasePortGeom = oTopologyLocate.GetBasePlatePort(oSlotWrapper.Penetrating)
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance

        Dim oMappedPorts As JCmnShp_CollectionAlias
        Set oMappedPorts = New Collection

        oSlotMappingRule.GetEmulatedPorts oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, oBasePort, oMappedPorts
        Set oBasePortGeom = oBasePort.Geometry
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
    
    oTopologyLocate.GetProjectedPointOnModelBody oBasePortGeom, _
                                        oPenetrationPoint, _
                                        oPointOnBasePlate, _
                                        oBasePlateNormal
    

   'Get Penetrated Normal & Projected Point on Penetrated Part
   Dim oPenetratedPort As IJPort
   Dim oPointOnPenetrated As IJDPosition
   Dim oPenetratedNormal As IJDVector
   
   If TypeOf oSlotWrapper.Penetrated Is IJPlatePart Then
      Dim oPenetratedPlate As New StructDetailObjects.PlatePart
      
      ' Penetrated is plate,use base or offset
      Set oPenetratedPlate.object = oSlotWrapper.Penetrated
      If nCollarSide = 0 Then
         Set oPenetratedPort = oPenetratedPlate.BasePort(BPT_Base)
      Else
         Set oPenetratedPort = oPenetratedPlate.BasePort(BPT_Offset)
      End If
   Else
      ' Penetrated is profile part,use web left port for now
      Dim oPenetratedProfile As New StructDetailObjects.ProfilePart
      
      Set oPenetratedProfile.object = oSlotWrapper.Penetrated
      Set oPenetratedPort = oPenetratedProfile.SubPort(JXSEC_WEB_LEFT)
   End If
   
   oTopologyLocate.GetProjectedPointOnModelBody oPenetratedPort.Geometry, _
                                                oPenetrationPoint, _
                                                oPointOnPenetrated, _
                                                oPenetratedNormal

   'Get Penetrating Web Normal & Projected Point on Web
   'Also get the Top Port
   Dim oWebPort As IJPort
   Dim oPointOnWeb As IJDPosition
   Dim oWebNormal As IJDVector

    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oPenetratingProfile As New StructDetailObjects.ProfilePart
        Set oPenetratingProfile.object = oSlotWrapper.Penetrating
        Set oWebPort = oPenetratingProfile.SubPort(eWebLeftOrWebRight)
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Set oWebPort = oMappedPorts.Item(CStr(eWebLeftOrWebRight))
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If

    oTopologyLocate.GetProjectedPointOnModelBody oWebPort.Geometry, _
                                                oPenetrationPoint, _
                                                oPointOnWeb, _
                                                oWebNormal

   Dim dDot As Double
   Dim dAngle As Double
   
   'Get the Angle Between the Base Plate and the Web
   If bGetWebBaseAngle = True Then
      Dim oPenetratedCrossBase As IJDVector
      Dim oPenetratedCrossWeb As IJDVector

      Set oPenetratedCrossBase = oPenetratedNormal.Cross(oBasePlateNormal)
      Set oPenetratedCrossWeb = oPenetratedNormal.Cross(oWebNormal)
   
      oPenetratedCrossBase.Length = 1
      oPenetratedCrossWeb.Length = 1
      dDot = oPenetratedCrossBase.Dot(oPenetratedCrossWeb)
   
      If dDot < -1 Then
         dDot = -1
      ElseIf dDot > 1 Then
         dDot = 1
      End If
      
      If Abs(dDot) < 0.000001 Then
         dAngle = HALF_PI
      Else
         dAngle = Atn(Sqr(1 / (dDot * dDot) - 1))
      End If
        
      If dDot < 0 Then
         dAngle = PI - dAngle
      End If
      
      dWebBaseAngle = dAngle
   End If
      
   Dim oPointOnProfile As IJDPosition
   Dim oProfileDir As IJDVector
   Dim oWireBody As IJWireBody

   If bGetPenetratingAngle = True Or bGetTopWebAngle = True Then
      'This appears to get the angle between the Base X ProfileDir and the penetrated
      'plate.  Returns 0 when they are perpendicular
      If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Set oWireBody = oPenetratingProfile.ParentSystem
        oTopologyLocate.GetProjectedPointOnModelBody oWireBody, _
                                                   oPenetrationPoint, _
                                                   oPointOnProfile, _
                                                   oProfileDir
      ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        'Use the Web x Base as the Profile Dir (to get the same results as above)
        Set oProfileDir = oBasePlateNormal.Cross(oWebNormal)
      End If
      
   End If

   ' Get angle between profile and penetrated part
   If bGetPenetratingAngle = True Then
      Dim oBaseCrossProfileDir As IJDVector
      Dim oBaseCrossPenetratedNormal As IJDVector
      
      Set oBaseCrossProfileDir = oBasePlateNormal.Cross(oProfileDir)
      Set oBaseCrossPenetratedNormal = oBasePlateNormal.Cross(oPenetratedNormal)
      oBaseCrossProfileDir.Length = 1
      oBaseCrossPenetratedNormal = 1
      dDot = oBaseCrossProfileDir.Dot(oBaseCrossPenetratedNormal)
      
      If dDot < -1 Then
         dDot = -1
      ElseIf dDot > 1 Then
         dDot = 1
      End If
      If Abs(dDot) < 0.000001 Then
         dAngle = HALF_PI
      Else
         dAngle = Atn(Sqr(1 / (dDot * dDot) - 1))
      End If
   
      If dDot < 0 Then
         dAngle = PI - dAngle
      End If
      
      dProfilePenetratedAngle = dAngle
   End If

    ' Get angle between top and web
    If bGetTopWebAngle = True Then
        
        Dim oWebRightPort As IJPort
        Dim oPointOnWebRight As IJDPosition
        Dim oWebRightNormal As IJDVector
        Dim oTopPort As IJPort
        Dim oPointOnTop As IJDPosition
        Dim oTopNormal As IJDVector
        
        If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
            Set oWebRightPort = oPenetratingProfile.SubPort(JXSEC_WEB_RIGHT)
            Set oTopPort = oPenetratingProfile.SubPort(JXSEC_TOP)
        ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
            Set oWebRightPort = oMappedPorts.Item(CStr(JXSEC_WEB_RIGHT))
            Set oTopPort = oMappedPorts.Item(CStr(JXSEC_TOP))
        Else
            'Unsupported Penetrating Object
            Exit Sub
        End If

        oTopologyLocate.GetProjectedPointOnModelBody oWebRightPort.Geometry, _
                                                oPenetrationPoint, _
                                                oPointOnWebRight, _
                                                oWebRightNormal
                                                
        oTopologyLocate.GetProjectedPointOnModelBody oTopPort.Geometry, _
                                                oPenetrationPoint, _
                                                oPointOnTop, _
                                                oTopNormal
        
        ' oWebTangent and oTopTangent were already on sketch plane,use it directly
        oTopNormal.Length = 1
        oWebRightNormal.Length = 1
        
        Dim oW As IJDVector
        Set oW = oWebRightNormal.Cross(oBasePlateNormal)
        
        Dim oProjectedTop As IJDVector
        Set oProjectedTop = oW.Cross(oTopNormal.Cross(oW))
        
        dDot = oProjectedTop.Dot(oWebRightNormal)
        If dDot < -1 Then
            dDot = -1
        ElseIf dDot > 1 Then
            dDot = 1
        End If
        
        If Abs(dDot) < 0.000001 Then
            dAngle = HALF_PI
        Else
            dAngle = Atn(Sqr(1 / (dDot * dDot) - 1))
        End If

        If dDot < 0 Then
            dAngle = PI - dAngle
        End If

        dTopWebAngle = dAngle
'
'      Dim sArrayOfEdgeIds() As String
'      Dim oPlaneToSliceProfile As IJPlane
'      Dim oProfileEdges As IJElements
'      Dim oComputeSurfaceHelper As StructGenericUtilities.ComputeSurfaceHelper
'      Dim oNullObject As Object
'
'      Set oComputeSurfaceHelper = New StructGenericUtilities.ComputeSurfaceHelper
'      Set oPlaneToSliceProfile = oComputeSurfaceHelper.CreateFinitePlane2( _
'                                           oNullObject, _
'                                           oNullObject, _
'                                           oPenetrationPoint.x, _
'                                           oPenetrationPoint.y, _
'                                           oPenetrationPoint.z, _
'                                           oPenetratedNormal.x, _
'                                           oPenetratedNormal.y, _
'                                           oPenetratedNormal.z, _
'                                           oPenetrationPoint.x - 1.5 * oPenetratingProfile.Height, _
'                                           oPenetrationPoint.y - 1.5 * oPenetratingProfile.Height, _
'                                           oPenetrationPoint.z - 1.5 * oPenetratingProfile.Height, _
'                                           oPenetrationPoint.x + 1.5 * oPenetratingProfile.Height, _
'                                           oPenetrationPoint.y + 1.5 * oPenetratingProfile.Height, _
'                                           oPenetrationPoint.z + 1.5 * oPenetratingProfile.Height)
'
'      Dim dExtend As Double
'      Dim oProfileStartPos As IJDPosition
'      Dim oProfileEndPos As IJDPosition
'
'      oWireBody.GetEndPoints oProfileStartPos, oProfileEndPos
'      If oPointOnProfile.DistPt(oProfileStartPos) < 0.0005 Or _
'         oPointOnProfile.DistPt(oProfileEndPos) < 0.0005 Then
'         dExtend = 3.5
'      Else
'         dExtend = -1
'      End If
'
'      sArrayOfEdgeIds = oTopologyLocate.IntersectProfilePartWithSketchingPlane( _
'                                        oPenetratingProfile.object, _
'                                        oPlaneToSliceProfile, _
'                                        dExtend, _
'                                        oProfileEdges)
'      Dim nIndex As Long
'      Dim oTopWire As Object
'      Dim oWebRightWire As Object
'      Dim sEdgeId As String
'
'      For nIndex = 1 To oProfileEdges.Count
'         sEdgeId = sArrayOfEdgeIds(nIndex - 1)
'         If sEdgeId = "514" Then
'            Set oTopWire = oProfileEdges.Item(nIndex)
'         End If
'
'         If sEdgeId = "258" Then
'            Set oWebRightWire = oProfileEdges.Item(nIndex)
'         End If
'         If Not oTopWire Is Nothing And Not oWebRightWire Is Nothing Then
'            Exit For
'         End If
'      Next
'
'      '''''''''''''''
'
'
'
'      If Not oTopWire Is Nothing And Not oWebRightWire Is Nothing Then
'         Dim oStartPos As IJDPosition
'         Dim oEndPos As IJDPosition
'         Dim oStartDir As IJDVector
'         Dim oEndDir As IJDVector
'         Dim oPenetratedCrossTop As IJDVector
'         Dim oTopTangent As IJDVector
'         Dim oWebTangent As IJDVector
'
'         Set oWireBody = oTopWire
'         oWireBody.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir
'
'         ' Set top tangent in the same direction as web normal
'         dDot = oStartDir.Dot(oWebNormal)
'         If dDot < 0 Then
'            Set oTopTangent = oStartDir.Clone
'            oTopTangent.Set -oStartDir.x, -oStartDir.y, -oStartDir.z
'         Else
'            Set oTopTangent = oStartDir
'         End If
'
'         Set oWireBody = oWebRightWire
'         oWireBody.GetEndPoints oStartPos, oEndPos, oStartDir, oEndDir
'
'         ' Set web tangent in the opposite direction to base plate normal
'         dDot = oStartDir.Dot(oBasePlateNormal)
'         If dDot > 0 Then
'            Set oWebTangent = oStartDir.Clone
'            oWebTangent.Set -oStartDir.x, -oStartDir.y, -oStartDir.z
'         Else
'            Set oWebTangent = oStartDir
'         End If
'
'         ' oWebTangent and oTopTangent were already on sketch plane,use it directly
'         oTopTangent.Length = 1
'         oWebTangent.Length = 1
'
'         dDot = oTopTangent.Dot(oWebTangent)
'         If dDot < -1 Then
'            dDot = -1
'         ElseIf dDot > 1 Then
'            dDot = 1
'         End If
'
'         If Abs(dDot) < 0.000001 Then
'            dAngle = HALF_PI
'         Else
'            dAngle = Atn(Sqr(1 / (dDot * dDot) - 1))
'         End If
'
'         If dDot < 0 Then
'            dAngle = PI - dAngle
'         End If
'
'         dTopWebAngle = dAngle
'      End If
   End If
   
End Sub

Public Function EstimateClipRightSideLength(ByVal oSlot As Object, _
                                            ByVal dAngle As Double) As Double
    
    'Get the Slot Wrapper
    Dim oSlotWrapper As New StructDetailObjects.Slot
    Set oSlotWrapper.object = oSlot
   
    'Get the Section Type, Depth, and Width
    Dim dProfileHeight As Double
    Dim sXSectionType As String
    Dim dProfileWidth As Double
    If TypeOf oSlotWrapper.Penetrating Is IJProfile Then
        Dim oProfileWrapper As New StructDetailObjects.ProfilePart
        Set oProfileWrapper.object = oSlotWrapper.Penetrating
        dProfileHeight = oProfileWrapper.Height
        dProfileWidth = oProfileWrapper.Width
        sXSectionType = oProfileWrapper.sectionType
    ElseIf TypeOf oSlotWrapper.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        
        Dim dWebTh As Double
        Dim dFlangeTh As Double
        oSlotMappingRule.GetSectionDimensions oSlotWrapper.Penetrating, oSlotWrapper.Penetrated, dProfileHeight, dProfileWidth, dWebTh, dFlangeTh
    Else
        'Unsupported Penetrating Object
        EstimateClipRightSideLength = 0
        Exit Function
    End If
    
    dProfileHeight = Round(dProfileHeight, 3)
    
    Dim dClipWidth As Double
    Dim dClipHeight As Double
    Dim dTopClearance As Double
    Dim dBottomClearance As Double
   
    dClipWidth = 0
    dClipHeight = 0
   
    Select Case sXSectionType
        Case "EA", "UA"
            If dProfileHeight >= 0.2 And dProfileHeight <= 0.45 Then
                dClipWidth = dProfileWidth + 0.065
                dTopClearance = 0.05
            
                If dProfileHeight = 0.2 Then
                    dBottomClearance = 0.04
                ElseIf dProfileHeight > 0.2 And dProfileHeight <= 0.3 Then
                    dBottomClearance = 0.05
                Else
                    dBottomClearance = 0.075
                End If
            End If
         
        Case "B"
            If dProfileHeight >= 0.18 Then
                dClipWidth = 0.12
                dTopClearance = 0.05
            
                If dProfileHeight <= 0.2 Then
                    dBottomClearance = 0.03
                Else
                    dBottomClearance = 0.05
                End If
            End If
         
        Case "FB"
            If dProfileHeight = 0.2 Then
                dClipWidth = 0.13
                dTopClearance = 0.05
                dBottomClearance = 0.03
            
            ElseIf dProfileHeight > 0.2 And dProfileHeight <= 0.3 Then
                dClipWidth = 0.13
                dTopClearance = 0.05
                dBottomClearance = 0.05
            
            ElseIf dProfileHeight > 0.3 Then
                dClipWidth = 0.16
                dTopClearance = 0.05
                dBottomClearance = 0.075
            End If
         
        Case "BUTL2"
            If dProfileHeight >= 0.2 And dProfileHeight <= 1 Then
                dClipWidth = dProfileWidth + 0.065
            
                If dProfileHeight >= 0.2 And dProfileHeight <= 0.3 Then
                    dTopClearance = 0.06
                    dBottomClearance = 0.05
               
                ElseIf dProfileHeight > 0.3 And dProfileHeight <= 0.45 Then
                    dTopClearance = 0.06
                    dBottomClearance = 0.075
               
                ElseIf dProfileHeight > 0.45 And dProfileHeight <= 0.6 Then
                    dTopClearance = 0.09
                    dBottomClearance = 0.075
               
                ElseIf dProfileHeight > 0.6 And dProfileHeight <= 0.8 Then
                    dTopClearance = 0.1
                    dBottomClearance = 0.1
               
                ElseIf dProfileHeight > 0.8 And dProfileHeight <= 1 Then
                    dTopClearance = 0.1
                    dBottomClearance = 0.15
                End If
            
            End If
         
        Case "BUT"
            If dProfileHeight >= 0.2 And dProfileHeight <= 1.5 Then
         
                dClipWidth = dProfileWidth + 0.065
            
                If dProfileHeight > 0.2 And dProfileHeight <= 0.3 Then
                    dTopClearance = 0.06
                    dBottomClearance = 0.05
               
                ElseIf dProfileHeight > 0.3 And dProfileHeight <= 0.45 Then
                    dTopClearance = 0.06
                    dBottomClearance = 0.075
               
                ElseIf dProfileHeight > 0.45 And dProfileHeight <= 0.6 Then
                    dTopClearance = 0.09
                    dBottomClearance = 0.075
               
                ElseIf dProfileHeight > 0.6 And dProfileHeight <= 0.8 Then
                    dTopClearance = 0.1
                    dBottomClearance = 0.1
               
                ElseIf dProfileHeight > 0.8 And dProfileHeight <= 1.5 Then
                    dTopClearance = 0.1
                    dBottomClearance = 0.15
                End If
            End If
         
        Case Else
      

    End Select
   
    If dClipWidth > 0 Then
        Dim dRightSideLength As Double

        dRightSideLength = 0
        dClipHeight = dProfileHeight - dTopClearance - dBottomClearance
        If dAngle > HALF_PI Then
            dAngle = dAngle - HALF_PI
        End If
        dRightSideLength = dClipHeight - dClipWidth * Tan(dAngle)
            
        EstimateClipRightSideLength = dRightSideLength
    Else
        EstimateClipRightSideLength = 0
    End If
   
End Function

Public Function CheckProfileOnHullPenetratingLBH(oSlotOrCollar As Object) As Boolean
   CheckProfileOnHullPenetratingLBH = False
   
    Dim oPenetrated As Object
    Dim oBasePlate As Object
    
    'Get the Base Plate From the Slot or Collar
    If TypeOf oSlotOrCollar Is IJCollarPart Then
        Dim oCollarWrapper As New StructDetailObjects.Collar
        Set oCollarWrapper.object = oSlotOrCollar
        
        Set oPenetrated = oCollarWrapper.Penetrated
        Set oBasePlate = oCollarWrapper.BasePlate
        
        Set oCollarWrapper = Nothing
    ElseIf TypeOf oSlotOrCollar Is IJStructFeature Then
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oSlotOrCollar
        
        If oStructFeature.get_StructFeatureType <> SF_Slot Then
            Exit Function
        End If
        
        Dim oSlotWrapper As New StructDetailObjects.Slot
        Set oSlotWrapper.object = oSlotOrCollar
        
        Set oPenetrated = oSlotWrapper.Penetrated
        Set oBasePlate = oSlotWrapper.BasePlate
        
        Set oSlotWrapper = Nothing
    End If

    'Check if the Stiffened Base Plate is a Hull
    Dim bStiffenedPlateIsHull As Boolean
    bStiffenedPlateIsHull = False
    If TypeOf oBasePlate Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oBasePlate
        If oPlate.plateType = Hull Then
            bStiffenedPlateIsHull = True
        End If
    Else
        Exit Function
    End If

    'Check if the Penetrated Plate is a Longitudinal Bulkhead
    Dim bPenetratedIsLBH As Boolean
    bPenetratedIsLBH = False
    If bStiffenedPlateIsHull Then
        If TypeOf oPenetrated Is IJPlate Then
            Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
            Set oPlatePartWrapper.object = oPenetrated
            
            If oPlatePartWrapper.plateType = LBulkheadPlate Then
                bPenetratedIsLBH = True
            End If
        Else
            Exit Function
        End If
    End If
   
    If bStiffenedPlateIsHull = True And bPenetratedIsLBH = True Then
        CheckProfileOnHullPenetratingLBH = True
    End If

End Function
'This method returns slot angle:
'  If no parameter rule associated with "SlotAngle", get the value from "SlotAngle" attribute
'  Otherwise,
'     If the angle is set by parameter rule,execute the rule and get the value from rule output
'     If user already overrode the value, get the value from "SlotAngle" attribute
Public Function GetSlotAngle(oCollar As Object) As Double
   On Error GoTo ErrorHandler
   GetSlotAngle = 0
   
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oSlot As Object
   
   Set oCollarWrapper.object = oCollar
   Set oSlot = oCollarWrapper.Slot
   
   Dim oSO As IJSmartOccurrence
   Dim oSI As IJSmartItem
   Dim bIsParameter As Boolean
   Dim sParameterName As String
   Dim oParameterRuleDef As IJDSymbolDefinition
   
   sParameterName = "SlotAngle"
   Set oSO = oSlot
   Set oSI = oSO.ItemObject
   bIsParameter = False
   
   Set oParameterRuleDef = oSI.ParameterRuleDef
   If Not oSI.ParameterRuleDef Is Nothing Then
      ' Execute parameter rule
      Dim oSmartOccHelper As IJSmartOccurrenceHelper
      Dim oOutputColl As IJDOutputCollection

      Set oSmartOccHelper = New CSmartOccurrenceCES
      On Error Resume Next
      Set oOutputColl = oSmartOccHelper.ExecuteParameterRule(oParameterRuleDef, oSO)
      On Error GoTo ErrorHandler
      
      Dim oOutputObj As Object
      Dim oParameterContent As IJDParameterContent
      Dim dAngle As Double
      
      If Not oOutputColl Is Nothing Then
          Set oOutputObj = oOutputColl.GetOutput(sParameterName)
          Set oParameterContent = oOutputObj
          dAngle = oParameterContent.UomValue
          bIsParameter = True
      End If
   End If
   
   Dim bSetByParameterRule As Boolean
   
   bSetByParameterRule = False
   If bIsParameter = True Then
      ' It's associated with a parameter rule,check if the value is set by rule
      Dim oOutputControl As IJOutputControl
      Dim bEditable As Boolean
      
      bEditable = False
      Set oOutputControl = oSlot
      bEditable = oOutputControl.OutputEditable(oParameterRuleDef, sParameterName)
      If bEditable = True Then
         ' User already overrode the parameter
         bSetByParameterRule = False
      Else
         ' The value is set by rule
         bSetByParameterRule = True
      End If
   End If
   
   If bSetByParameterRule = False Then
      ' Not set by rule, get the value from attribute
      Dim oSDHelper As New StructDetailObjects.Helper
      Dim vAngleValue As Variant
      Dim vResultType As EResultType
      Dim sTableName As String
      Dim nCodeListValue As Long
      Dim sShortDesc As String
      Dim sLongDesc As String
   
      oSDHelper.GetCustomAttributeParamValue oSlot, _
                                          "IJUASlotParameter", _
                                          sParameterName, _
                                          vAngleValue, _
                                          vResultType, _
                                          sTableName, _
                                          nCodeListValue, _
                                          sShortDesc, _
                                          sLongDesc
      dAngle = vAngleValue
   End If
   
   If dAngle > 0 Then
      GetSlotAngle = dAngle
   End If
   
ErrorHandler:
   Exit Function
   
End Function

'Get all the profile parts of the given profile system
Public Sub GetLeafSystemsFromRoot(oRootSystem As Object, ppLeafSystems As Collection)
On Error GoTo ErrorHandler
    
    Dim oRoot As IJDesignParent
    Dim oChildColl As IJDObjectCollection
    Dim oObject As Object
    
    Set oRoot = oRootSystem
    Set ppLeafSystems = New Collection
    
    oRoot.GetChildren oChildColl
    
    If TypeOf oRootSystem Is IJPlate Then
        For Each oObject In oChildColl
               If TypeOf oObject Is IJPlateSystem Then
                    ppLeafSystems.Add oObject
               End If
        Next
    ElseIf TypeOf oRootSystem Is IJStiffener Then
        For Each oObject In oChildColl
               If TypeOf oObject Is IJStiffenerSystem Then
                    ppLeafSystems.Add oObject
               End If
        Next
    End If

Set oRoot = Nothing
Set oObject = Nothing
Set oChildColl = Nothing

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, "Helper.bas", "GetProfilePartFromSystem").Number
                  
End Sub

Public Sub GetLeafPart(oLeafSystem As IJSystem, ByRef pLeafPart As Object)
On Error GoTo ErrorHandler
        
        Dim oLeafPart As Object
        Dim oChildren As IJDObjectCollection
        
        Dim oParent As IJDesignParent
        Set oParent = oLeafSystem
        
        If TypeOf oLeafSystem Is IJPlateSystem Then
            oParent.GetChildren oChildren
            For Each oLeafPart In oChildren
                If TypeOf oLeafPart Is IJPlatePart Then
                    Set pLeafPart = oLeafPart
                End If
            Next
        ElseIf TypeOf oLeafSystem Is IJStiffenerSystem Then
            oParent.GetChildren oChildren
            For Each oLeafPart In oChildren
                If TypeOf oLeafPart Is IJProfilePart Then
                    Set pLeafPart = oLeafPart
                End If
            Next
        End If
        
Set oChildren = Nothing
Set oParent = Nothing
Set oLeafPart = Nothing

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, "Helper.bas", "GetProfilePartFromSystem").Number

End Sub

 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub ConstructBaseEdgePC2(ByVal eBaseEdgeId As GSCADSDCreateModifyUtilities.JXSEC_CODE, _
                               ByVal oMD As IJDMemberDescription, _
                               ByVal oResourceManager As IUnknown, _
                               ByRef oObject As Object)

On Error GoTo ErrorHandler
    Dim oCollarWrapper As New StructDetailObjects.Collar
    Set oCollarWrapper.object = oMD.CAO
    
    Dim oCollarDefCM As New SMCollarRules.CollarDefCM
    
      
    Dim oCollarBottomPort As IJPort
    Set oCollarBottomPort = oCollarWrapper.SubPort(eBaseEdgeId)
    
    Dim oLastBasePlatePort As IJPort
    
    Dim pSecondBasePCValid As Boolean
    pSecondBasePCValid = oCollarDefCM.IsSecondBasePCValid(oCollarWrapper.object, oLastBasePlatePort)
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
    
Cleanup:

    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ConstructBaseEdgePC").Number
End Sub

'This method gives connected objects sharing the same parent as that of the input object, however is expensive
'as it processes all the assembly connected objects
Public Sub GetConnectedLeafPartsOfSameParent(oPart As Object, ppLeafParts As Collection)
On Error GoTo ErrorHandler
    
    Dim oSDObjectsHelper As New StructDetailObjects.Helper
    
    Dim oRootSystem As Object
    Set oRootSystem = oSDObjectsHelper.Object_RootParentSystem(oPart)

    Dim nConnData As Long
    Dim zConnData() As ConnectionData
    Dim ncount As Long
    
    Dim nRootSystem As Object
'    Dim ncount As Integer
    Dim oObject As Object
    Set ppLeafParts = New Collection
    
    Dim oParent As IJDesignParent, ppChildren As IJDObjectCollection, pObject As Object
    
    If TypeOf oPart Is IJPlatePart Then
        
        oSDObjectsHelper.Object_AppConnections oPart, AppConnectionType_Assembly, nConnData, zConnData()
        
        For ncount = 1 To nConnData
            If TypeOf zConnData(ncount).ToConnectable Is IJPlatePart Then
                Set nRootSystem = oSDObjectsHelper.Object_RootParentSystem(zConnData(ncount).ToConnectable)
                If nRootSystem Is oRootSystem Then
                    Set oObject = zConnData(ncount).ToConnectable
                    ppLeafParts.Add zConnData(ncount).ToConnectable
                End If
                Set nRootSystem = Nothing
            End If
        Next

    ElseIf TypeOf oPart Is IJProfilePart Then
        
        oSDObjectsHelper.Object_AppConnections oPart, AppConnectionType_Assembly, nConnData, zConnData()
        
        For ncount = 1 To nConnData
            If TypeOf zConnData(ncount).ToConnectable Is IJProfilePart Then
                Set nRootSystem = oSDObjectsHelper.Object_RootParentSystem(zConnData(ncount).ToConnectable)
                If nRootSystem Is oRootSystem Then
                    ppLeafParts.Add zConnData(ncount).ToConnectable
                End If
                Set nRootSystem = Nothing
            End If
        Next
    ElseIf TypeOf oPart Is IJPlate Then
        Set oParent = oPart
        oParent.GetChildren ppChildren
        For Each pObject In ppChildren
            If TypeOf oObject Is IJPlatePart Then
                GetConnectedLeafPartsOfSameParent oObject, ppLeafParts
            End If
        Next
        
    ElseIf TypeOf oPart Is IJProfile Then
        Set oParent = oPart
        oParent.GetChildren ppChildren
        For Each pObject In ppChildren
            If TypeOf oObject Is IJProfilePart Then
                GetConnectedLeafPartsOfSameParent oObject, ppLeafParts
            End If
        Next
    Else
        Exit Sub
    End If

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, "Helper.bas", "GetProfilePartFromSystem").Number
                  
End Sub

'Get the parts connected to the lateral port of an object(The connection should be of Lateral edge to Lateral Edge).
Public Sub GetLateralConnectedParts(oPart As Object, ppLeafParts As Collection)
On Error GoTo ErrorHandler

    Set ppLeafParts = New Collection
    
    Dim oSDObjectsHelper As New StructDetailObjects.Helper
    
    Dim oConnectable As IJConnectable
    If TypeOf oPart Is IJConnectable Then
        Set oConnectable = oPart
    End If
    If oConnectable Is Nothing Then
        Exit Sub
    End If
    
    Dim oPortsOnPart As IJElements, oLateralPorts As Collection
    Set oLateralPorts = New Collection
    oConnectable.enumPorts oPortsOnPart, PortFace
    
    If oPortsOnPart Is Nothing Then
        Exit Sub
    End If
    
    Dim oStructPort As IJStructPort, oPort As Object
    Dim ctx As eUSER_CTX_FLAGS, otherPortCtx As eUSER_CTX_FLAGS
    
    For Each oPort In oPortsOnPart
        Set oStructPort = oPort
        ctx = oStructPort.ContextID
        If ctx And CTX_BASE Then
        
        ElseIf ctx And CTX_OFFSET Then
        
        Else
            oLateralPorts.Add oPort
        End If
    Next
    
    Set oPort = Nothing
    If oLateralPorts.Count = 0 Then
        Exit Sub
    End If
    
    Dim oLPort As IJPort, oConnections As IJElements, oConnection As Object
    Dim oStructconnection As IMSStructConnection.IJStructConnection
    Dim oOtherPort As IJStructPort, oOtherConnectable As IJConnectable
    
    For Each oLPort In oLateralPorts
        oLPort.enumConnections oConnections, ConnectionAssembly, ConnectionStandard
        
        If Not oConnections Is Nothing Then
            For Each oConnection In oConnections
                If TypeOf oConnection Is IJStructConnection Then
                    Set oStructconnection = oConnection
                    oStructconnection.GetOtherPortAndConnectable oLPort, oOtherPort, oOtherConnectable
                    
                    otherPortCtx = oOtherPort.ContextID
                    If otherPortCtx = CTX_BASE Then
                        Exit For
                    ElseIf otherPortCtx = CTX_OFFSET Then
                        Exit For
                    Else
                        ppLeafParts.Add oOtherConnectable
                    End If
                End If
            Next
        End If
    Next
    
Cleanup:
    Set oSDObjectsHelper = Nothing
    Set oConnectable = Nothing
    Set oPortsOnPart = Nothing
    Set oLateralPorts = Nothing
    Set oStructPort = Nothing
    Set oLPort = Nothing
    Set oConnections = Nothing
    Set oConnection = Nothing
    Set oStructconnection = Nothing
    Set oOtherPort = Nothing
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, "CollarRulesUtils.bas", "GetLateralConnectedParts").Number
                  
End Sub


