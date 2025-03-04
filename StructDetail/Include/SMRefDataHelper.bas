Attribute VB_Name = "SMRefDataHelper"
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012-16, Intergraph Corporation.  All rights reserved.
'
'  File        : SMRefDataHelper.bas
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'   17/04/2012 - GH -CR-174918
'           Added  New methods 'GetStiffnedPlatePort_At_Position' to get the plate port(base or offset) as bounding.
'           and the method 'IsFreeEndCutFromExtend_ConvexKnuckle' to check wheather the FreeEndCut is From Extend
'   3-May-2013 - Alligators -TR-211987
'           Added input validation check in IfSeamsExistWithInDistance.
'
'    21/Nov/2013 - svsmylav
'          TR-234194 'CheckIfACisWithinCFRange' and 'GetFeatureContourRngBox' new methods are added.
'    30/Jan/2014 - GH  -    TR-228731  Created New method GetKnucklesAndSeamsonProfile() method
'                           updated IfSeamsExistWithInDistance() method to handle when seam exists on stiffener cases
'    12/May/2015 - GH- CR-260982 - Added New method IsValueBetween() to check the given value is with in the range
'
'         DI-259156: Deleted 'CheckIfACisWithinCFRange' method.
'*********************************************************************************************
Private Const MODULE As String = "S:\StructDetail\Data\Include\SMRefDataHelper.bas"
Const TOLERANCE_VALUE = 0.000011
Public Const C_BaseSide = "Base"
Public Const C_OffsetSide = "Offset"



'************************************************************************
'  For Tripping, Bracket By plane
' Method : IsBracket ------- checkes whether object is a Tripping Bracket
'                             or Bracket by plane
'
'  Inputs --------- Any object can be sent into the method
'
'  Output --------- True or False
'                   True --- if the object is Tripping Bracket or a Bracket by plane
'                   False --- if the object is neither Tripping Bracket nor Bracket by plane
'************************************************************************

Public Function IsBracket(oObject As Object) As Boolean

    On Error GoTo ErrorHandler
    Const sMETHOD = "IsBracket"

    IsBracket = False

    ' --------------------------------
    ' If a part, see if system-derived
    ' --------------------------------
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper

    Dim oRootObject As Object ' either standalone or system
    If TypeOf oObject Is IJPlate Then
        oStructDetailHelper.IsPartDerivedFromSystem oObject, oRootObject, True

        If oRootObject Is Nothing Then
            Set oRootObject = oObject
        End If
    Else
        Exit Function
    End If

    ' ----------------------------------------------------------------------
    ' If root object is a bracket-type plate part, it is a standalone object
    ' ----------------------------------------------------------------------
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils

    Dim plateType As StructPlateType
    plateType = CollarPlate
    If TypeOf oRootObject Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oRootObject
        plateType = oPlate.plateType
    End If

    If TypeOf oRootObject Is IJPlatePart And plateType = BracketPlate Then
        IsBracket = True
    ElseIf oPlateUtil.IsTrippingBracket(oRootObject) Then
        IsBracket = True
    ElseIf oPlateUtil.IsBracketByPlane(oRootObject) Then
        IsBracket = True
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

'Function to find if object is a collar or not
Public Function IsCollar(oObject As Object) As Boolean

    On Error GoTo ErrorHandler
    
    IsCollar = False
    
    If oObject Is Nothing Then Exit Function
    
    If TypeOf oObject Is IJPlate Then
        Dim oPlateUtil As IJPlateAttributes
        Set oPlateUtil = New PlateUtils
    
        Dim plateType As StructPlateType
    
        Dim oPlate As IJPlate
        Set oPlate = oObject
        plateType = oPlate.plateType
    
        If TypeOf oObject Is IJPlatePart And plateType = CollarPlate Then
            IsCollar = True
        End If
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsCollar").Number

End Function

'***********************************************************************
' METHOD:  GetAngleBetweenCornerEdgePorts
'
' DESCRIPTION:  Gets the angles between two edge ports of corner feature
'               (this method can also be used for getting the angle between corner feature
'                face port and any edge port)
'
' INPUTS:  Any two input ports of the corner feature
'
' OUPUT:  Angle between the two ports
'***********************************************************************
Public Function GetAngleBetweenCornerEdgePorts(oEdgePort1 As IJPort, oEdgePort2 As IJPort) As Double

    On Error GoTo ErrorHandler
    Const sMETHOD = "GetAngleBetweenCornerEdgePorts"
    
    If oEdgePort1 Is Nothing Or _
       oEdgePort2 Is Nothing Then
      
        GoTo ErrorHandler
    
    End If
    
    
    Dim oPartInfo As New PartInfo
    Dim oNormalOnEdgePort1 As IJDVector
    Dim oNormalOnEdgePort2 As IJDVector
    
    Dim bApproxmationUsed As Boolean
    
    Set oNormalOnEdgePort1 = oPartInfo.GetPortNormal(oEdgePort1, bApproxmationUsed)
    Set oNormalOnEdgePort2 = oPartInfo.GetPortNormal(oEdgePort2, bApproxmationUsed)
    
    
    
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    
    Dim dAngle As Double
    

    oNormalOnEdgePort1.Get x, y, z
    oNormalOnEdgePort1.Set -x, -y, -z
          
    oNormalOnEdgePort2.Get x, y, z
    oNormalOnEdgePort2.Set -x, -y, -z
          
    dAngle = oNormalOnEdgePort1.Angle(oNormalOnEdgePort2, oNormalOnEdgePort1)
    dAngle = GetPI - dAngle
    
    
    GetAngleBetweenCornerEdgePorts = dAngle
    
    Set oNormalOnEdgePort1 = Nothing
    Set oNormalOnEdgePort1 = Nothing
    Set oPartInfo = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'***********************************************************************
' METHOD: GetSlotAtFeature
'
' DESCRIPTION: Gets the slot feature on support1/support 2 to which input
'              corner feature is related i.e. if slot height changes, the
'              corner feature geometry changes in a corresponding manner.
'
' INPUTS: Corner feature Smart Occurrence
'
' OUPUT: Slot object
'***********************************************************************
Public Function GetSlotAtFeature(oCornerFeatureSO As Object) As Object

Const sMETHOD = "GetSlotAtFeature"
On Error GoTo ErrHandler

    Set GetSlotAtFeature = Nothing 'incase no slot is found

    Const MIN_DISTANCE = 0.0001

'   Declare variables
    Dim oStructFeature As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Dim i As Long
    Dim j As Long
    Dim iC As Long
    Dim n As Long
    Dim oObj As Object
    Dim oModelBody As IJDModelBody
    Dim dMinimumDistance As Double
    Dim oBasePort1 As IJPort
    Dim oBasePort2 As IJPort
    Dim oCornerPos As IJDPosition
    Dim oCornerFeature As IJSDOCornerFeature
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim nPort1CFACs As Long
    Dim nPort2CFACs As Long
    Dim aPort1CFACData() As ConnectionData
    Dim aPort2CFACData() As ConnectionData
    Dim nSupport1ACs As Long
    Dim nSupport2ACs As Long
    Dim aSupport1ACData() As ConnectionData
    Dim aSupport2ACData() As ConnectionData
    Dim pIJStructFeatUtils As IJSDFeatureAttributes
    Dim oReqPortOnSupp1 As IJPort
    Dim oCheckPort As IJPort
    Dim AppConnection As IJAppConnectionType
    Dim bFoundSlot As Boolean
    Dim oPart As Object
    Dim oReqPortOnSupp2 As IJPort
    
    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    Set oCornerFeature = New StructDetailObjectsex.CornerFeature
    Set oCornerFeature.object = oCornerFeatureSO
    If oCornerFeature.object Is Nothing Then
        Exit Function
    ElseIf TypeOf oCornerFeature.object Is IJStructFeature Then
        Set oStructFeature = oCornerFeature.object
        If oStructFeature.get_StructFeatureType <> SF_CornerFeature Then
            Exit Function
        End If
    Else
        Exit Function
    End If

    oCornerFeature.GetLocationOfCornerFeature oCornerPos
    ' -------------------------------------------------------------------------------
    ' Retreive base ports of support 1 and support 2
    ' -------------------------------------------------------------------------------
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set pIJStructFeatUtils = New SDFeatureUtils
    
    'Check type of part and get BasePort
    Set oPart = oCornerFeature.GetPartObject

    If TypeOf oPart Is IJProfile Then
        Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oPart
        Set oBasePort1 = oSDO_ProfilePart.BasePortBeforeTrim(BPT_Lateral)
        Set oBasePort2 = oSDO_ProfilePart.BasePortBeforeTrim(BPT_Base)
        Set oSDO_ProfilePart = Nothing
    ElseIf TypeOf oPart Is IJPlate Then
        Dim oSDO_PlatePart As StructDetailObjects.PlatePart
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oPart
        Set oBasePort1 = oSDO_PlatePart.baseport(BPT_Lateral)
        Set oBasePort2 = oSDO_PlatePart.baseport(BPT_Base)
        Set oSDO_PlatePart = Nothing
    End If
    
'   Get all connections using the Corner Feature’s port 1
    oSDO_Helper.Object_AppConnections oBasePort1, AppConnectionType_Assembly, nPort1CFACs, aPort1CFACData()

'  Loop through all connections using the Corner Feature’s port 1
'  Use the AC connection geometry and corner feature location to find the closest AC (discard other ACs)
    Set oReqPortOnSupp1 = Nothing
    For i = 1 To nPort1CFACs
        If TypeOf aPort1CFACData(i).ToConnectedPort Is IJPort Then
            Set oCheckPort = aPort1CFACData(i).ToConnectedPort
            Set oModelBody = oCheckPort.Geometry
            oModelBody.GetMinimumDistanceFromPosition oCornerPos, Nothing, dMinimumDistance
            If dMinimumDistance < MIN_DISTANCE Then
                Set oReqPortOnSupp1 = aPort1CFACData(i).ToConnectedPort
                Exit For
            End If
        End If
    Next i
    If Not oReqPortOnSupp1 Is Nothing Then
        oSDO_Helper.Object_AppConnections oReqPortOnSupp1.Connectable, AppConnectionType_Assembly, nSupport1ACs, aSupport1ACData()
    Else
        Exit Function
    End If
    Set oReqPortOnSupp1 = Nothing
    
'   Get all connections using the Corner Feature’s port 2
    oSDO_Helper.Object_AppConnections oBasePort2, AppConnectionType_Assembly, nPort2CFACs, aPort2CFACData()

'  Loop through all connections using the Corner Feature’s port 2
'  Use the AC connection geometry and corner feature location to find the closest AC (discard other ACs)
    Set oReqPortOnSupp2 = Nothing
    For i = 1 To nPort2CFACs
        If TypeOf aPort2CFACData(i).ToConnectedPort Is IJPort Then
            Set oCheckPort = aPort2CFACData(i).ToConnectedPort
            Set oModelBody = oCheckPort.Geometry
            oModelBody.GetMinimumDistanceFromPosition oCornerPos, Nothing, dMinimumDistance
            If dMinimumDistance < MIN_DISTANCE Then
                Set oReqPortOnSupp2 = aPort2CFACData(i).ToConnectedPort
                Exit For
            End If
        End If
    Next i
    If Not oReqPortOnSupp2 Is Nothing Then
        oSDO_Helper.Object_AppConnections oReqPortOnSupp2.Connectable, AppConnectionType_Assembly, nSupport2ACs, aSupport2ACData()
    Else
        Exit Function
    End If
    Set oReqPortOnSupp2 = Nothing
    Set oCornerPos = Nothing
    
    ' -------------------------------------------------------------------
    ' Get common AC between support 1 and support 2
    ' -------------------------------------------------------------------
    bFoundSlot = False
    For i = 1 To nSupport1ACs
        For j = 1 To nSupport2ACs
            If aSupport1ACData(i).ConnectingPort Is aSupport2ACData(j).ToConnectedPort Then
                Set AppConnection = aSupport1ACData(i).AppConnection
                If AppConnection.Behavior = ConnectionPenetration Then
                    Set oMemberObjects = aSupport1ACData(i).AppConnection
                    n = oMemberObjects.Count
                    For iC = 1 To n
                        Set oObj = oMemberObjects.Item(iC)
                        If TypeOf oObj Is IJStructFeature Then
                            Set oStructFeature = oObj
                            'Check for slot feature
                            If oStructFeature.get_StructFeatureType = SF_Slot Then
                                Set GetSlotAtFeature = oObj
                                bFoundSlot = True
                                Exit For
                            End If
                        End If
                    Next iC
                End If
            End If
            If bFoundSlot = True Then Exit For
        Next j
        If bFoundSlot = True Then Exit For
    Next i
    
    'Clean up
    Set oBasePort1 = Nothing
    Set oBasePort2 = Nothing
    Set oCheckPort = Nothing
    Set oCornerFeature = Nothing
    Set oMemberObjects = Nothing
    Set oModelBody = Nothing
    Set oObj = Nothing
    Set oPart = Nothing
    Set oSDO_Helper = Nothing
    Set oStructFeature = Nothing
    Set pIJStructFeatUtils = Nothing

    Exit Function
    
ErrHandler:
        Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function


'************************************************************************
' Method : Equal --- checkes whether two double variables have equal
'          values within a given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function Equal(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean
    If (LeftVariable >= (RightVariable - Tolerance)) And _
        (LeftVariable <= (RightVariable + Tolerance)) Then
        Equal = True
    Else
        Equal = False
    End If
End Function

'************************************************************************
' Method : GreaterThan --- checkes whether left side double variable is
'          greater than right side double variable, uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'
'  NOTE: GreaterThanZero method need to be used if right side variable
'        is zero '0#' for comparison
'************************************************************************
Public Function GreaterThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable > (RightVariable - Tolerance)) Then
        GreaterThan = True
    Else
        GreaterThan = False
    End If

End Function

'************************************************************************
' Method : GreaterThanZero --- checkes whether left side double variable
'          is greater than zero '0#', uses given tolerance
'
'  Inputs --------- Left side double variable and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function GreaterThanZero(ByVal LeftVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If LeftVariable > Abs(Tolerance) Then
        GreaterThanZero = True
    Else
        GreaterThanZero = False
    End If

End Function

'************************************************************************
' Method : GreaterthanOrEqualTo --- checkes whether left side double
'          variable is greater than or equal to right side double variable
'          uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function GreaterThanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable >= (RightVariable - Tolerance)) Then
        GreaterThanOrEqualTo = True
    Else
        GreaterThanOrEqualTo = False
    End If

End Function

'************************************************************************
' Method : LessThan --- checkes whether left side double variable is
'          less than right side double variable, uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'
'  NOTE: LessThanZero method need to be used if right side variable
'        is zero '0#' for comparison
'************************************************************************
Public Function LessThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable < (RightVariable + Tolerance)) Then
        LessThan = True
    Else
        LessThan = False
    End If

End Function

'************************************************************************
' Method : LessThanZero --- checkes whether left side double variable
'          is less than zero '0#', uses given tolerance
'
'  Inputs --------- Left side double variable and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************

Public Function LessThanZero(ByVal LeftVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If LeftVariable < -Abs(Tolerance) Then
        LessThanZero = True
    Else
        LessThanZero = False
    End If

End Function

'************************************************************************
' Method : LessThanOrEqualTo --- checkes whether left side double
'          variable is less than or equal to right side double variable
'          uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function LessThanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable <= (RightVariable + Tolerance)) Then
        LessThanOrEqualTo = True
    Else
        LessThanOrEqualTo = False
    End If

End Function

'************************************************************************
' Method : IsValueBetween --- checkes whether given double value lies in the given range
'
'
'  Inputs --------- One double value, two limits and Exclusive(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function IsValueBetween(dValue As Double, dLowerlimit As Double, dUpperlimit As Double, _
                            Optional bLimitsExclusive As Boolean = True, Optional ByVal Tolerance As Double = TOLERANCE_VALUE)

    IsValueBetween = False
    
    If bLimitsExclusive Then
        
        If (dValue > (dLowerlimit + Tolerance)) And (dValue < (dUpperlimit - Tolerance)) Then
            IsValueBetween = True
        End If
    Else
        If (dValue >= (dLowerlimit - Tolerance)) And (dValue <= (dUpperlimit + Tolerance)) Then
            IsValueBetween = True
        End If
    End If

End Function





Public Sub UpdateDependentEdgeFeature(oPhyConn As Object)

Const sMETHOD = "UpdateDependentEdgeFeature"
On Error GoTo ErrHandler

    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    If Not TypeOf oPhyConn Is IJStructPhysicalConnection Then
        Exit Sub
    End If

    Dim oSDOPhysConn As New StructDetailObjects.PhysicalConn
    Set oSDOPhysConn.object = oPhyConn
    
    
    ' For each part
    ' -------------
    Dim oPartSupport As IJPartSupport
    Dim oFeaturesList As Collection
    Dim oFeature As IUnknown
    Dim oStructFeature As IJStructFeature
    Dim featureType As StructFeatureTypes
    Dim oFeatureLoc As IJDPosition
    Dim bForceUpdate As Boolean
    
    Dim exitflag As Boolean
    exitflag = False
    
    Dim oWB As IJWireBody
    Dim oComplexStrings As IJElements
    Dim oCurve As IJCurve
    Dim iCount As Long
    Dim jCount As Long
    Dim kCount As Long
    Dim oPlate As IJPlate
    
    For iCount = 1 To 2
        ' Code will proceed only when object is a plate (but not hull) or a Profile/Member
        If iCount = 1 Then
            If TypeOf oSDOPhysConn.ConnectedObject1 Is IJPlate Then
                Set oPlate = oSDOPhysConn.ConnectedObject1
                If oPlate.plateType = Hull Then Exit For
            ElseIf TypeOf oSDOPhysConn.ConnectedObject1 Is IJProfile Then
            
            ElseIf TypeOf oSDOPhysConn.ConnectedObject1 Is ISPSMemberPartPrismatic Then
            
            Else
            ' Not proceeding if the Object is Neither Plate, Profile or Prismatic member
                Exit For
            End If
            
            If TypeOf oSDOPhysConn.ConnectedObject1 Is ISPSMemberPartPrismatic Then
                ' Object is Member
                Set oPartSupport = New MemberPartSupport
                Set oPartSupport.Part = oSDOPhysConn.ConnectedObject1
            Else
                ' Object is Plate or Profile
                Set oPartSupport = New PartSupport
            Set oPartSupport.Part = oSDOPhysConn.ConnectedObject1
            End If
        Else
            If TypeOf oSDOPhysConn.ConnectedObject2 Is IJPlate Then
                Set oPlate = oSDOPhysConn.ConnectedObject2
                If oPlate.plateType = Hull Then Exit For
            ElseIf TypeOf oSDOPhysConn.ConnectedObject2 Is IJProfile Then
            
            ElseIf TypeOf oSDOPhysConn.ConnectedObject2 Is ISPSMemberPartPrismatic Then
            
            Else
            ' Not proceeding if the Object is Neither Plate, Profile or Prismatic member
                Exit For
            End If

            If TypeOf oSDOPhysConn.ConnectedObject2 Is ISPSMemberPartPrismatic Then
                Set oPartSupport = New MemberPartSupport
                Set oPartSupport.Part = oSDOPhysConn.ConnectedObject2
            Else
                'oSDOPhysConn.ConnectedObject2Type
                Set oPartSupport = New PartSupport
            Set oPartSupport.Part = oSDOPhysConn.ConnectedObject2
        End If
        
        End If
        
        oPartSupport.GetFeatures oFeaturesList

        ' ----------------------------
        ' For each feature on the part
        ' ----------------------------
        Dim nFeatures As Long
        nFeatures = oFeaturesList.Count
        
        For jCount = 1 To nFeatures
            Set oFeature = oFeaturesList.Item(jCount)
            
            bForceUpdate = False
            
            ' -------------------------
            ' If edge feature
            ' -------------------------
            If TypeOf oFeature Is IJStructFeature Then
                Set oStructFeature = oFeature
                featureType = oStructFeature.get_StructFeatureType
                                
                Select Case featureType
                    Case SF_EdgeFeature
                        
                        Dim pIJStructFeatUtils As IJSDFeatureAttributes
                        Set pIJStructFeatUtils = New SDFeatureUtils
                        
                        Dim oppEdgePort As IJPort
                        Dim oppLocationPoint As Object
                        
                        pIJStructFeatUtils.Get_Inputs_EdgeCut oStructFeature, oppEdgePort, oppLocationPoint
                        
                        Dim oP3d As IngrGeom3D.Point3d
                        Set oP3d = New Point3d
                        
                        If TypeOf oppLocationPoint Is Point3d Then
                            Set oP3d = oppLocationPoint
                        Else 'Considering case when placed using seam or knuckle method
                            Dim oTempPos As IJDPosition
                            Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
                            Set oTopologyLocate = New TopologyLocate
        
                            Set oTempPos = oTopologyLocate.FindIntersectionPoint(oppEdgePort, oppLocationPoint)
                            
                            ' If no intersection point exists, knuckle is "Bent" and does directly intersect edge port
                            ' Need projected intersection point
                            If oTempPos Is Nothing Then
                                If TypeOf oppLocationPoint Is IJKnuckle Then
                                    Dim oKnuckleObject As PlateKnuckle
                                    Set oKnuckleObject = oppLocationPoint
                                    Set oTempPos = GetProjectedKnuckleIntersection(oppEdgePort, oKnuckleObject)
                                    Set oKnuckleObject = Nothing
                                End If
                            End If
                            oP3d.SetPoint oTempPos.x, oTempPos.y, oTempPos.z
                        End If
                        
                        Dim dx As Double
                        Dim dy As Double
                        Dim dz As Double
                        
                        oP3d.GetPoint dx, dy, dz
                        
                        Dim dMinDist As Double
                        Dim dSrcX As Double
                        Dim dSrcY As Double
                        Dim dSrcZ As Double
                        Dim dInX As Double
                        Dim dInY As Double
                        Dim dInZ  As Double
                        
                        If TypeOf oSDOPhysConn.object Is IJWireBody Then
                            Set oWB = oSDOPhysConn.object
        
                            oWB.GetComplexStrings Nothing, oComplexStrings
                            For kCount = 1 To oComplexStrings.Count
                                Set oCurve = oComplexStrings.Item(kCount)
                                
                                exitflag = oCurve.IsPointOn(dx, dy, dz)
                                
                                If exitflag = False Then
                                    oCurve.DistanceBetween oP3d, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
                                    'ReCompute Edge Feature if Point lies within a minimum distance of 10 mm from PC
                                    If dMinDist < 0.01 Then exitflag = True
                                End If
                                'Since point lies on PC, Update dependent Edge Feature
                                If exitflag = True Then Exit For
                            Next kCount
                        End If
                        
                        If exitflag = True Then
                            bForceUpdate = True
                        End If
                End Select
            End If
         
            If bForceUpdate Then
                ForceUpdateSmartItem oFeature
            End If
        Next jCount
    Next iCount
    
    Exit Sub

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


'Method IfSeamsExistWithInDistance --- 'This Function Return the distance between the seam and pLocation, if there is more than one seam,
'it will return the distance which is  between the furthest seam and pLocation

'Inputs: 1) root object of a plate or a profile is to be passed
'        2) location from which seams are looked for
'        3) Specified distance

'Outputs: 1) Distance of seam(that lies with in distance) from the location(pLocation)
'         2)Direction of the vector that joins the location and its projected
            'point on seam landing curve

'oDirec: pass in the profile secondery oritation
'bolNeedSameSide: True: means you want the seam which is on the Web Right Side
'                 False: Means you want the seam which is on the Web Left Side

Public Function IfSeamsExistWithInDistance(pLocation As IJDPosition, _
                                           oRootObject As Object, _
                                           dWithInDistance As Double, _
                                           dSeamDistance As Double, _
                                           Optional oDirec As IJDVector, _
                                           Optional bolNeedSameSide As Boolean = False, _
                                           Optional bSeam_Feature As Boolean, _
                                           Optional dSearchDist As Double, _
                                           Optional oCFFacePort As IJPort) As Boolean
                                           
    If pLocation Is Nothing Or oRootObject Is Nothing Then
        Exit Function
    End If
    
    IfSeamsExistWithInDistance = False
    bSeam_Feature = False
    
    Dim oJSeam As IJSeam
    Dim oWireBody As IUnknown
    Dim oLandingCurve As IJLandCurve
    Dim oSeamAttributes As IJSeamAttributes
    Dim oSeamUtils As SeamUtils
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Dim oVector As IJDVector
    Dim oPntOnWire As IJDPosition
    Dim dDistance As Double
    Dim oStructDetailHelper As New StructDetailObjects.Helper
    Dim oListOfAllSeams As Collection
    
    Dim bolDetermineSide As Boolean
    bolDetermineSide = False
    If Not oDirec Is Nothing Then
        bolDetermineSide = True
    End If
    
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    Set oSeamUtils = New SeamUtils
    Set oSeamAttributes = oSeamUtils
    
'    oStructDetailHelper.GetSeamsOnObject oRootObject, oListOfAllSeams
    Set oListOfAllSeams = New Collection
    
    
    If TypeOf oRootObject Is IJPlate Then
    
        GetSeamsOnPlate oRootObject, oListOfAllSeams
        
        dSeamDistance = 0
        
        Dim iIndex As Long
        Dim nSeams As Long
        
        nSeams = oListOfAllSeams.Count
        For iIndex = 1 To nSeams
            ' For Each Seam
            ' Retrieve it's Geometry (WireBody)
            
    
            Set oJSeam = oListOfAllSeams.Item(iIndex)
            oSeamAttributes.GetLandingCurveFromSeam oJSeam, oLandingCurve
             
            ' Retrieve the nearest Point on the Seam from the Penetration Point
            ' by Projecting the Penetration Point onto the Seam's WireBody
            Set oWireBody = oLandingCurve
            oTopologyLocate.GetProjectedPointOnModelBody oWireBody, pLocation, _
                                                  oPntOnWire, oVector
        
            Dim oTempVector As IJDVector
            Set oTempVector = oPntOnWire.Subtract(pLocation)
            'Set oVector = oTempVector
                                               
            'Determine the seam is on which side of the location. if its a profile
            'it will be decided if it is on web right or web left.
            Dim bolSameSide As Boolean
            bolSameSide = False
            If bolDetermineSide = True Then
                Dim dFlag As Double
                dFlag = oDirec.x * oTempVector.x + oDirec.y * oTempVector.y + oDirec.z * oTempVector.z
                If dFlag > 0 Then
                   bolSameSide = True
                Else
                   bolSameSide = False
                End If
            End If
                                                                                
            If (Not oVector Is Nothing) And (bolNeedSameSide = bolSameSide) Then
                If oVector.Length > 0.0001 Then
                    ' Calculate distance between the Point on Seam and Penetration Point
                    dDistance = oPntOnWire.DistPt(pLocation)
                    
                    If LessThanOrEqualTo(dDistance, dWithInDistance) Then
                        IfSeamsExistWithInDistance = True
                        If GreaterThan(dDistance, dSeamDistance) Then
                            dSeamDistance = dDistance
                        End If
                    End If
                    If GreaterThanOrEqualTo(dDistance, dWithInDistance) And LessThanOrEqualTo(dDistance, (dWithInDistance + dSearchDist)) Then
                        bSeam_Feature = True
                        dSeamDistance = dDistance
                    End If
                    If LessThanOrEqualTo(dDistance, dWithInDistance) And GreaterThanOrEqualTo(dDistance, (dWithInDistance - dSearchDist)) Then
                         bSeam_Feature = True
                        dSeamDistance = dDistance
                    End If
                End If
            End If
            
            Set oVector = Nothing
            Set oPntOnWire = Nothing
            
            Set oJSeam = Nothing
            Set oWireBody = Nothing
            Set oLandingCurve = Nothing
        Next iIndex
        
        
    ElseIf TypeOf oRootObject Is IJProfile Then
    
        Dim oKnuckleColl As IJElements
        Dim oSeamPointColl As New Collection
        
        dSeamDistance = 0
        'Get Profile Knuckles and Seam Points on Profile object
        GetKnucklesAndSeamsonProfile oRootObject, oKnuckleColl, oSeamPointColl
        
        For iIndex = 1 To oKnuckleColl.Count
            oListOfAllSeams.Add oKnuckleColl.Item(iIndex)
        Next
        
        For iIndex = 1 To oSeamPointColl.Count
            oListOfAllSeams.Add oSeamPointColl.Item(iIndex)
        Next
        
        Dim oSplitPositon As IJDPosition
        Set oSplitPositon = New DPosition
        
        'Loop through each item and find the nearest split
        For iIndex = 1 To oListOfAllSeams.Count
            
            'Get the Position of the Split
            If TypeOf oListOfAllSeams.Item(iIndex) Is IJPoint Then
                
                Dim oSplitPoint As IJPoint
                Dim dx As Double, dy As Double, dz As Double
                
                'Convert Point to Position
                Set oSplitPoint = oListOfAllSeams.Item(iIndex)
                
                oSplitPoint.GetPoint dx, dy, dz
                oSplitPositon.Set dx, dy, dz
                
            ElseIf TypeOf oListOfAllSeams.Item(iIndex) Is IJDModelBody Then
                
                Dim oModelBody As IJDModelBody
                Dim dMinDist As Double
                
                'Get the Position of the Split
                Set oModelBody = oListOfAllSeams.Item(iIndex)
                oModelBody.GetMinimumDistanceFromPosition pLocation, oSplitPositon, dMinDist
            Else
                'Not yet Handled
            End If
            
            If Not oSplitPositon Is Nothing Then
                'Vector from CF location to Seam location
                Set oVector = oSplitPositon.Subtract(pLocation)
                
                'Determine on which the split exits
                If oVector.Dot(oDirec) > 0 Then
                    bolSameSide = True
                Else
                    bolSameSide = False
                End If
                
                If bolSameSide Then
                    
                    dMinDist = 0
                    
                    Dim oProjectedPos As IJDPosition
                    Set oProjectedPos = New DPosition
                    
                    Set oModelBody = Nothing
                    Set oModelBody = oCFFacePort.Geometry
                    
                    'Get the Split Position on the Corner Feature Face Port
                    oModelBody.GetMinimumDistanceFromPosition oSplitPositon, oProjectedPos, dMinDist
                    
                    If (Not oVector Is Nothing) And (bolNeedSameSide = bolSameSide) Then
                        If oVector.Length > 0.0001 Then
                            ' Calculate distance between the Point on Seam and Penetration Point
                            dDistance = oProjectedPos.DistPt(pLocation)
                            
                            If LessThanOrEqualTo(dDistance, dWithInDistance) Then
                                IfSeamsExistWithInDistance = True
                                If GreaterThan(dDistance, dSeamDistance) Then
                                    dSeamDistance = dDistance
                                End If
                            End If
                            
                            'Check if split falls in given Range
                            If GreaterThanOrEqualTo(dDistance, dWithInDistance) And LessThanOrEqualTo(dDistance, (dWithInDistance + dSearchDist)) Then
                                bSeam_Feature = True
                                dSeamDistance = dDistance
                            End If
                            
                            'Check if split falls in given Range
                            If LessThanOrEqualTo(dDistance, dWithInDistance) And GreaterThanOrEqualTo(dDistance, (dWithInDistance - dSearchDist)) Then
                                 bSeam_Feature = True
                                dSeamDistance = dDistance
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If


CleanUp:
    Set oVector = Nothing
    Set oJSeam = Nothing
    Set oKnuckleColl = Nothing
    Set oSplitPoint = Nothing
    Set oSeamPointColl = Nothing
    Set oModelBody = Nothing
    Set oProjectedPos = Nothing


    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, "Seam.cls", "IfSeamsExistWithInDistance").Number
End Function


'GetSeamsOnPlate -- used to fetch all the seams on a given object.
'wrapper GetSeamsOnObject necessitates passing root object whereas this method can take
'input as any object. be it a root parent or a leaf plate.

Public Sub GetSeamsOnPlate(oSystemUn As IUnknown, oCollection As Collection)

  On Error GoTo ErrorHandler
  Const sMETHOD = "GetSeamsOnPlate"
  
    If (TypeOf oSystemUn Is IJPlate) And (TypeOf oSystemUn Is IJSystem) Then
      
      Dim oParent As IJSystem
      Dim lCount As Long
      Dim sChildren As String
      Dim oChildren As IJDTargetObjectCol
      
      If TypeOf oSystemUn Is IJProfile Then
          Exit Sub
      End If
      
      Set oParent = oSystemUn
      Set oChildren = oParent.GetChildren
      
      lCount = oChildren.Count
      Dim i As Integer
      For i = 1 To lCount
          Dim oObject As Object
          Set oObject = oChildren.Item(i)
          If TypeOf oObject Is IJSeam Then
              oCollection.Add oObject
          ElseIf (TypeOf oObject Is IJPlate) Then
              GetSeamsOnPlate oObject, oCollection
          End If
      Next
    End If
    
Exit Sub
ErrorHandler:
Err.Raise LogError(Err, MODULE, sMETHOD).Number


End Sub

'************************************************************************
'  Method : GetPI - returns PI value
'************************************************************************
Public Function GetPI() As Double
    GetPI = 3.14159265358979
End Function

Public Sub ForceUpdateSmartItem(oObject As Object)

On Error GoTo ErrorHandler
Const sMETHOD = "ForceUpdateSmartItem"
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute WebCut before FlangeCuts
    ' interface IJStructGeometry : {6034AD40-FA0B-11d1-B2FD-080036024603}
    Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools

    Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
    
    On Error Resume Next
    If TypeOf oObject Is IJAssemblyConnection Then
        oStructAssocTools.UpdateObject oObject, _
                                   "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}"
    Else
        oStructAssocTools.UpdateObject oObject, _
                                   "{6034AD40-FA0B-11d1-B2FD-080036024603}"
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    Set oStructAssocTools = Nothing

    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub


Public Function GetProjectedKnuckleIntersection(oEdgePort As Object, oKnuckleObject As Object) As IJDPosition

On Error GoTo ErrorHandler
Const sMETHOD = "GetProjectedKnuckleIntersection"
    
    Dim oModelBodyUtil As SGOModelBodyUtilities
    Dim oWireBodyUtil As SGOWireBodyUtilities
    Dim oPointOnWire As IJDPosition
    Dim oPointOnSurface As IJDPosition
    Dim oClosestPoint As IJDPosition
    Dim oTangentVec As IJDVector
    Dim dDistance As Double
    Dim oWireEdgePortGeom As IJWireBody
    Dim oWireEdgePort As IJPort
    
    Set oModelBodyUtil = New SGOModelBodyUtilities
    Set oWireBodyUtil = New SGOWireBodyUtilities
    Set oWireEdgePort = oEdgePort
    Set oWireEdgePortGeom = oWireEdgePort.Geometry
    oModelBodyUtil.GetClosestPointsBetweenTwoBodies oWireEdgePortGeom, _
                                           oKnuckleObject, _
                                           oPointOnWire, _
                                           oPointOnSurface, _
                                           dDistance

    ' Get the tangent vector along the wire edge port at the Position returned.
    oWireBodyUtil.GetClosestPointOnWire oWireEdgePortGeom, _
                                        oPointOnWire, _
                                        oClosestPoint, _
                                        oTangentVec
                                        
    Set GetProjectedKnuckleIntersection = oClosestPoint
    
    Exit Function
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

'********************************************************************
'  METHOD: RelatedPortBeforeCut
' DESCRIPTION:
'   Calculates the Intersection Point between two Ports on a given Plane
'
'   In:
'       oPort As IJPort - Port to retrieve prior Related Port for
'
'   Out:
'       RelatedPortBeforeCut As IJPort - Related Port prior to SDCut_AE
'
' Note:
'
'********************************************************************
Public Function RelatedPortBeforeCut(oPort As IJPort, _
                            Optional bGlobalPorts As Boolean = False) As IJPort
On Error GoTo ErrorHandler
Const MT = "RelatedPortBeforeCut"

    Const strGeneratedPlatePartAEProgId = "CreatePlatePart.GeneratePlatePart_AE.1"
    Const strStandAlonePlatePartAEProgId = "CreatePlatePart.CreatePlatePart_AE.1"
    Const strGeneratedProfilePartAEProgId = "ProfilePartActiveEntities.ProfilePartGeneration_AE.1"
    Const strStandAloneProfilePartTrimAEProgId = "ProfilePartActiveEntities.ProfileTrim_AE.1"
    Const strStandAloneProfilePartAEProgID = "ProfilePartActiveEntities.ProfilePartCreation.1"
    Const strSketchFeatureCutAEProgID = "SketchFeature.SDCutAE.1"

    Dim oStructPort As IJStructPortEx
    Set oStructPort = oPort
    
    If TypeOf oPort.Connectable Is IJPlatePart Then
        Set RelatedPortBeforeCut = oStructPort.RelatedPort(strStandAlonePlatePartAEProgId, _
                                                           False, bGlobalPorts)
        
        If (oPort Is RelatedPortBeforeCut) Then
            Set RelatedPortBeforeCut = oStructPort.RelatedPort(strGeneratedPlatePartAEProgId, _
                                                               False, bGlobalPorts)
        End If
    Else
        ' Generated Profile part with  Bound , stand alone profile part with or without boundary
        If TypeOf oPort.Connectable Is IJProfilePart Then
            Set RelatedPortBeforeCut = oStructPort.RelatedPort(strGeneratedProfilePartAEProgId, _
                                                            False, bGlobalPorts)
             
            ' Generated profile part without Boundary
            If (oPort Is RelatedPortBeforeCut) Then
                Set RelatedPortBeforeCut = oStructPort.RelatedPort(strGeneratedProfilePartAEProgId, _
                                                            False, bGlobalPorts)
            End If
            
            If (oPort Is RelatedPortBeforeCut) Then
                'StandAlone not sure if this is required, it does not harm though
                Set RelatedPortBeforeCut = oStructPort.RelatedPort(strStandAloneProfilePartAEProgID, _
                                                            False, bGlobalPorts)
            End If
            
            ' This may not be required
            If (oPort Is RelatedPortBeforeCut) Then
                Set RelatedPortBeforeCut = oStructPort.RelatedPort(strSketchFeatureCutAEProgID, _
                                                            True, bGlobalPorts)
            End If
        End If
    End If
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function



Public Sub GetOtherPortAndConnectable(oThisPort As IJPort, oOtherPort As IJPort, oOtherConnectable As IJConnectable)
On Error GoTo ErrorHandler

    Dim ppEnumConnections As IJElements
    Dim oStructconnection As IMSStructConnection.IJStructConnection
    Dim oStructDetailHelper As New StructDetailObjects.Helper

    oThisPort.enumConnections ppEnumConnections, ConnectionAssembly, ConnectionStandard

    If Not ppEnumConnections Is Nothing Then
        Dim nCount As Long
        Dim nIndex As Long
        Dim sFlag As Long
        nCount = ppEnumConnections.Count
        If nCount = 1 Then
            Set oStructconnection = ppEnumConnections.Item(nCount)
            oStructconnection.GetOtherPortAndConnectable oThisPort, oOtherPort, oOtherConnectable
        ElseIf nCount > 1 Then ' to be enhanced if it has more than 2 connections
            Set oStructconnection = ppEnumConnections.Item(1)
            oStructconnection.GetOtherPortAndConnectable oThisPort, oOtherPort, oOtherConnectable
        End If
    End If
Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "GetOtherPortAndConnectable").Number
End Sub

'Pass an AC, see if the connected object1 and object2 are mutually bounded
Public Function IsMutualBound(oAC As StructDetailObjects.AssemblyConn) As Boolean
    On Error GoTo ErrorHandler
    Const METHOD_NAME = "IsMutualBound"
    
    IsMutualBound = False
    
    Dim pObject1 As Object
    Dim pObject2 As Object
    
    Set pObject1 = oAC.ConnectedObject1
    Set pObject2 = oAC.ConnectedObject2
    
    Dim oHelper As New StructDetailObjects.Helper
    Set pObject1 = oHelper.Object_RootParentSystem(pObject1)
    Set pObject2 = oHelper.Object_RootParentSystem(pObject2)
    
    If oHelper.ObjectType(pObject1) = SDOBJECT_BEAM Or oHelper.ObjectType(pObject1) = SDOBJECT_STIFFENER Or oHelper.ObjectType(pObject1) = SDOBJECT_STIFFENERSYSTEM Then
        If oHelper.ObjectType(pObject2) = SDOBJECT_BEAM Or oHelper.ObjectType(pObject2) = SDOBJECT_STIFFENER Or oHelper.ObjectType(pObject2) = SDOBJECT_STIFFENERSYSTEM Then
            'proceed
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    Dim iPort1TopoContext As IMSStructConnection.eUSER_CTX_FLAGS
    Dim iPort1CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Call oAC.Port1Topology(iPort1TopoContext, iPort1CrossSectEntity)
    
    
    Dim iPort2TopoContext As IMSStructConnection.eUSER_CTX_FLAGS
    Dim iPort2CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Call oAC.Port2Topology(iPort2TopoContext, iPort2CrossSectEntity)
    
    If iPort1TopoContext = CTX_LATERAL Or iPort2TopoContext = CTX_LATERAL Then
        IsMutualBound = False
    ElseIf (iPort1TopoContext = CTX_BASE Or iPort1TopoContext = CTX_OFFSET) And (iPort2TopoContext = CTX_BASE Or iPort2TopoContext = CTX_OFFSET) Then
        IsMutualBound = True
    Else
        IsMutualBound = False
    End If
    
    If IsMutualBound = True Then
        If pObject1 Is pObject2 Then    'to avoid split, knuckled profiles
            IsMutualBound = False
        End If
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
End Function

'************************************************************************
' GetProfileKnuckleType: given AC/WebCut/FlangeCut get profile knuckle's
' manufacturing method and knuckle object.
'************************************************************************
Public Sub GetProfileKnuckleType(ByVal oObject As Object, iKnuckleMfgMethod As Integer, _
                        Optional oKnuckleObject As Object = Nothing, Optional bIsConvex As Boolean = False)

  On Error GoTo ErrorHandler
  Const METHOD_NAME = "GetProfileKnuckleType"

    Dim bFromKnuckle As Boolean
    bFromKnuckle = False
    Dim oProfileKnuckle As IJProfileKnuckle

    Dim oSD_AssemblyConn As StructDetailObjects.AssemblyConn
    Dim oWebCut As StructDetailObjects.WebCut
    Dim oFlangeCut As StructDetailObjects.FlangeCut
    Dim oProfileKnuckleMfg As IJProfileKnuckleMfg
    Dim oFeature As IJStructFeature
    Dim oProfileKnuHelper As IJProfileKnuckleHelper

    If TypeOf oObject Is IJAssemblyConnection Then
        Set oSD_AssemblyConn = New StructDetailObjects.AssemblyConn
        Set oSD_AssemblyConn.object = oObject
        bFromKnuckle = oSD_AssemblyConn.FromKnuckle(oProfileKnuckle)

    ElseIf TypeOf oObject Is IJStructFeature Then
        Set oFeature = oObject
        If oFeature.get_StructFeatureType = SF_WebCut Then
            Set oWebCut = New StructDetailObjects.WebCut
            Set oWebCut.object = oObject
            bFromKnuckle = oWebCut.FromKnuckle(oProfileKnuckle)

        ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
            Set oFlangeCut = New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oObject
            bFromKnuckle = oFlangeCut.FromKnuckle(oProfileKnuckle)
        End If
    End If

    If bFromKnuckle Then
        Set oProfileKnuckleMfg = oProfileKnuckle
        iKnuckleMfgMethod = oProfileKnuckleMfg.ManufacturingMethod
        Set oKnuckleObject = oProfileKnuckle
        Set oProfileKnuHelper = New IMSProfileKnuckleEntity.ProfileKnuckleHelper
        bIsConvex = oProfileKnuHelper.IsConvexKnuckle(oProfileKnuckle)
    Else
        iKnuckleMfgMethod = -1
    End If

  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetProfileKnuckleType").Number
End Sub
'***********************************************************************************************
'    Function      : IsFreeEndCutFromExtend_ConvexKnuckle
'
'    Description   : This method determines if Free EndCut is because of 'Extend' Mfg. option and
'                    returs plate port on which  stiffener is placed at the knuckle position).
'
'    Parameters    :
'          Input    Bounded port of Free EndCut
'
'          Outputs: optionally plate's port is returned as IJPort
'
'    Return        : 'True' if following two conditions are satisfied: (i) this FreeEndCut is from 'Extend' Mfg. option and
'                    (ii) stiffened plate's port is obtained (this will be used as bounding port).
'                    otherwise 'False'
'
'***********************************************************************************************
Public Function IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedPort As Object, _
        Optional oBoundingPlatePort As IJPort = Nothing) As Boolean
    On Error GoTo ErrorHandler
    Dim strError As String
    Dim bIsBase As Boolean
    
    IsFreeEndCutFromExtend_ConvexKnuckle = False 'Initialize
    
    'Check if input bounded object is a port
    Dim oPort As IJPort
    If oBoundedPort Is Nothing Then
        Exit Function ' *** Exit ***
    Else
        If TypeOf oBoundedPort Is IJPort Then
            Set oPort = oBoundedPort
            'Continue only if it is a profile
            If Not TypeOf oPort.Connectable Is IJStiffener Then
                Exit Function ' *** Exit ***
            End If
        Else
            'For 'Extend' Mfg. option we get either base port or offset port of the stiffener as
            ' bounded object - current case is not that,so exit.
            Exit Function ' *** Exit ***
        End If
    End If
    
    'Declare variables
    Dim oSDU_FreeECUtil As GSCADStructDetailUtil.StructDetailFreeEndUtil
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oProfileKnuHelper As IMSProfileKnuckleEntity.ProfileKnuckleHelper
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    Set oPort = oBoundedPort

    bIsBase = False 'Initialize
    Set oSDO_ProfilePart.object = oPort.Connectable
    
    If oBoundedPort Is oSDO_ProfilePart.BasePortBeforeTrim(BPT_Base) Then
        bIsBase = True
    End If
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim bIsFromExtend As Boolean
    Dim oKnuckleObj As Object
    Dim oExtendedPos As Object
    Dim oKnucklePoint As IJPoint
    Dim oKnucklePosition As IJDPosition
    Dim oKnuckleHelper As IMSProfileKnuckleEntity.ProfileKnuckleHelper
    
    Set oKnuckleHelper = New IMSProfileKnuckleEntity.ProfileKnuckleHelper
    Set oKnucklePosition = New DPosition
    Set oSDU_FreeECUtil = New StructDetailFreeEndUtil
    
    'Get knuckle position
    oSDU_FreeECUtil.IsStiffEndFromExtendKnuckleEx oSDO_ProfilePart.object, bIsBase, _
                                                 bIsFromExtend, oKnuckleObj, oExtendedPos
                            
    If bIsFromExtend Then
        Dim oProfileKnuckle As IJProfileKnuckle
        Set oProfileKnuckle = oKnuckleObj
        'Check the knuckle is Psedo or not
        If oProfileKnuckle.IsPseudo Then
            Set oKnucklePoint = oKnuckleObj
            oKnucklePoint.GetPoint dx, dy, dz
            oKnucklePosition.x = dx
            oKnucklePosition.y = dy
            oKnucklePosition.z = dz
    
            If oKnuckleHelper.IsConvexKnuckle(oKnuckleObj) Then
                Set oBoundingPlatePort = GetStiffnedPlatePort_At_Position(oSDO_ProfilePart.object, _
                                                                        oKnucklePosition)
                IsFreeEndCutFromExtend_ConvexKnuckle = True
            End If
        End If
    End If

    
CleanUp:
    Set oPort = Nothing
    Set oSDO_ProfilePart = Nothing
    Set oSDU_FreeECUtil = Nothing
    Set oStructDetailHelper = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number

End Function

'***********************************************************************************************
'    Function      : GetStiffnedPlatePort_At_Position
'
'    Description   : This method returns the Plate part's base or offset port on which profile part
'                    is layed at the given location.
'
'    Parameters    : oprofilePart or port, Position Location
'          Input
'
'    Return        : Plate port as IJPort
'
'***********************************************************************************************

Public Function GetStiffnedPlatePort_At_Position(oProfilePart As Object, oPositon As IJDPosition) As IJPort
    
    On Error GoTo ErrorHandler

    Set GetStiffnedPlatePort_At_Position = Nothing
    
    'Declare variables
    Dim oProfilehelper As IJProfileAttributes
    Dim oPrimOrien As IJDVector
    Dim oSecOrien As IJDVector
    Dim oPlateAttributes As New GSCADCreateModifyUtilities.USSHelper
    Dim oSD_PlatePart As New StructDetailObjects.PlatePart
    Dim oSDO_ProfilePart As New StructDetailObjects.ProfilePart
    Dim oPlatePartColl As IJDTargetObjectCol
    Dim oMountingFace As IJSurfaceBody
    Dim oPlateModelBody As IJDModelBody
    Dim oPlateSurface As IJSurfaceBody
    Dim oClosestPos As IJDPosition
    Dim oPlateNormal  As IJDVector
    Dim oLeafPlate As IJPlate
    Dim oSystem As IJSystem
    Dim oPlateSystem As Object
    Dim oStiffenedPlate As Object
    Dim oObject As Object
    Dim bIsSplitted As Boolean
    Dim bIsPlateSystem As Boolean
    Dim iCount As Integer
    Dim dMinDistance As Double
    Dim oModelBody As IJDModelBody
    
    Set oSDO_ProfilePart.object = oProfilePart
    oSDO_ProfilePart.GetStiffenedPlate oPlateSystem, bIsPlateSystem
    
    'Check Plate is Standalone or not and Select the PlatePart of Knuckle
    If bIsPlateSystem = True Then
        'Get Plate at given position
        oPlateAttributes.GetLeafPlateNearestToPoint oPlateSystem, oPositon, _
                                                    oLeafPlate, bIsSplitted
        Set oSystem = oLeafPlate
        Set oPlatePartColl = oSystem.GetChildren
        
        For iCount = 1 To oPlatePartColl.Count
            Set oObject = oPlatePartColl.Item(iCount)
            If TypeOf oObject Is IJPlatePart Then
                Set oModelBody = oObject
                oModelBody.GetMinimumDistanceFromPosition oPositon, oClosestPos, dMinDistance
                If dMinDistance < 0.001 Then
                    Set oStiffenedPlate = oObject
                    Exit For
                End If
            End If
        Next
    Else
        Set oStiffenedPlate = oPlateSystem
    End If
      
    Set oClosestPos = Nothing
    dMinDistance = 0
    'Get plate part
    Set oSD_PlatePart.object = oStiffenedPlate
    
    Set oProfilehelper = New ProfileUtils
    oProfilehelper.GetProfileOrientation oProfilePart, oPositon, oSecOrien, oPrimOrien
    Set oMountingFace = oSDO_ProfilePart.MountingFacePort.Geometry
    
    ' Check with Base port of plate
    Set oPlateModelBody = oSD_PlatePart.baseport(BPT_Base).Geometry
    oPlateModelBody.GetMinimumDistanceFromPosition oPositon, oClosestPos, dMinDistance
    Set oPlateSurface = oSD_PlatePart.baseport(BPT_Base).Geometry
    oPlateSurface.GetNormalFromPosition oClosestPos, oPlateNormal
    
    'Check both vectors are on same direction then required port is base
    If oPlateNormal.Dot(oPrimOrien) > 0 Then
        Set GetStiffnedPlatePort_At_Position = oSD_PlatePart.baseport(BPT_Base)
        GoTo CleanUp
    Else
        Set GetStiffnedPlatePort_At_Position = oSD_PlatePart.baseport(BPT_Offset)
    End If
  Exit Function
    
CleanUp:

    Set oProfilehelper = Nothing
    Set oModelBody = Nothing
    Set oPrimOrien = Nothing
    Set oSecOrien = Nothing
    Set oPlateAttributes = Nothing
    Set oSD_PlatePart = Nothing
    Set oSDO_ProfilePart = Nothing
    Set oPlatePartColl = Nothing
    Set oMountingFace = Nothing
    Set oPlateModelBody = Nothing
    Set oPlateSurface = Nothing
    Set oClosestPos = Nothing
    Set oPlateNormal = Nothing
    Set oLeafPlate = Nothing
    Set oSystem = Nothing
    Set oPlateSystem = Nothing
    Set oStiffenedPlate = Nothing

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetStiffnedPlatePort_At_Position", "").Number

End Function

'*************************************************************************
'Function
'   GetFeatureContourRngBox
'
'Abstract
'   Given the corner feature, this function returns rangebox of the projected feature contour
'
'Inputs
'   oCornerFeature As Object
'   ogCFRngBox As GBox
'
'Return
'   Rangebox
'
'Exceptions
'
'***************************************************************************
Public Function GetFeatureContourRngBox(oCornerFeature As Object) As GBox
  On Error GoTo ErrorHandler

    Dim sMETHOD As String
    sMETHOD = "GetFeatureContourRngBox"
    
    'Note: for scallop we have just 10 mm offset below the support 1/2 so we would need to augment
    ' the rangebox along resultant vector to determine AC below CF cases correctly.
    Dim oStructFeatUtils As IJSDFeatureAttributes
    Dim oFacePort As Object
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    
    Set oStructFeatUtils = New SDFeatureUtils
      
    oStructFeatUtils.get_CornerCutInputsEx oCornerFeature, _
                                         oFacePort, _
                                         oEdgePort1, _
                                         oEdgePort2
    'get the PCs that both edge ports are involved and recompute them
    If Not TypeOf oEdgePort1 Is IJPort Or Not TypeOf oEdgePort2 Is IJPort Then
       Exit Function
    End If
    
    Dim oPartInfo As PartInfo
    Set oPartInfo = New PartInfo
    Dim oEdgePort1NormalVec As IJDVector
    Dim oEdgePort2NormalVec As IJDVector
    Dim oResultantVec As IJDVector
    Dim bApproimationUsed As Boolean
    
    'Since normals of edge-port 1 and edge-port 2 are in direction pointing away from
    ' bounded part, resultant vector would increase size of the rangebox in this direction
    Set oEdgePort1NormalVec = oPartInfo.GetPortNormal(oEdgePort1, bApproimationUsed)
    Set oEdgePort2NormalVec = oPartInfo.GetPortNormal(oEdgePort2, bApproimationUsed)
    Set oResultantVec = oEdgePort1NormalVec.Add(oEdgePort2NormalVec)
    
    Dim dXmin As Double
    Dim dYmin As Double
    Dim dZmin As Double
    Dim dXmax As Double
    Dim dYmax As Double
    Dim dZmax As Double
    
    'Note: below condition would set value of either (i) dXmin  or (ii) dXmax.
    ' So dXmax=0 or dXmin=0 respectively; similar logic is used to store Y and Z values.
    If Sgn(oResultantVec.x) = -1 Then
        dXmin = oResultantVec.x
    Else
        dXmax = oResultantVec.x
    End If
    
    If Sgn(oResultantVec.y) = -1 Then
        dYmin = oResultantVec.y
    Else
        dYmax = oResultantVec.y
    End If
        
    If Sgn(oResultantVec.z) = -1 Then
        dZmin = oResultantVec.z
    Else
        dZmax = oResultantVec.z
    End If
    
    'Get corner feature contour
    Dim oStGenContour As IJStructGenericContour
    Set oStGenContour = oCornerFeature
    
    'Get contour wire body and its range
    Dim oWireBdy As IJWireBody
    Dim oRange As IJRangeAlias
    Dim gWireBdyRngBox As GBox
    
    oStGenContour.GetAttributedWire oWireBdy
    Set oRange = oWireBdy
    gWireBdyRngBox = oRange.GetRange()
    
    'Get CF contour projection vector
    Dim oElems As IJElements
    Dim oProjVec As IJDVector
    
    oStGenContour.GetProjectionVectors oElems
    Set oProjVec = oElems.Item(1)
    oProjVec.Length = oStGenContour.MaximumExtrusionDistance 'Update to required extrusion
    
    'Adjust the range for extrusion of contour
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = Abs(oProjVec.x)
    dy = Abs(oProjVec.y)
    dz = Abs(oProjVec.z)

    GetFeatureContourRngBox.m_low.x = gWireBdyRngBox.m_low.x + dXmin - dx
    GetFeatureContourRngBox.m_low.y = gWireBdyRngBox.m_low.y + dYmin - dy
    GetFeatureContourRngBox.m_low.z = gWireBdyRngBox.m_low.z + dZmin - dz
    
    GetFeatureContourRngBox.m_high.x = gWireBdyRngBox.m_high.x + dXmax + dx
    GetFeatureContourRngBox.m_high.y = gWireBdyRngBox.m_high.y + dYmax + dy
    GetFeatureContourRngBox.m_high.z = gWireBdyRngBox.m_high.z + dZmax + dz
    
CleanUp:
    Set oRange = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function


'***********************************************************************************************
'    Function      : GetKnucklesAndSeamsonProfile
'
'    Description   : This method helps to Return the Profile Knuckles and Seam points on the Profile system
'
'
'***********************************************************************************************
Public Sub GetKnucklesAndSeamsonProfile(oProfileSystem As Object, oProfileKnucklesColl As IJElements, oSeamPointColl As Collection)

 Const sMETHOD = "GetKnucklesAndSeamsonProfile"
 On Error GoTo ErrHandler

    If oProfileSystem Is Nothing Then
        Exit Sub
    End If
        
    Dim oProfileAttributes As GSCADProfilePartSemanticsLib.ProfilePartSemanticsUtil
    Set oProfileAttributes = New GSCADProfilePartSemanticsLib.ProfilePartSemanticsUtil
    
    Dim oStructUtils As IJStructEntityUtils
    Set oStructUtils = New StructEntityUtils
    
    'Get All the split Knuckles on the Profile
    oProfileAttributes.GetProfileKnucklesForPart oProfileSystem, pkmmSplit, oProfileKnucklesColl
    
    Dim oPlanningSeamColl As New Collection
    
    'Get All the Seam points on the Profile
    Set oSeamPointColl = oStructUtils.GetSeamPointsOnProfile(oProfileSystem, sstDesignSeam)
    
    Set oPlanningSeamColl = oStructUtils.GetSeamPointsOnProfile(oProfileSystem, sstPlanningSeam)
    
    Dim iCount As Integer
    
    For iCount = 1 To oPlanningSeamColl.Count
        oSeamPointColl.Add oPlanningSeamColl.Item(iCount)
    Next
          
    
Exit Sub

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Function:   degreeToRadian
'   Inputs:
'               Degree
'
'   Outputs:
'               Radian
'
'   Description:
'               This function converts degrees into radians
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function degreeToRadian(degree As Double) As Double
    Dim PI As Double
    
    PI = 4 * Atn(1)
    degreeToRadian = (degree * PI) / 180
    
End Function
