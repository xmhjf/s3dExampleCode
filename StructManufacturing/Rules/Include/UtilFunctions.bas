Attribute VB_Name = "UtilFunctions"
'*******************************************************************
'  Copyright (C) 2004 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    UtilFunctions.bas
'
'  History:
'       MJV         August 17, 2004  Initial release
'
'
'******************************************************************
Option Explicit
Private Const sSOURCEFILE As String = "\UtilFunctions.bas"
Public Const C_TOL001 = 0.001

'********************************************************************
' Routine: GetOrderPortsList
'
' Description: GetOrderPortsList
'       Given a List of Edge Ports
'       Determine the Order of the Edges Ports so that they are continuous
'       with the First Edge being the Edge with minimum X,|Y|,Z
'
' Inputs:
'   oPortList   - Edge Port List ( IJPort objects)
'
' Outputs:
'   oOrderIndex   - Edge Port Index List ( Long Index Values into the Edge Port List)
'
'
' Notes:
'   Minimum (X,|Y|,Z) Point is defined as:
'       Sort Point in ascending order by their X-coord,
'       then absolute value of y-coord, then z-coord.
'
'********************************************************************
Public Sub GetOrderPortsList(oPortList As IJElements, _
                             oOrderIndex() As Long)
Const MT = "PortUtilities.GetOrderPortsList"
On Error GoTo ErrorHandler


    
    Dim nPorts As Long
    Dim iIndex1 As Long
    Dim iIndex2 As Long
    Dim iIndex3 As Long
    Dim iMinIndex As Long
    Dim iMinOrder As Long
    
    Dim dValue1 As Double
    Dim dValue2 As Double
    Dim dValue3 As Double
    
    Dim bContinue As Boolean
    
    Dim aOrderIndex() As Long
    Dim aOrderFlags() As Long
    Dim aPortXvalue() As Double
    Dim aPortYvalue() As Double
    Dim aPortZvalue() As Double
    Dim aPortMinXvalue() As Double
    Dim aPortMinYvalue() As Double
    Dim aPortMinZvalue() As Double
    Dim aPortMaxXvalue() As Double
    Dim aPortMaxYvalue() As Double
    Dim aPortMaxZvalue() As Double
    
    Dim oPort1 As IJPort
    Dim oport2 As IJPort
    Dim oLowPos As IJDPosition
    Dim oHighPos As IJDPosition
    Dim oIntersectObject1 As IUnknown
    
    Dim oNameUtil As SDNameRulesUtilHelper
    Dim oIntersect As IIntersect
    Dim oTopologyIntersect As IJDTopologyIntersect
    
    On Error Resume Next
    
    nPorts = oPortList.Count
    ReDim oOrderIndex(nPorts)
    
    Dim sText As String
    sText = "nPorts = " & Str(nPorts)
    
    Set oNameUtil = New GSCADSDNameRulesUtil.SDNameRulesUtilHelper
    Set oIntersect = New Intersect
    Set oTopologyIntersect = New DGeomOpsIntersect
    
    ' Initialize data arrays for each Port
    ' For each Edge Port,
    '   retieve it's min/max Range Points
    '   calculate middle of Port's Range Box
    ReDim aOrderIndex(nPorts)
    ReDim aOrderFlags(nPorts)
    ReDim aPortXvalue(nPorts)
    ReDim aPortYvalue(nPorts)
    ReDim aPortZvalue(nPorts)
    ReDim aPortMinXvalue(nPorts)
    ReDim aPortMinYvalue(nPorts)
    ReDim aPortMinZvalue(nPorts)
    ReDim aPortMaxXvalue(nPorts)
    ReDim aPortMaxYvalue(nPorts)
    ReDim aPortMaxZvalue(nPorts)
    
    For iIndex1 = 1 To nPorts
        aOrderFlags(iIndex1) = 0
        aOrderIndex(iIndex1) = 0
        oNameUtil.GetRangeCorners oPortList.Item(iIndex1), oLowPos, oHighPos
        aPortMinXvalue(iIndex1) = oLowPos.x
        aPortMinYvalue(iIndex1) = oLowPos.y
        aPortMinZvalue(iIndex1) = oLowPos.z
        
        aPortMaxXvalue(iIndex1) = oHighPos.x
        aPortMaxYvalue(iIndex1) = oHighPos.y
        aPortMaxZvalue(iIndex1) = oHighPos.z
        
        aPortXvalue(iIndex1) = (oLowPos.x + oHighPos.x) / 2#
        aPortYvalue(iIndex1) = (oLowPos.y + oHighPos.y) / 2#
        aPortZvalue(iIndex1) = (oLowPos.z + oHighPos.z) / 2#
        
        Set oLowPos = Nothing
        Set oHighPos = Nothing
    Next iIndex1
    
    ' Starting with the first Port Edge
    ' Find an Edge that intersects the First Port Edge
    iIndex1 = 1
    bContinue = True
    While bContinue
    
        iIndex3 = 0
        bContinue = False
        aOrderFlags(iIndex1) = -1
        Set oPort1 = oPortList.Item(iIndex1)
        For iIndex2 = 1 To nPorts
            ' Skip the current Edge Port if it has already been Ordered
            If aOrderFlags(iIndex2) = 0 Then
                iIndex3 = iIndex2
                Set oport2 = oPortList.Item(iIndex2)
                oIntersect.GetCommonGeometry oPort1.Geometry, _
                                              oport2.Geometry, _
                                              oIntersectObject1, _
                                              False
                If oIntersectObject1 Is Nothing Then
                    ' Expect the GetCommonGeometry to return valid geometry
                    ' from adjacent Ports but if not,
                    ' see if the Ports intersect
                    oTopologyIntersect.PlaceIntersectionObject Nothing, _
                                                            oPort1.Geometry, _
                                                            oport2.Geometry, _
                                                            Null, _
                                                            oIntersectObject1
                End If
            
                Set oport2 = Nothing
                ' if the current Port Edge (Index2) intersects the Base Port Edge (Index1)
                '   set the Ordered flag for the Base Port Edge to
                '   point to the current Port Edge
                '   Use current Port Edge as the next Base Port Edge
                If Not oIntersectObject1 Is Nothing Then
                    Set oIntersectObject1 = Nothing
                    aOrderFlags(iIndex1) = iIndex2
                    iIndex1 = iIndex2
                    bContinue = True
                    Exit For
                End If
            End If
        Next iIndex2
        
        Set oPort1 = Nothing
        
    Wend
            
    ' Verify That all Port Edges have been Ordered
    bContinue = True
    For iIndex1 = 1 To nPorts
        If aOrderFlags(iIndex1) = 0 Then
            ' not all Port Edges have been ordered
            ' default to input order
            bContinue = False
        
            For iIndex2 = 1 To nPorts
                oOrderIndex(iIndex2) = iIndex2
            Next iIndex2
        
            Exit Sub
        End If
    Next iIndex1
    
    ' Determine the Port Edge at Minimum (X,|Y|,Z) Point
    iIndex1 = 1
    iIndex2 = 1
    iMinIndex = 1
    iMinOrder = 1
    aOrderIndex(iIndex2) = iIndex1
    Set oLowPos = New AutoMath.DPosition
    Set oHighPos = New AutoMath.DPosition
        
    bContinue = True
    While bContinue
        bContinue = False
        iIndex3 = aOrderFlags(iIndex1)
        If iIndex3 > 0 Then
            iIndex2 = iIndex2 + 1
            aOrderIndex(iIndex2) = iIndex3
            oLowPos.Set aPortMinXvalue(iMinOrder), _
                        aPortMinYvalue(iMinOrder), aPortMinZvalue(iMinOrder)
            oHighPos.Set aPortMinXvalue(iIndex3), _
                         aPortMinYvalue(iIndex3), aPortMinZvalue(iIndex3)
            dValue1 = oLowPos.DistPt(oHighPos)
            
            ' Check if current Edge
            '   has minimum X value / has minimum abs(Y) value / has minimum Z value
            If dValue1 > C_TOL001 Then
                If IsMinimumPoint(oHighPos, oLowPos) Then
                    iMinIndex = iIndex2
                    iMinOrder = iIndex3
                End If
            Else
                ' The current Edges are at the same minimum Corner
                ' want the Edge that has least change in Y direction
                dValue1 = aPortMaxYvalue(iIndex3) - aPortMinYvalue(iIndex3)
                dValue2 = aPortMaxYvalue(iMinOrder) - aPortMinYvalue(iMinOrder)
                dValue3 = Abs(dValue1) - Abs(dValue2)
                If dValue3 >= C_TOL001 And dValue1 < dValue2 Then
                    iMinIndex = iIndex2
                    iMinOrder = iIndex3
                Else
                    ' The current Edges are at the same minimum Corner
                    ' The current Edges have identical change in Y direction
                    ' want the Edge at minimum Z middle value
                    dValue1 = Abs(aPortZvalue(iIndex3)) - Abs(aPortZvalue(iMinOrder))
                    If dValue1 >= C_TOL001 And _
                        aPortZvalue(iIndex3) < aPortZvalue(iMinOrder) Then
                        iMinIndex = iIndex2
                        iMinOrder = iIndex3
                    Else
                        dValue1 = Abs(aPortXvalue(iIndex3)) - Abs(aPortXvalue(iMinOrder))
                        If dValue1 >= C_TOL001 And _
                            aPortXvalue(iIndex3) < aPortXvalue(iMinOrder) Then
                            iMinIndex = iIndex2
                            iMinOrder = iIndex3
                        End If
                    End If
                End If
                
            End If
            
            bContinue = True
            iIndex1 = iIndex3
        End If
    Wend
        
    ' Have determined Edge Port with Minimum X,|Y|,Z values
    ' Fill the OrderIndex array with Port Indexes indicating the Order
    iIndex2 = 0
    If iMinIndex < nPorts Then
        For iIndex1 = iMinIndex To nPorts
            iIndex2 = iIndex2 + 1
            oOrderIndex(iIndex2) = aOrderIndex(iIndex1)
        Next iIndex1
    End If
        
    If iMinIndex > 1 Then
        For iIndex1 = 1 To iMinIndex - 1
            iIndex2 = iIndex2 + 1
            oOrderIndex(iIndex2) = aOrderIndex(iIndex1)
        Next iIndex1
    End If
        
        
    ' Determine if the Port Edges have been ordered in 'anti- Clockwise' (Left Hand)
    ' or ordered in 'Clockwise' (Right hand) order
    ' Return the List in 'Clockwise' (Right Hand) Order
    '$$$
    '$$$ the code below kinda works for determining the Travel of Direction
    '$$$ (ClockWise .vs. Counter Clockwise)
    '$$$ it could be improved and should be improved when time permits or with TR
    
    Dim dDot1 As Double
    Dim bReverse As Boolean
    
    Dim oCross2 As IJDVector
    Dim oVector12 As IJDVector
    Dim oVector23 As IJDVector
    Dim oPostion1 As IJDPosition
    Dim oPostion2 As IJDPosition
    Dim oPostion3 As IJDPosition
        
    Set oPostion1 = New AutoMath.DPosition
    Set oPostion2 = New AutoMath.DPosition
    Set oPostion3 = New AutoMath.DPosition
        
    dDot1 = 0#
    bReverse = False
    
    iIndex2 = oOrderIndex(1)
    oPostion1.Set aPortXvalue(iIndex2), aPortYvalue(iIndex2), aPortZvalue(iIndex2)
    
    iIndex2 = oOrderIndex(2)
    oPostion2.Set aPortXvalue(iIndex2), aPortYvalue(iIndex2), aPortZvalue(iIndex2)
    
    For iIndex1 = 3 To nPorts + 2
        If iIndex1 > nPorts Then
            iIndex2 = oOrderIndex(iIndex1 - nPorts)
        Else
            iIndex2 = oOrderIndex(iIndex1)
        End If
        
        oPostion3.Set aPortXvalue(iIndex2), aPortYvalue(iIndex2), aPortZvalue(iIndex2)
            
        Set oVector12 = oPostion2.Subtract(oPostion1)
        Set oVector23 = oPostion3.Subtract(oPostion2)
            
        oVector12.Length = 1#
        oVector23.Length = 1#
            
        ' if the Cross Product is mainly Positive
        ' Assume Counter ColckWise direction, want Clockwise direction
        Set oCross2 = oVector12.Cross(oVector23)
        oCross2 = 1#
        If Abs(oCross2.z) > Abs(oCross2.x) Then
            If Abs(oCross2.z) > Abs(oCross2.y) Then
                If oCross2.z >= C_TOL001 Then
                    bReverse = True
                    Exit For
                End If
            Else
                If oCross2.y >= C_TOL001 Then
                    bReverse = True
                    Exit For
                End If
            End If
        ElseIf Abs(oCross2.y) > Abs(oCross2.x) Then
            If oCross2.y >= C_TOL001 Then
                bReverse = True
                Exit For
            End If
        Else
            If oCross2.x >= C_TOL001 Then
                bReverse = True
                Exit For
            End If
        End If
        
        Set oCross2 = Nothing
        oPostion1.Set oPostion2.x, oPostion2.y, oPostion2.z
        oPostion2.Set oPostion3.x, oPostion3.y, oPostion3.z
    Next iIndex1
        
    If bReverse Then
        For iIndex1 = 1 To nPorts
            aOrderFlags(nPorts - iIndex1 + 1) = oOrderIndex(iIndex1)
        Next iIndex1
        
        For iIndex1 = 1 To nPorts
            oOrderIndex(iIndex1) = aOrderFlags(iIndex1)
        Next iIndex1
    End If
    
    Exit Sub
    
ErrorHandler:
   Err.Raise LogError(Err, sSOURCEFILE, MT).Number
            
End Sub
'***********************************************************************
' METHOD:  IsMinimumPoint
'
' DESCRIPTION:  Returns the Location of a IJStructFeature object.
'   inputs:
'           oPointToCheck   - Point to Check
'           oMinPoint       - Base Point (current Minimum Point)
'
'   outputs:
'           IsMinimumPoint  - TRUE
'                               if PointToCheck is Minimum Point
'                             FALSE
'                               if oMinPoint is Minimum Point
'
'Note:
'   Minimum (X,|Y|,Z) Point is defined as:
'       Sort Point in ascending order by their X-coord,
'       then absolute value of y-coord, then z-coord.
'
'***********************************************************************
Public Function IsMinimumPoint(oPointToCheck As IJDPosition, _
                               oMinPoint As IJDPosition) As Boolean
Const sMETHOD As String = "IsMinimumPoint"
On Error GoTo ErrorHandler
    
    On Error Resume Next
    IsMinimumPoint = False
            
    Dim dValue1 As Double
    Dim dValue2 As Double
    Dim dValue3 As Double
    
    dValue1 = Abs(oPointToCheck.x - oMinPoint.x)
    dValue2 = Abs(oPointToCheck.y - oMinPoint.y)
    dValue3 = Abs(oPointToCheck.z - oMinPoint.z)
                
    ' Check if PointToCheck
    '   has minimum X value / has minimum abs(Y) value / has minimum Z value
    If dValue1 >= C_TOL001 And oPointToCheck.x < oMinPoint.x Then
        IsMinimumPoint = True
    ElseIf dValue2 >= C_TOL001 And Abs(oPointToCheck.y) < Abs(oMinPoint.y) Then
        IsMinimumPoint = True
    ElseIf dValue3 >= C_TOL001 And oPointToCheck.z < oMinPoint.z Then
        IsMinimumPoint = True
    End If
            

Exit Function

ErrorHandler:
  Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number

End Function
'**********************************************************
'* Determine the Lateral face port from a given edge port *
'**********************************************************
Public Sub Get_FacePortFromEdgePort(oEdgePort As IJStructPort, _
                                    oFacePort As IJStructPort)

Const sMETHOD As String = "Get_FacePortFromEdgePort"
On Error GoTo ErrorHandler
Dim sText As String

    Dim iIndex As Long
    Dim nPorts As Long
    
    Dim lXId As Long
    Dim lCtxId As Long
    Dim lOptId As Long
    Dim lOprId As Long
    
    Dim oPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oConnectable As IJConnectable
    Dim oStructConnectable As IJStructConnectable
    Dim oFacePortList As IJElements
    
    Set oPort = oEdgePort
    Set oConnectable = oPort.Connectable
    Set oStructConnectable = oConnectable
    
    Dim oTopologyLocate As GSCADStructGeomUtilities.IJTopologyLocate
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate

    lCtxId = oEdgePort.ContextID
    lOptId = oEdgePort.OperationID
    lOprId = oEdgePort.OperatorID
    lXId = oEdgePort.SectionID
    
    oStructConnectable.enumConnectablePortsByOperationAndTopology _
                                oFacePortList, _
                                vbNullString, _
                                JS_TOPOLOGY_FILTER_ALL_LFACES, _
                                True

    nPorts = oFacePortList.Count

    For iIndex = 1 To nPorts
        If TypeOf oFacePortList.Item(iIndex) Is IJStructPort Then
            Set oStructPort = oFacePortList.Item(iIndex)
            If oStructPort.ContextID = CTX_LATERAL_LFACE And _
               oStructPort.OperationID = lOptId And _
               oStructPort.OperatorID = lOprId Then
                Set oFacePort = oStructPort
                Exit For
               End If
        End If
    Next iIndex
  
Exit Sub

ErrorHandler:
  Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
  
End Sub
Public Function IsFlangedBracket(oPlatePart As IJPlatePart) As Boolean
    Dim bIsFlanged As Boolean
    
    bIsFlanged = False
    IsFlangedBracket = False
    
    ' Get the root system of this part
    Dim oParentSystem As IJSystem
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    On Error Resume Next
    ' True below means "recursive" - Navigate past design splits
    ' to retrieve the root system

    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oParentSystem, True
    
'    Set oStructDetailHelper = Nothing
    If oParentSystem Is Nothing Then
        ' Part is a stand-alone plate part
        ' These cannot be flanged (yet)
        Exit Function
    End If
    Dim oPlateSystem As IJPlateSystem
    Set oPlateSystem = oParentSystem
    Dim oPlateFlange_AE As IJPlateFlange_AE
    
    ' Get the flange AE (if any)
    
    'reserved parameter below is to support eventual multiple
    ' flanges.  If a plate system has multiple free-edges, each
    ' could theoretically be flanged.  This is not yet supported
    Set oPlateFlange_AE = oPlateSystem.FlangeActiveEntity(Nothing)
    If oPlateFlange_AE Is Nothing Then
        ' system is not flanged - part cannot be
        GoTo CleanUp
        Exit Function
    End If
    Dim oFlangedLeafPlate As IJPlate
    
    ' ask the flange AE for the affected leaf system
    
    ' true means get the affected leaf system
    ' currently, flanges cannot cross design splits
    Set oFlangedLeafPlate = oPlateFlange_AE.GetAffectedPlate(True)
    'get the immediate parent (leaf system) of the input part
    Dim oLeafSystem As IJSystem
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oLeafSystem, False      ' Used to be True
    Dim oLeafPlate As IJPlate
    Set oLeafPlate = oLeafSystem

    ' see if this part's leaf system is the one that is flanged
    If oLeafPlate Is oFlangedLeafPlate Then
        bIsFlanged = True
    End If

    IsFlangedBracket = bIsFlanged
    
CleanUp:
    Set oLeafPlate = Nothing
    Set oStructDetailHelper = Nothing
    Set oFlangedLeafPlate = Nothing
    Set oPlateFlange_AE = Nothing
    Set oPlateSystem = Nothing
    Set oParentSystem = Nothing
    
End Function

Public Function GetCatalogResourceMgr() As IJDPOM
    Const METHOD As String = "GetCatalogResourceMgr"

    Dim oCmnAppUtil As IJDCmnAppGenericUtil
    
    Set oCmnAppUtil = New CmnAppGenericUtil
    
    oCmnAppUtil.GetCatalogConnection GetCatalogResourceMgr
    
CleanUp:
    Set oCmnAppUtil = Nothing
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, sSOURCEFILE, METHOD).Number
End Function


Public Function IsSmartItemExist(oCatalogPOM As IJDPOM, strItemName As String) As Boolean

On Error GoTo CleanUp
    
    Dim oCatalogObj As Object
    Set oCatalogObj = oCatalogPOM.GetObject(strItemName)
    If Not oCatalogObj Is Nothing Then
        IsSmartItemExist = True
    Else
        IsSmartItemExist = False
    End If
    Exit Function

CleanUp:
    Set oCatalogObj = Nothing
    
End Function

Public Function EvaluateProfileCurvature(oProfileObj As Object) As ProfileCurvature

On Error GoTo CleanUp

    ' Before doing anything, get curvature type and set as candidate return value.
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfileObj
    EvaluateProfileCurvature = oProfileWrapper.CurvatureType

    ' Now you can do other stuff that could possibly overwrite this value.

    Dim oSDProfilePart As New StructDetailObjects.ProfilePart
    Set oSDProfilePart.object = oProfileObj
    
    If oSDProfilePart.ProfileType = sptEdgeReinforcement Then
        Dim strSectioName As String
        strSectioName = oSDProfilePart.SectionName
        
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
        Set oPartSupport.Part = oProfileObj

        Dim oProfilePartSupport As IJProfilePartSupport
        Set oProfilePartSupport = oPartSupport

        Dim bThicknessCentered As Boolean
        Dim oVector As IJDVector
        Dim oLandCurveWB As IJWireBody
        oProfilePartSupport.GetProfilePartLandingCurve oLandCurveWB, oVector, bThicknessCentered
        
        Dim oStartPos As IJDPosition, oEndPos As IJDPosition
        oLandCurveWB.GetEndPoints oStartPos, oEndPos

        Dim dProfileStraightLength As Double
        dProfileStraightLength = oStartPos.DistPt(oEndPos)
        
        Dim strQuery As String
        strQuery = "SELECT BowStringDepth From JUAMfgProfileBowStringLimit WHERE " & _
                   "SectionName = '" & strSectioName & "' AND " & _
                   "ProfileStraightLengthMin < " & CStr(Round(dProfileStraightLength, 3) + 0.001) & " AND " & _
                   "ProfileStraightLengthMax >=" & CStr(Round(dProfileStraightLength, 3) - 0.001)
        'MsgBox strQuery

        Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
        Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper

        Dim oQueryOutputValues() As Variant
        oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
        
        If (UBound(oQueryOutputValues) >= LBound(oQueryOutputValues)) Then
            Dim dMaxBowStringDepth As Double
            dMaxBowStringDepth = CDbl(oQueryOutputValues(LBound(oQueryOutputValues)))
            'MsgBox "dMaxBowStringDepth = " & dMaxBowStringDepth

            Dim dBowStringDepth As Double, dRadiusOfBow As Double
            dBowStringDepth = oProfileWrapper.GetDepthOfBowString(0, -1#, 0, Nothing, Nothing, CTX_INVALID, Nothing, dProfileStraightLength, dRadiusOfBow)
            'MsgBox "dBowStringDepth = " & dBowStringDepth
            'MsgBox "dRadiusOfBow = " & dRadiusOfBow

            If dBowStringDepth < dMaxBowStringDepth Then
                EvaluateProfileCurvature = PROFILE_CURVATURE_Straight
            End If
        End If
    End If
    
CleanUp:
    Set oPartSupport = Nothing
    Set oProfilePartSupport = Nothing
    Set oLandCurveWB = Nothing
    Set oProfileWrapper = Nothing
    Set oVector = Nothing
    Set oMfgCatalogQueryHelper = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    
    Exit Function
End Function

Public Function CheckIfProfileIsPlaneBend(oProfileObj As Object) As Boolean
Const sMETHOD As String = "CheckIfProfileIsPlaneBend"
On Error GoTo CleanUp

    Dim strSectioName As String
    Dim dProfileStraightLength As Double, dBowStringDepth As Double, dMaxBowStringDepth As Double
    Dim bThicknessCentered As Boolean
    Dim oVector As IJDVector
    Dim oPartSupport As IJPartSupport
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oLandCurveWB As IJWireBody
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oQueryOutputValues() As Variant
    Dim strQuery As String
    Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Dim oSDProfilePart As New StructDetailObjects.ProfilePart
    
    Set oProfileWrapper.object = oProfileObj
    Set oSDProfilePart.object = oProfileObj
    
    strSectioName = oSDProfilePart.SectionName
    
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oPartSupport.Part = oProfileObj
    Set oProfilePartSupport = oPartSupport
    oProfilePartSupport.GetProfilePartLandingCurve oLandCurveWB, oVector, bThicknessCentered
    
    oLandCurveWB.GetEndPoints oStartPos, oEndPos
    dProfileStraightLength = oStartPos.DistPt(oEndPos)
    
    Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
    strQuery = "SELECT BowStringDepth From JUAMfgProfileBowStringLimit WHERE (SectionName = '" + strSectioName + "') AND (ProfileStraightLengthMin < " + CStr(dProfileStraightLength) + ") AND (ProfileStraightLengthMax >" + CStr(dProfileStraightLength) + ")"
    'MsgBox strQuery

    oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
    
    If (UBound(oQueryOutputValues) >= LBound(oQueryOutputValues)) Then
        Dim dRadiusOfBow As Double
        dMaxBowStringDepth = oQueryOutputValues(0)
        'MsgBox "dMaxBowStringDepth = " & dMaxBowStringDepth
        dBowStringDepth = oProfileWrapper.GetDepthOfBowString(0, -1#, 0, Nothing, Nothing, CTX_INVALID, Nothing, dProfileStraightLength, dRadiusOfBow)
        'MsgBox "dBowStringDepth = " & dBowStringDepth
        'MsgBox "dRadiusOfBow = " & dRadiusOfBow
        If dBowStringDepth < dMaxBowStringDepth Then
            CheckIfProfileIsPlaneBend = True
        End If
    End If
    
CleanUp:
    Set oPartSupport = Nothing
    Set oProfilePartSupport = Nothing
    Set oLandCurveWB = Nothing
    Set oProfileWrapper = Nothing
    Set oVector = Nothing
    Set oMfgCatalogQueryHelper = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
End Function

