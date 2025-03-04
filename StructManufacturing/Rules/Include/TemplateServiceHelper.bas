Attribute VB_Name = "TemplateServiceHelper"
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: TubeTemplateHelper
'
'  Abstract:    Common helper module for Tube Template functions
'
'  History:
'      Siva        19th Oct 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "TubeTemplateHelper.bas"

' ***********************************************************************************
' Public Function GetTubeSurfaceInfo()
'
' Description:  Gets the tube surface from side input
'
' ***********************************************************************************
Private Function GetTubeSurfaceInfo(ByVal oMemberPart As Object, strTemplateSide As String) As Object
    Const METHOD = "GetTubeSurfaceInfo"
    On Error GoTo ErrorHandler
    
    Dim oTemplateHelper     As IJMfgTemplateHelper
    Set oTemplateHelper = New MfgTemplateHelper
    
    Dim lSectionID  As Long
    
    If strTemplateSide = "Outer" Then
        lSectionID = 3073  ' Outer Tube
    Else
        lSectionID = 3074  ' Inner Tube
    End If
    
    Dim oSurfaceElements    As IJElements
    Set oSurfaceElements = oTemplateHelper.GetContoursOfGivenType(oMemberPart, JS_TOPOLOGY_FILTER_ALL_LFACES, CTX_LATERAL_LFACE, lSectionID)
    
    Set GetTubeSurfaceInfo = oSurfaceElements.Item(1)
       
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateMarkAtPosition()
'
' Description:  It create A mark of fixed length at the given position along given vector
'
' ***********************************************************************************
Public Function CreateLineAtPosition(oMarkPosition As IJDPosition, oMarkVector As IJDVector, dMarkLength As Double, Optional bCenterMark As Boolean = False) As Object
    Const METHOD = "CreateLineAtPosition"
    On Error GoTo ErrorHandler
    
    Dim oMarkLine As IJLine
    Set oMarkLine = New Line3d
    
    If bCenterMark = True Then
        oMarkVector.length = 0.5 * dMarkLength
        oMarkLine.DefineBy2Points oMarkPosition.x - oMarkVector.x, oMarkPosition.y - oMarkVector.y, oMarkPosition.z - oMarkVector.z, oMarkPosition.x + oMarkVector.x, oMarkPosition.y + oMarkVector.y, oMarkPosition.z + oMarkVector.z
    Else
        oMarkVector.length = dMarkLength
        oMarkLine.DefineBy2Points oMarkPosition.x, oMarkPosition.y, oMarkPosition.z, oMarkPosition.x + oMarkVector.x, oMarkPosition.y + oMarkVector.y, oMarkPosition.z + oMarkVector.z
    End If
    
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oMarkLine, False
    
    Set CreateLineAtPosition = oCS
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateLineFromPoints()
'
' Description:  It create A mark of fixed length at the given position along given vector
'
' ***********************************************************************************
Public Function CreateLineFromPoints(oStartPos As IJDPosition, oEndPos As IJDPosition) As Object
    Const METHOD = "CreateLineFromPoints"
    On Error GoTo ErrorHandler
    
    Dim oMarkLine As IJLine
    Set oMarkLine = New Line3d
    
    oMarkLine.DefineBy2Points oStartPos.x, oStartPos.y, oStartPos.z, oEndPos.x, oEndPos.y, oEndPos.z
    
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oMarkLine, False
    
    Set CreateLineFromPoints = oCS
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function GetDirectionString()
'
' Description:  Returns the direction string based on input vector direction and input position
'
' ***********************************************************************************
Public Function GetDirectionString(oVector As IJDVector, oPos As IJDPosition) As String
    Const METHOD = "GetDirectionString"
    On Error GoTo ErrorHandler
    
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oVector.Get dNormalX, dNormalY, dNormalZ
    oPos.Get dRootX, dRootY, dRootZ
    
    Dim strDirection As String
    
    If dNormalX >= 0.5 And dNormalX > Abs(dNormalY) And dNormalX > Abs(dNormalZ) Then
        strDirection = "F"
    ElseIf dNormalY >= 0.5 And dNormalY > Abs(dNormalX) And dNormalY > Abs(dNormalZ) Then
        
        If dRootY > 0.001 Then
            strDirection = "O" ' Outer
        Else
            strDirection = "I" ' Inner
        End If
        
    ElseIf dNormalZ >= 0.5 And dNormalZ > Abs(dNormalX) And dNormalZ > Abs(dNormalY) Then
        strDirection = "U"
    End If
    
    If dNormalX < -0.5 And dNormalX < dNormalY And dNormalX < dNormalZ Then
        strDirection = "A"
    ElseIf dNormalY < -0.5 And dNormalY < dNormalZ Then
        
        If dRootY > 0.001 Then
            strDirection = "I" ' Inner
        Else
            strDirection = "O" ' Outer
        End If
        
    ElseIf dNormalZ < -0.5 And dNormalZ < dNormalX And dNormalZ < dNormalY Then
        strDirection = "D"
    End If
    
    GetDirectionString = strDirection
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function GetOppositeDirectionString()
'
' Description:  Returns the direction string opposite to input direction string
'
' ***********************************************************************************
Public Function GetOppositeDirectionString(strinputDir As String) As String
    Const METHOD = "GetOppositeDirectionString"
    On Error GoTo ErrorHandler
    
    Dim strOutDir As String
    If strinputDir = "F" Then
        strOutDir = "A"
    ElseIf strinputDir = "A" Then
        strOutDir = "F"
    ElseIf strinputDir = "O" Then
        strOutDir = "I"
    ElseIf strinputDir = "I" Then
        strOutDir = "O"
    ElseIf strinputDir = "P" Then
        strOutDir = "S"
    ElseIf strinputDir = "S" Then
        strOutDir = "P"
    ElseIf strinputDir = "U" Then
        strOutDir = "D"
    ElseIf strinputDir = "D" Then
        strOutDir = "U"
    End If
    
    GetOppositeDirectionString = strOutDir
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function GetReferenceCurve()
'
' Description:  creates reference curve i.e, intersection of plate represent profile orientation and tube surface
'
' ***********************************************************************************
Public Function GetReferenceCurve(oMemberPart As Object, oSurfaceBody As Object, oBasePlane As IJPlane, oTopCurve As IJCurve) As Object
Const METHOD = "GetReferenceCurve"
On Error GoTo ErrorHandler
    
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
        
    oBasePlane.GetNormal dNormalX, dNormalY, dNormalZ
    oBasePlane.GetRootPoint dRootX, dRootY, dRootZ
    
    ' Create Ref plane
    Dim oRefPlane    As IJPlane
    Set oRefPlane = New Plane3d
    
    oRefPlane.SetNormal dNormalX, dNormalY, dNormalZ
    oRefPlane.SetRootPoint dRootX, dRootY, dRootZ
    
    ' Create RefPlaneNormal object
  
    Dim oRefPlaneNormal As IJDVector
    Set oRefPlaneNormal = New DVector
    oRefPlaneNormal.Set dNormalX, dNormalY, dNormalZ
    oRefPlaneNormal.length = oTopCurve.length / 2
    
    ' Create root pos object
    Dim oBasePlaneRootPos    As IJDPosition
    Set oBasePlaneRootPos = New DPosition
    
    oBasePlaneRootPos.Set dRootX, dRootY, dRootZ
    
    Dim oNewRootPos     As IJDPosition
    Set oNewRootPos = oBasePlaneRootPos.Offset(oRefPlaneNormal)
    
    oRefPlane.SetRootPoint oNewRootPos.x, oNewRootPos.y, oNewRootPos.z
    
    Dim oMfgGeomUtilwrapper As IJDMfgGeomUtilWrapper
    Set oMfgGeomUtilwrapper = New MfgGeomUtilWrapper

    Dim oPlaneSheetBody   As Object
    Set oPlaneSheetBody = oMfgGeomUtilwrapper.CreateFinitePlane(oRefPlane)
            
    Dim oProfileSupport As IJProfilePartSupport
    Set oProfileSupport = New ProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oProfileSupport
    Set oPartSupport.Part = oMemberPart

    ' Find the LandingCurve and thickness direction of the web
    Dim oLandingCurve As IJWireBody
    Dim oThicknessDirection As IJDVector
    oProfileSupport.GetProfilePartLandingCurve oLandingCurve, oThicknessDirection, False, SideA

    Dim oModelBodyLC  As IJDModelBody
    Set oModelBodyLC = oLandingCurve
    
    Dim dMinDist    As Double
    Dim oClosestPos As IJDPosition, oClosestPos2   As IJDPosition
    
    oModelBodyLC.GetMinimumDistance oPlaneSheetBody, oClosestPos, oClosestPos2, dMinDist
    
    Dim oXVector As IJDVector, oYVector As IJDVector
    Dim oOrigin  As IJDPosition
    
    oProfileSupport.GetOrientation oClosestPos, oXVector, oYVector, oOrigin
    
    oXVector.length = 1
    oYVector.length = 1
    
    Dim oNormalVector   As IJDVector
    Set oNormalVector = oXVector.Cross(oYVector)
    
    oNormalVector.length = 1
    
    ' define ref plane with new calculated pos as root point and orientation plane normal as normal vector
    'Dim dX As Double, dY As Double, dZ As Double, dNormalX As Double, dNormalY As Double, dNormalZ As Double
    oRefPlane.DefineByPointNormal oClosestPos.x, oClosestPos.y, oClosestPos.z, oNormalVector.x, oNormalVector.y, oNormalVector.z
    
    Dim oRefPlaneSheetBody   As Object
    Set oRefPlaneSheetBody = oMfgGeomUtilwrapper.CreateFinitePlane(oRefPlane)
    
    Dim oIntersector As IJDTopologyIntersect
    Set oIntersector = New DGeomOpsIntersect
    
    Dim oRefCurveWB As Object
    oIntersector.PlaceIntersectionObject Nothing, oRefPlaneSheetBody, oSurfaceBody, Nothing, oRefCurveWB

    Set GetReferenceCurve = oRefCurveWB

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Public Function CreateGeom2D()
'
' Description:  It create Geom2D of input type for the curve supplied
'
' ***********************************************************************************
Public Function CreateGeom2D(oGeomCS As IJComplexString, eGeomType As StrMfgGeometryType, Optional oGeom2DMoniker As IMoniker, Optional strMarkName As String) As IJMfgGeom2d
    Const METHOD = "CreateGeom2D"
    On Error GoTo ErrorHandler

    Dim oSystemMark     As IJMfgSystemMark
    Dim oMarkingInfo    As MarkingInfo
    Dim oGeom2D         As IJMfgGeom2d

    Dim oSystemMarkFactory As IJMfgSystemMarkFactory
    Dim oGeom2DFactory As IJMfgGeom2dFactory

    Set oSystemMarkFactory = New MfgSystemMarkFactory
    Set oGeom2DFactory = New MfgGeom2dFactory
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    ' Create a SystemMark object to store additional information
    Set oSystemMark = oSystemMarkFactory.Create(oResourceManager)

    ' QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark

    If Not strMarkName = "" Then
        oMarkingInfo.Name = strMarkName
    End If

    Set oGeom2D = oGeom2DFactory.Create(oResourceManager)
    oGeom2D.PutGeometry oGeomCS
    oGeom2D.PutGeometrytype eGeomType

    If Not oGeom2DMoniker Is Nothing Then
        oGeom2D.PutMoniker oGeom2DMoniker
    End If

    oSystemMark.Set2dGeometry oGeom2D
    
    Set CreateGeom2D = oGeom2D

    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oGeom2D = Nothing
    Set oGeomCS = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
