Attribute VB_Name = "TemplateMarkingHelper"
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Common helper module for Template marking functions
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "TemplateMarkingHelper.bas"

Public Enum enumTemplateShipDir
    eTopline = 0        ' Based on Top Line
    eBottomLine = 1     ' Based on bottom line
    eGlobal = 2         ' Based on Global directions
End Enum

' ***********************************************************************************
' Public Function CreateMarkAtPosition()
'
' Description:  It create A mark of fixed length at the given position along given vector
'
' ***********************************************************************************
Public Function CreateMarkAtPosition(oMarkPosition As IJDPosition, oMarkVector As IJDVector, MarkLength As Double, Optional bCenterMark As Boolean = False) As IJComplexString
    Const METHOD = "CreateMarkAtPosition"
    On Error GoTo ErrorHandler
    
    Dim oMarkLine As IJLine
    Set oMarkLine = New Line3d
    
    oMarkVector.length = MarkLength / 2
    
    If bCenterMark = True Then
        oMarkLine.DefineBy2Points oMarkPosition.x - oMarkVector.x, oMarkPosition.y - oMarkVector.y, oMarkPosition.z - oMarkVector.z, oMarkPosition.x + oMarkVector.x, oMarkPosition.y + oMarkVector.y, oMarkPosition.z + oMarkVector.z
    Else
        oMarkLine.DefineBy2Points oMarkPosition.x, oMarkPosition.y, oMarkPosition.z, oMarkPosition.x + oMarkVector.x, oMarkPosition.y + oMarkVector.y, oMarkPosition.z + oMarkVector.z
    End If
    
    Dim oCS As IJComplexString
    Set oCS = New ComplexString3d
    oCS.AddCurve oMarkLine, False
    
    Dim oCurve As IJCurve
    Set oCurve = oCS
    
    Dim dCurveLength     As Double
    dCurveLength = oCurve.length
    
    ' If the length of projects CS is more then 60mm then trim it.
    If dCurveLength > (MarkLength + 0.001) Then
    
        Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        
        Dim oStartPos As IJDPosition
        Set oStartPos = New DPosition
        
        Dim oEndPos As IJDPosition
        Set oEndPos = New DPosition
        
        oStartPos.Set dStartX, dStartY, dStartZ
        oEndPos.Set dEndX, dEndY, dEndZ
        
        Dim oMfgRuleHelper As MfgRuleHelpers.Helper
        Set oMfgRuleHelper = New Helper
        
        ' we need Fitting marks with length 15 mm
        If oStartPos.DistPt(oMarkPosition) > oEndPos.DistPt(oMarkPosition) Then
            oMfgRuleHelper.TrimCurveEnds oCS, dCurveLength - MarkLength, oStartPos
        Else
            oMfgRuleHelper.TrimCurveEnds oCS, dCurveLength - MarkLength, oEndPos
        End If
    End If
    
    Set CreateMarkAtPosition = oCS
    
CleanUp:
    Set oMarkLine = Nothing
    Set oCS = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function CreateGeom3D()
'
' Description:  It create Geom3D of input type for the curve supplied
'
' ***********************************************************************************
Public Function CreateGeom3D(oGeomCS As IJComplexString, eGeomType As StrMfgGeometryType, Optional oGeom3DMoniker As IMoniker, Optional strMarkName As String) As IJMfgGeom3d
    Const METHOD = "CreateGeom3D"
    On Error GoTo ErrorHandler
    
    Dim oGeom3dFactory      As IJMfgGeom3dFactory
    Set oGeom3dFactory = New MfgGeom3dFactory
    
    Dim oResourceManager    As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    If Not strMarkName = "" Then
        Dim oSystemMark     As IJMfgSystemMark
        Dim oMarkingInfo    As MarkingInfo
        
        Dim oSystemMarkFactory As IJMfgSystemMarkFactory
        Set oSystemMarkFactory = New MfgSystemMarkFactory
    
        ' Create a SystemMark object to store additional information
        Set oSystemMark = oSystemMarkFactory.Create(oResourceManager)
    
        ' QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark
        oMarkingInfo.Name = strMarkName
    End If
    
    Dim oGeom3D             As IJMfgGeom3d
    Set oGeom3D = oGeom3dFactory.Create(oResourceManager)
    oGeom3D.PutGeometry oGeomCS
    oGeom3D.PutGeometrytype eGeomType
    
    If Not oGeom3DMoniker Is Nothing Then
        oGeom3D.PutMoniker oGeom3DMoniker
    End If

    oSystemMark.Set3dGeometry oGeom3D
    
    Set CreateGeom3D = oGeom3D

    Set oSystemMark = Nothing
    Set oMarkingInfo = Nothing
    Set oGeom3D = Nothing
    Set oGeomCS = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Public Function GetActiveConnection
'
' Description: Helper function to get the active connection
'
' ***********************************************************************************
Public Function GetActiveConnection() As IJDAccessMiddle
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    oCmnAppGenericUtil.GetActiveConnection GetActiveConnection

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'****************************************************************************************************
'Description
'   Returns the GetActiveConnectionName.
'****************************************************************************************************
Public Function GetActiveConnectionName() As String

    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    GetActiveConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing
    
End Function

' ***********************************************************************************
' Public Function GetMarkVector()
'
' Description:  It create vector that will be used construct marks
'
' ***********************************************************************************
Public Function GetMarkVector(oMfgTemplate As IJMfgTemplate, oCurve As IJCurve, oPosition As IJDPosition) As IJDVector
    Const METHOD = "GetMarkVector"
    On Error GoTo ErrorHandler
    
    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    Dim oTemplatePlane  As IJPlane
    Set oTemplatePlane = oMfgTemplateReport.GetPlane
    
    Dim dNormalX As Double, dNormalY As Double, dNormalZ  As Double
    oTemplatePlane.GetNormal dNormalX, dNormalY, dNormalZ
    
    Dim oPlaneNormal    As IJDVector
    Set oPlaneNormal = New DVector
    
    oPlaneNormal.Set dNormalX, dNormalY, dNormalZ
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oTanVec As IJDVector
    Set oTanVec = oMfgGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = oTanVec.Cross(oPlaneNormal)
    
    Set GetMarkVector = oMarkVec
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : The purpose of this routine is to calculate the intersection between any two given
'            objects. The call will be delegated to the G&T PlaceIntersectionObject routine
'--------------------------------------------------------------------------------------------------
Public Function GetIntersection(pIntersectedObject As Object, pIntersectingObject As Object) As Object
On Error GoTo ErrorHandler
Const METHOD = "GetIntersection"

    ' Find the intersection.
    Dim oGeometryIntersector    As IMSModelGeomOps.DGeomOpsIntersect
    Set oGeometryIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    On Error Resume Next 'Needed for continuing with next skid mark if intersection fails
    Dim oIntersectionUnknown    As IUnknown        ' Resultant intersection.
    oGeometryIntersector.PlaceIntersectionObject Nothing, pIntersectedObject, pIntersectingObject, Nothing, oIntersectionUnknown
    
    On Error GoTo ErrorHandler
    Set GetIntersection = oIntersectionUnknown
    Set oGeometryIntersector = Nothing
    Set oIntersectionUnknown = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Create and plane given the orrt point and the normal.
'            Use the geometry factory to create a transient plane 3d object and return the plane
'--------------------------------------------------------------------------------------------------
Public Function CreatePlane(oRootPoint As IJDPosition, oNormalVec As IJDVector) As IJPlane
On Error GoTo ErrorHandler
Const METHOD = "CreatePlane"
        
    Dim oGeometryFactory    As IngrGeom3D.GeometryFactory
    Dim oPlane3D            As IngrGeom3D.IPlanes3d
    
    ' create persistent point
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oPlane3D = oGeometryFactory.Planes3d
    Set CreatePlane = oPlane3D.CreateByPointNormal(Nothing, oRootPoint.x, oRootPoint.y, oRootPoint.z, oNormalVec.x, oNormalVec.y, oNormalVec.z)
    
    Set oGeometryFactory = Nothing
    Set oPlane3D = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'***********************************************************************
' METHOD:  CreateShipDirectionMarks
'
' DESCRIPTION: Returns the Geom3D collection that represent template ship direction marks
'***********************************************************************
Public Function CreateShipDirectionMarks(oMfgTemplate As IJMfgTemplate, eTemplateDir As StrMfgTemplateDirection) As IJElements
Const METHOD = "CreateShipDirectionMarks"
On Error GoTo ErrorHandler

    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection

    Dim oMfgTemplateReport  As IJMfgTemplateReport
    Set oMfgTemplateReport = oMfgTemplate
    
    Dim oBCTLPosition    As IJDPosition
    Set oBCTLPosition = oMfgTemplateReport.GetPoint(BaseControlTopLinePoint)
    
    Dim oBCLPosition    As IJDPosition
    Set oBCLPosition = oMfgTemplateReport.GetPoint(BaseControlPoint)
    
    Dim dBCTP_BCP_Len   As Double
    dBCTP_BCP_Len = oBCTLPosition.DistPt(oBCLPosition)
    
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper
    
    Dim oTemplateTopLine   As IJLine
    Set oTemplateTopLine = oMfgTemplate.TopLine
    
    If oTemplateTopLine Is Nothing Then
        Exit Function
    End If
    
    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = oMfgGeomChild.GetParent
    
    Dim eTemplateSetType As MfgTemplateSetType
    eTemplateSetType = oMfgTemplateSet.GetTemplateSetType
    
    Dim oTemplateTopLineCS   As IJComplexString
    Set oTemplateTopLineCS = New ComplexString3d
    oTemplateTopLineCS.AddCurve oTemplateTopLine, False
      
    Dim oStartPos   As IJDPosition, oEndPos As IJDPosition, oTempPos As IJDPosition
    Dim oPrimaryMarkVec     As IJDVector
    Dim oSecondaryMarkVec   As IJDVector
    Dim oSideLine_Vec       As IJDVector
    Dim oResultVec          As IJDVector
    Dim oShipDirPosition    As IJDPosition
    Dim oTopLineDirection   As IJDVector
    Dim oUDirection         As IJDVector
    Dim oLocalCoordSystem   As IJLocalCoordinateSystem
    Dim dDotProduct         As Double
    
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oLocalCoordSystem = oMfgTemplate
        
    If eTemplateDir = eTopline Then
    
        Dim dTopLineLen     As Double
        Dim oTopLineWB  As IJWireBody
        Set oTopLineWB = oMfgRuleHelper.ComplexStringToWireBody(oTemplateTopLineCS)
    
        oTopLineWB.GetEndPoints oStartPos, oEndPos
        
        Set oTopLineDirection = oStartPos.Subtract(oEndPos)
        Set oUDirection = oLocalCoordSystem.XAxis
        
        dDotProduct = oTopLineDirection.Dot(oUDirection)
        
        If dDotProduct < 0 Then
            Set oTempPos = oStartPos
            Set oStartPos = oEndPos
            Set oEndPos = oTempPos
        End If
                
        dTopLineLen = oStartPos.DistPt(oEndPos)
    
        Set oPrimaryMarkVec = oStartPos.Subtract(oEndPos)
        oPrimaryMarkVec.length = dTopLineLen / 4
        
        Set oSideLine_Vec = oBCLPosition.Subtract(oBCTLPosition)
        oSideLine_Vec.length = dBCTP_BCP_Len / 4

        Set oResultVec = oPrimaryMarkVec.Add(oSideLine_Vec)
        
        Dim oTopLineMidPos  As IJDPosition
        Set oTopLineMidPos = oMfgRuleHelper.GetMiddlePoint(oTopLineWB)
        Set oShipDirPosition = oTopLineMidPos.Offset(oResultVec)
        
    ElseIf eTemplateDir = eBottomLine Then
    
        Dim dBottomLineLen          As Double
        Dim oTemplateBottomLine     As IJCurve
        Set oTemplateBottomLine = oMfgTemplate.GetTemplateLocationMarkLine
        
        dBottomLineLen = oTemplateBottomLine.length
        
        Set oPrimaryMarkVec = oMfgGeomHelper.GetTangentByPointOnCurve(oTemplateBottomLine, oBCLPosition)
        oPrimaryMarkVec.length = dBottomLineLen / 4
        
        Set oSideLine_Vec = oBCTLPosition.Subtract(oBCLPosition)
        oSideLine_Vec.length = dBCTP_BCP_Len / 4
        
        Set oResultVec = oPrimaryMarkVec.Add(oSideLine_Vec)
        Set oShipDirPosition = oBCLPosition.Offset(oResultVec)
        
    Else    ' Global direction
        ' No IMPL
    End If
    
    oResultVec.length = 1
    oPrimaryMarkVec.length = 1
    
    Dim oNormalVec As IJDVector
    Set oNormalVec = New DVector
    oNormalVec.Set 0, 1, 0
    
    Dim oRootPoint As IJDPosition
    Set oRootPoint = New DPosition
    
    Dim oCenPlane   As IJPlane
    Set oCenPlane = CreatePlane(oRootPoint, oNormalVec)
    
    On Error Resume Next    ' Intersection routine can fail if there is no intersection
    
    Dim oIntersectionObj    As Object
    Set oIntersectionObj = oMfgGeomHelper.IntersectCurveWithPlane(oTemplateTopLineCS, oCenPlane)
    
    On Error GoTo ErrorHandler
    
    Dim bCenterLine As Boolean
    If Not oIntersectionObj Is Nothing Then
        bCenterLine = True
    End If
       
    Dim oMarkCS     As IJComplexString
    
    If eTemplateSetType = PlateTemplate Then
        Set oMarkCS = CreateMarkAtPosition(oShipDirPosition, oPrimaryMarkVec, TEMPLATE_SHIP_DIR_PRIMARY_LENGTH, False)
    Else ' ProfileTemplate
        Set oMarkCS = CreateMarkAtPosition(oShipDirPosition, oPrimaryMarkVec, PR_TEMPLATE_SHIP_DIR_PRIMARY_LENGTH, False)
    End If
    
    oPrimaryMarkVec.length = 1
    
    Dim strMarkingName As String
    strMarkingName = GetMarkingName(bCenterLine, oShipDirPosition, oPrimaryMarkVec)
    
    Dim oMfgGeom3D As IJMfgGeom3d
    Set oMfgGeom3D = CreateGeom3D(oMarkCS, STRMFG_DIRECTION, , strMarkingName)
    
    oGeomElems.Add oMfgGeom3D
    Set oMfgGeom3D = Nothing
    Set oMarkCS = Nothing
    
    Dim oTemplatePlane  As IJPlane
    Set oTemplatePlane = oMfgTemplateReport.GetPlane
    
    Dim dNormalX As Double, dNormalY As Double, dNormalZ  As Double
    oTemplatePlane.GetNormal dNormalX, dNormalY, dNormalZ
    
    Dim oPlaneNormal    As IJDVector
    Set oPlaneNormal = New DVector
    
    oPlaneNormal.Set dNormalX, dNormalY, dNormalZ
    
    Set oSecondaryMarkVec = oPrimaryMarkVec.Cross(oPlaneNormal)
    
    ' below code is to get ship direction marks like below
        ' Y
        '
        '
        '
        ' ' ' ' ' ' X
    If oSecondaryMarkVec.Dot(oSideLine_Vec) < 0 Then
        oSecondaryMarkVec.length = 1
    Else
        oSecondaryMarkVec.length = -1
    End If
    
    If eTemplateSetType = PlateTemplate Then
        Set oMarkCS = CreateMarkAtPosition(oShipDirPosition, oSecondaryMarkVec, TEMPLATE_SHIP_DIR_SECONDARY_LENGTH, False)
    Else ' ProfileTemplate
        Set oMarkCS = CreateMarkAtPosition(oShipDirPosition, oSecondaryMarkVec, PR_TEMPLATE_SHIP_DIR_SECONDARY_LENGTH, False)
    End If
    
    Dim oMGHelper   As IJMfgMGHelper
    Set oMGHelper = New MfgMGHelper
    
    Dim oRevMarkCS  As IJComplexString
    oMGHelper.ReverseComplexString oMarkCS, oRevMarkCS
    
    oSecondaryMarkVec.length = 1
    
    strMarkingName = GetMarkingName(bCenterLine, oShipDirPosition, oSecondaryMarkVec)
    Set oMfgGeom3D = CreateGeom3D(oRevMarkCS, STRMFG_DIRECTION, , strMarkingName)
    
    oGeomElems.Add oMfgGeom3D
    Set oMfgGeom3D = Nothing
    
    Set CreateShipDirectionMarks = oGeomElems
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'***********************************************************************
' METHOD:  CreateTemplatePlaneFromContour
'
' DESCRIPTION: Returns the plane with template contours as boundaries
'***********************************************************************
Public Function CreateTemplatePlaneFromContour(oMfgTemplate As IJMfgTemplate) As IJPlane
Const METHOD = "CreateTemplatePlaneFromContour"
On Error GoTo ErrorHandler
   
    Dim oContourElements As IJElements
    Set oContourElements = oMfgTemplate.GetSideBoundaries
    
    If oContourElements Is Nothing Then
        Set oContourElements = New JObjectCollection
    End If
    
    Dim oTemplateBottomLine As Object
    Set oTemplateBottomLine = oMfgTemplate.GetTemplateLocationMarkLine
    
    oContourElements.Add oTemplateBottomLine
    
    Dim oTemplateTopLine As Object
    Set oTemplateTopLine = oMfgTemplate.TopLine
    
    oContourElements.Add oTemplateBottomLine
    
    Dim oComplexString  As IJComplexString
    Set oComplexString = New ComplexString3d
    
    Dim iIndex  As Integer
    For iIndex = 1 To oContourElements.Count
        Dim oCurve As IJComplexString
        Set oCurve = oContourElements.Item(iIndex)
        
        Dim oTempElements   As IJElements
        oCurve.GetCurves oTempElements
        
        oComplexString.SetCurves oTempElements
    Next
    
    Dim oPlane    As IJPlane
    Set oPlane = New Plane3d
    oPlane.SetBoundary 1, oComplexString
    
    Set CreateTemplatePlaneFromContour = oPlane
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'***********************************************************************
' METHOD:  GetTemplateOrientation
'
' DESCRIPTION: Returns the enum that represent template orientation
'***********************************************************************
Public Function GetTemplateOrientation(oTemplatePlane As IJPlane) As StrMfgTemplateDirection
Const METHOD = "GetTemplateOrientation"
On Error GoTo ErrorHandler
   
    Dim dX As Double, dY As Double, dZ As Double
    oTemplatePlane.GetNormal dX, dY, dZ
    Dim eTemplateOrient As StrMfgTemplateDirection
    
    If dX >= 0.5 And dX > Abs(dY) And dX > Abs(dZ) Then
        eTemplateOrient = MfgTDTransversal
    ElseIf dY >= 0.5 And dY > Abs(dZ) Then
        eTemplateOrient = MfgTDLongitudinal
    Else
        eTemplateOrient = MfgTDWaterline
    End If
    
    GetTemplateOrientation = eTemplateOrient
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'***********************************************************************
' METHOD:  GetMarkingName
'
' DESCRIPTION: Returns the string that represent the global ship direction
'***********************************************************************
Public Function GetMarkingName(bCenterCross As Boolean, oPosition As IJDPosition, oVector As IJDVector) As String
    Const METHOD As String = "GetMarkingName"
    On Error GoTo ErrorHandler
        
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim dPosX As Double, dPosY As Double, dPosZ As Double
    
    oVector.Get dNormalX, dNormalY, dNormalZ
    oPosition.Get dPosX, dPosY, dPosZ
    
    Dim strDirection    As String
    
    If dNormalX >= 0.5 And dNormalX > Abs(dNormalY) And dNormalX > Abs(dNormalZ) Then
        strDirection = "F"
    ElseIf dNormalY >= 0.5 And dNormalY > Abs(dNormalX) And dNormalY > Abs(dNormalZ) Then
        
        If dPosY > 0.001 Then
            strDirection = "O" ' Outer
        Else
            strDirection = "I" ' Inner
        End If
        
        If bCenterCross = True Then
            strDirection = "P"
        End If
        
    ElseIf dNormalZ >= 0.5 And dNormalZ > Abs(dNormalX) And dNormalZ > Abs(dNormalY) Then
        strDirection = "U"
    End If
    
    If dNormalX < -0.5 And dNormalX < dNormalY And dNormalX < dNormalZ Then
        strDirection = "A"
    ElseIf dNormalY < -0.5 And dNormalY < dNormalZ Then
        
        If dPosY > 0.001 Then
            strDirection = "I" ' Inner
        Else
            strDirection = "O" ' Outer
        End If
        
        If bCenterCross = True Then
            strDirection = "S"
        End If
        
    ElseIf dNormalZ < -0.5 And dNormalZ < dNormalX And dNormalZ < dNormalY Then
        strDirection = "D"
    End If
    
    GetMarkingName = strDirection
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

