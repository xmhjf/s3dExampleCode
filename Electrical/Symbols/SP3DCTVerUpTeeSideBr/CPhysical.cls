VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          dkl
'   Creation Date:  Wednesday, Jun 6 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'    This symbol is made to function as a Vertical Upward Three-Way, reference Zhen Hua Catalog.
'
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------    -----    ------------------
'  6.JUN.2007     dkl       CR-114880 Created the symbol.
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    
    Dim parActualWidth1 As Double
    Dim parActualDepth1 As Double
    Dim parTangentLength As Double
    Dim parActualWidth3 As Double
    Dim parActualDepth3 As Double
    Dim parHeaderLength As Double
    Dim parBranchLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHeaderLength = arrayOfInputs(2)
    parBranchLength = arrayOfInputs(3)
'    parTangentLength1 = arrayOfInputs(4)
'    parTangentLength2 = arrayOfInputs(5)
'    parTangentLength3 = arrayOfInputs(6)
    
    iOutput = 0
'   Retrieve Part properties: TangentLength
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    
    Dim dTangentLength1 As Double   'Tangent length along Port 1.
    Dim dTangentLength2 As Double   'Tangent length along Port 2.
    Dim dTangentLength3 As Double   'Tangent length along Port 3.
    
'   The following Part data Bases are implemented,
'   1) Vertical Segmented Upward Three Way, Side branch, Symmetrical, specified by Header Length, Branch Length and Tangent Length (Default Implementation).
'   2) Vertical Segmented Upward Three Way, Side branch, Asymmetrical, specified by Header Length, Branch Length, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
    Dim lCableTrayPartDataBasis As Integer
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    Select Case lCableTrayPartDataBasis
    
        Case Is <= 1, 79     'Vertical Segmented Upward Three Way, Side branch, Symmetrical, specified by
                             ' Header Length, Branch Length and Tangent Length (Default Implementation).
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            
        Case 81      'Vertical Segmented Upward Three Way, Side branch, Asymmetrical, specified by
                     'Header Length, Branch Length, Tangent Length 1, Tangent Length 2 and Tangent Length 3.
            dTangentLength1 = arrayOfInputs(4)
            dTangentLength2 = arrayOfInputs(5)
            dTangentLength3 = arrayOfInputs(6)

        Case Else
            GoTo ErrorLabel:
    End Select
    
'Check to validate that if the tangentLength is zero, set it to a very small value
    If CmpDblLessThanOrEqualTo(dTangentLength1, 0) Then dTangentLength1 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength2, 0) Then dTangentLength2 = 0.0001
    If CmpDblLessThanOrEqualTo(dTangentLength3, 0) Then dTangentLength3 = 0.0001
        
'   Port 1 position
    ' Retrieve Ports 1 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth1, parActualDepth1)

    Dim dHalfDepth1 As Double
    dHalfDepth1 = parActualDepth1 / 2
    
    Dim dHalfWidth1 As Double
    dHalfWidth1 = parActualWidth1 / 2
        
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    oPort1.Set -parHeaderLength / 2, 0, 0
    
    ' Defining points for creating the linestring at port 1.
    Dim dPort1S(0 To 11)  As Double
    ' Tray top edge at the side opposite to the side of the branch.
    dPort1S(0) = oPort1.x
    dPort1S(1) = oPort1.y + dHalfWidth1
    dPort1S(2) = oPort1.z + dHalfDepth1
    ' Tray bottom edge at the side opposite to the side of the branch.
    dPort1S(3) = oPort1.x
    dPort1S(4) = oPort1.y + dHalfWidth1
    dPort1S(5) = oPort1.z - dHalfDepth1
    ' Tray bottom edge at side of the side-branch.
    dPort1S(6) = oPort1.x
    dPort1S(7) = oPort1.y - dHalfWidth1
    dPort1S(8) = oPort1.z - dHalfDepth1
    ' Tray top edge at side of the side-branch.
    dPort1S(9) = oPort1.x
    dPort1S(10) = oPort1.y - dHalfWidth1
    dPort1S(11) = oPort1.z + dHalfDepth1
    
'   Port 2 position
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    oPort2.Set parHeaderLength / 2, 0, 0
    
'   Port 3 position
    '   Retrieve Ports 1 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(3, oPartFclt, parActualWidth3, parActualDepth3)
    
    Dim dHalfWidth3 As Double
    dHalfWidth3 = parActualWidth3 / 2
    
    Dim dHalfDepth3 As Double
    dHalfDepth3 = parActualDepth3 / 2
    
    Dim oPort3 As AutoMath.DPosition 'Port 3 center point
    Set oPort3 = New AutoMath.DPosition
    oPort3.Set 0, dHalfWidth1 - dHalfDepth3, parBranchLength - dHalfDepth1
    
    ' Defining points for creating the linestring at port 3.
    Dim dPort3S(0 To 11)  As Double

    dPort3S(0) = oPort3.x - dHalfWidth3
    dPort3S(1) = oPort3.y + dHalfDepth3
    dPort3S(2) = oPort3.z
    
    dPort3S(3) = oPort3.x - dHalfWidth3
    dPort3S(4) = oPort3.y - dHalfDepth3
    dPort3S(5) = oPort3.z
    
    dPort3S(6) = oPort3.x + dHalfWidth3
    dPort3S(7) = oPort3.y - dHalfDepth3
    dPort3S(8) = oPort3.z

    dPort3S(9) = oPort3.x + dHalfWidth3
    dPort3S(10) = oPort3.y + dHalfDepth3
    dPort3S(11) = oPort3.z

'   Construct Tray bottom
    Dim oProjectionVector As AutoMath.DVector
    Set oProjectionVector = New AutoMath.DVector
    oProjectionVector.Set 1, 0, 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oLineString As IngrGeom3D.LineString3d
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dPort1S)
    
    Dim objTrayBottom As Object
    Set objTrayBottom = PlaceProjection(m_OutputColl, oLineString, oProjectionVector, _
                                                                    parHeaderLength, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTrayBottom
    Set objTrayBottom = Nothing
    Set oLineString = Nothing

' Construct Left segment.
    ' Defining Points for creating the left side of the branch.
    Dim dPointsDefiningLeftBranch(0 To 11) As Double
    
    dPointsDefiningLeftBranch(0) = dPort1S(0)
    dPointsDefiningLeftBranch(1) = dPort1S(1)
    dPointsDefiningLeftBranch(2) = dPort1S(2)

    dPointsDefiningLeftBranch(3) = dPointsDefiningLeftBranch(0) + dTangentLength1
    dPointsDefiningLeftBranch(4) = dPointsDefiningLeftBranch(1)
    dPointsDefiningLeftBranch(5) = dPointsDefiningLeftBranch(2)

    dPointsDefiningLeftBranch(6) = dPort3S(0)
    dPointsDefiningLeftBranch(7) = dPort3S(1)
    dPointsDefiningLeftBranch(8) = dPort3S(2) - dTangentLength3
    
    dPointsDefiningLeftBranch(9) = dPointsDefiningLeftBranch(6)
    dPointsDefiningLeftBranch(10) = dPointsDefiningLeftBranch(7)
    dPointsDefiningLeftBranch(11) = dPort3S(2)
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dPointsDefiningLeftBranch)
         
    Dim oLineLeft As IngrGeom3D.Line3d
    Set oLineLeft = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningLeftBranch(0), dPointsDefiningLeftBranch(1), dPointsDefiningLeftBranch(2), _
        dPointsDefiningLeftBranch(0), dPointsDefiningLeftBranch(1) - dHalfDepth1 * 2, dPointsDefiningLeftBranch(2))
       
    Dim objSurfset As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
    Set objSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
              oLineString, oLineLeft, 0, 0, stnorm, ednorm, False)
              
' Set the output
    Dim ObjSurface As Object
    For Each ObjSurface In objSurfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
    Next ObjSurface
    Set ObjSurface = Nothing
    Set oLineLeft = Nothing

' Construct Right segment.
    ' Defining Points for creating the branch segment on the right.
    Dim dPointsDefiningRightBranch(0 To 11) As Double
    dPointsDefiningRightBranch(0) = oPort2.x
    dPointsDefiningRightBranch(1) = oPort2.y + dHalfWidth1
    dPointsDefiningRightBranch(2) = oPort2.z + dHalfDepth1

    dPointsDefiningRightBranch(3) = dPointsDefiningRightBranch(0) - dTangentLength2
    dPointsDefiningRightBranch(4) = dPointsDefiningRightBranch(1)
    dPointsDefiningRightBranch(5) = dPointsDefiningRightBranch(2)

    dPointsDefiningRightBranch(6) = dPort3S(9)
    dPointsDefiningRightBranch(7) = dPort3S(10)
    dPointsDefiningRightBranch(8) = dPort3S(11) - dTangentLength3
    
    dPointsDefiningRightBranch(9) = dPointsDefiningRightBranch(6)
    dPointsDefiningRightBranch(10) = dPointsDefiningRightBranch(7)
    dPointsDefiningRightBranch(11) = dPort3S(11)
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dPointsDefiningRightBranch)
       
    Dim oLineRight As IngrGeom3D.Line3d
    Set oLineRight = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningRightBranch(0), dPointsDefiningRightBranch(1), dPointsDefiningRightBranch(2), _
        dPointsDefiningRightBranch(0), dPointsDefiningRightBranch(1) - dHalfDepth1 * 2, dPointsDefiningRightBranch(2))
               
    Set objSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
              oLineString, oLineRight, 0, 0, stnorm, ednorm, False)
              
' Set the output
    Dim ObjSurface1 As Object
    For Each ObjSurface1 In objSurfset
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface1
    Next ObjSurface1
    Set objSurfset = Nothing
    Set ObjSurface1 = Nothing
    Set oLineRight = Nothing

' Construct the Branch base
    Dim oLeftInclinedLine As IngrGeom3D.Line3d
    Set oLeftInclinedLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningLeftBranch(3), dPointsDefiningLeftBranch(4), dPointsDefiningLeftBranch(5), _
        dPointsDefiningLeftBranch(6), dPointsDefiningLeftBranch(7), dPointsDefiningLeftBranch(8))
    
    Dim oBranchSegment As Collection
    Set oBranchSegment = New Collection
    oBranchSegment.Add oLeftInclinedLine
    
    Dim oLeftTangent As IngrGeom3D.Line3d
    Set oLeftTangent = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningLeftBranch(6), dPointsDefiningLeftBranch(7), dPointsDefiningLeftBranch(8), _
        dPointsDefiningLeftBranch(9), dPointsDefiningLeftBranch(10), dPointsDefiningLeftBranch(11))
    oBranchSegment.Add oLeftTangent
    
    Dim oHorLineatPort3 As IngrGeom3D.Line3d
    Set oHorLineatPort3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningLeftBranch(9), dPointsDefiningLeftBranch(10), dPointsDefiningLeftBranch(11), _
        dPointsDefiningRightBranch(9), dPointsDefiningRightBranch(10), dPointsDefiningRightBranch(11))
    oBranchSegment.Add oHorLineatPort3
    
    Dim oRightTangent As IngrGeom3D.Line3d
    Set oRightTangent = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningRightBranch(9), dPointsDefiningRightBranch(10), dPointsDefiningRightBranch(11), _
        dPointsDefiningRightBranch(6), dPointsDefiningRightBranch(7), dPointsDefiningRightBranch(8))
    oBranchSegment.Add oRightTangent
    
    Dim oRightInclinedLine As IngrGeom3D.Line3d
    Set oRightInclinedLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningRightBranch(6), dPointsDefiningRightBranch(7), dPointsDefiningRightBranch(8), _
        dPointsDefiningRightBranch(3), dPointsDefiningRightBranch(4), dPointsDefiningRightBranch(5))
    oBranchSegment.Add oRightInclinedLine
    
    Dim oHorLineatBranchBaseBottom As IngrGeom3D.Line3d
    Set oHorLineatBranchBaseBottom = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPointsDefiningRightBranch(3), dPointsDefiningRightBranch(4), dPointsDefiningRightBranch(5), _
        dPointsDefiningLeftBranch(3), dPointsDefiningLeftBranch(4), dPointsDefiningLeftBranch(5))
    oBranchSegment.Add oHorLineatBranchBaseBottom
    
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set dPointsDefiningLeftBranch(3), dPointsDefiningLeftBranch(4), dPointsDefiningLeftBranch(5)
    
    Dim objBranchSegmentString  As IngrGeom3D.ComplexString3d
    Set objBranchSegmentString = PlaceTrCString(oStPoint, oBranchSegment)
    
    oProjectionVector.Set 0, -1, 0
    
    Dim objBranchBase As IngrGeom3D.Plane3d
    Set objBranchBase = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
            dPointsDefiningLeftBranch(3), dPointsDefiningLeftBranch(4), dPointsDefiningLeftBranch(5), _
            oProjectionVector.x, oProjectionVector.y, oProjectionVector.z)
    Call objBranchBase.AddBoundary(objBranchSegmentString)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchBase
    Set objBranchBase = Nothing
    Set oProjectionVector = Nothing
    Set objBranchSegmentString = Nothing
    Set oStPoint = Nothing
    Set oBranchSegment = Nothing
    Set oGeomFactory = Nothing
    
' Place Port 1
    Dim oDir        As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    
    Dim oRadialOrient As AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oRadialOrient.Set 0, 0, 1
    
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
       
' Place Port 2
    Set oDir = New AutoMath.DVector
    oDir.Set 1, 0, 0
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oRadialOrient.Set 0, 0, 1

    oPortLocation.Set oPort2.x - dInsertionDepth * oDir.x, oPort2.y - dInsertionDepth * oDir.y, oPort2.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort2 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing
        
' Place Port 3
    Set oDir = New AutoMath.DVector
    oDir.Set 0, 0, 1
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oRadialOrient.Set 0, -1, 0

    oPortLocation.Set oPort3.x - dInsertionDepth * oDir.x, oPort3.y - dInsertionDepth * oDir.y, oPort3.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 3, oPortLocation, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oPort3 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPortLocation = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
