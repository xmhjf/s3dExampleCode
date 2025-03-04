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
'    This symbol is made to function as a Vertical Downward Three-Way, reference Zhen Hua Catalog.
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
    
    Dim parActualWidth As Double
    Dim parActualDepth As Double
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
'   1) Vertical Segmented Downward Three Way, Side branch, Symmetrical, specified by Header Length, Branch Length and Tangent Length (Default Implementation).
'   2) Vertical Segmented Downward Three Way, Side branch, Asymmetrical, specified by Header Length, Tee Branch Length, Tangent Length 1, Tangent Length 2 and Tangent Length.
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
        
        Case Is <= 1, 83     'Vertical Segmented Downward Three Way, Side branch, Symmetrical, specified by
                             'Header Length, Branch Length and Tangent Length (Default Implementation).
            parTangentLength = oTrayPart.TangentLength
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dTangentLength3 = parTangentLength
            
        Case 85      'Vertical Segmented Downward Three Way, Side branch, Asymmetrical, specified by
                     ' Header Length, Tee Branch Length, Tangent Length 1, Tangent Length 2 and Tangent Length.
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
    
'   Retrieve Ports 1 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(1, oPartFclt, parActualWidth, parActualDepth)
        
'   Port 1 position
    Dim dHalfDepth1 As Double
    dHalfDepth1 = parActualDepth / 2
    
    Dim dHalfWidth1 As Double
    dHalfWidth1 = parActualWidth / 2
       
    Dim oPort1 As AutoMath.DPosition 'Port 1 center point
    Set oPort1 = New AutoMath.DPosition
    oPort1.Set -parHeaderLength / 2, 0, 0
    
    ' Defining points for creating the linestring at port 1.
    Dim dPort1S(0 To 11)  As Double
    ' Tray top edge at the side of the flat base surface.
    dPort1S(0) = oPort1.x
    dPort1S(1) = oPort1.y + dHalfWidth1
    dPort1S(2) = oPort1.z + dHalfDepth1
    ' Tray bottom edge at the side of the flat base surface.
    dPort1S(3) = oPort1.x
    dPort1S(4) = oPort1.y + dHalfWidth1
    dPort1S(5) = oPort1.z - dHalfDepth1
    ' Tray bottom edge at the side oppsoite to the flat base surface.
    dPort1S(6) = oPort1.x
    dPort1S(7) = oPort1.y - dHalfWidth1
    dPort1S(8) = oPort1.z - dHalfDepth1
    ' Tray top edge at the side oppsoite to the flat base surface.
    dPort1S(9) = oPort1.x
    dPort1S(10) = oPort1.y - dHalfWidth1
    dPort1S(11) = oPort1.z + dHalfDepth1
    
'   Port 2 position
    Dim oPort2 As AutoMath.DPosition 'Port 2 center point
    Set oPort2 = New AutoMath.DPosition
    oPort2.Set parHeaderLength / 2, 0, 0
    
    ' Defining points for creating the linestring at port 2.
    Dim dPort2S(0 To 11)  As Double
    ' Tray top edge at the side of the flat base surface.
    dPort2S(0) = oPort2.x
    dPort2S(1) = oPort2.y + dHalfWidth1
    dPort2S(2) = oPort2.z + dHalfDepth1
    ' Tray bottom edge at the side of the flat base surface.
    dPort2S(3) = oPort2.x
    dPort2S(4) = oPort2.y + dHalfWidth1
    dPort2S(5) = oPort2.z - dHalfDepth1
    ' Tray bottom edge at the side oppsoite to the flat base surface.
    dPort2S(6) = oPort2.x
    dPort2S(7) = oPort2.y - dHalfWidth1
    dPort2S(8) = oPort2.z - dHalfDepth1
    ' Tray top edge at the side oppsoite to the flat base surface.
    dPort2S(9) = oPort2.x
    dPort2S(10) = oPort2.y - dHalfWidth1
    dPort2S(11) = oPort2.z + dHalfDepth1
    
'   Port 3 position
    ' Retrieve Ports 3 properties: Actual Width and Depth
    Call RetrieveCableTrayPortProperties(3, oPartFclt, parActualWidth3, parActualDepth3)
    
    Dim oPort3 As AutoMath.DPosition 'Port 3 center point
    Set oPort3 = New AutoMath.DPosition
    oPort3.Set 0, dHalfWidth1 - parActualDepth3 / 2, -parBranchLength + dHalfDepth1
    
    ' Defining points for creating the linestring at port 3.
    Dim dPort3S()  As Double
    ReDim dPort3S(0 To 11)
    dPort3S(0) = oPort3.x - parActualWidth3 / 2
    dPort3S(1) = oPort3.y + parActualDepth3 / 2
    dPort3S(2) = oPort3.z
    
    dPort3S(3) = oPort3.x - parActualWidth3 / 2
    dPort3S(4) = oPort3.y - parActualDepth3 / 2
    dPort3S(5) = oPort3.z
    
    dPort3S(6) = oPort3.x + parActualWidth3 / 2
    dPort3S(7) = oPort3.y - parActualDepth3 / 2
    dPort3S(8) = oPort3.z

    dPort3S(9) = oPort3.x + parActualWidth3 / 2
    dPort3S(10) = oPort3.y + parActualDepth3 / 2
    dPort3S(11) = oPort3.z

'   Construct header front side.
    Dim oVerLineRightSideHeader As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oVerLineRightSideHeader = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPort1S(6), dPort1S(7), dPort1S(8), _
        dPort1S(9), dPort1S(10), dPort1S(11))
    
    Dim oHorLineAlongRightSideHeader As IngrGeom3D.Line3d
    Set oHorLineAlongRightSideHeader = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        dPort1S(6), dPort1S(7), dPort1S(8), _
        dPort2S(6), dPort2S(7), dPort2S(8))
    
    Dim objHeaderSideSurface As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
    Set objHeaderSideSurface = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
        oHorLineAlongRightSideHeader, oVerLineRightSideHeader, 0, 0, stnorm, ednorm, False)
              
' Set the output
    Dim ObjSurface As Object
    For Each ObjSurface In objHeaderSideSurface
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
    Next ObjSurface
    Set ObjSurface = Nothing
    Set objHeaderSideSurface = Nothing
    Set oVerLineRightSideHeader = Nothing
    Set oHorLineAlongRightSideHeader = Nothing

' Construct the Base surface adjoining port 1.
    'Defining Points for creating Base surface adjoining port 1.
    'Point at Base surface adjoining port 1 - in the Inner half i.e along +Y axis.
    Dim oPointBaseSurfPort1InHalf As AutoMath.DPosition
    Set oPointBaseSurfPort1InHalf = New AutoMath.DPosition
    oPointBaseSurfPort1InHalf.Set dPort1S(3) + dTangentLength1, dPort1S(4), dPort1S(5)
    
    'Point at Base surface adjoining port 1 - in the Outer half i.e along -Y axis.
    Dim oPointBaseSurfPort1OutHalf As AutoMath.DPosition
    Set oPointBaseSurfPort1OutHalf = New AutoMath.DPosition
    oPointBaseSurfPort1OutHalf.Set dPort1S(6) + dTangentLength1, dPort1S(7), dPort1S(8)
    
    Dim dPoints() As Double
    ReDim dPoints(0 To 11)
    
    dPoints(0) = oPointBaseSurfPort1InHalf.x
    dPoints(1) = oPointBaseSurfPort1InHalf.y
    dPoints(2) = oPointBaseSurfPort1InHalf.z
    
    dPoints(3) = oPointBaseSurfPort1OutHalf.x
    dPoints(4) = oPointBaseSurfPort1OutHalf.y
    dPoints(5) = oPointBaseSurfPort1OutHalf.z
    
    dPoints(6) = dPort1S(6)
    dPoints(7) = dPort1S(7)
    dPoints(8) = dPort1S(8)
    
    dPoints(9) = dPort1S(3)
    dPoints(10) = dPort1S(4)
    dPoints(11) = dPort1S(5)

    Dim BottomSurfacePort1 As IngrGeom3D.Plane3d
    Set BottomSurfacePort1 = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), BottomSurfacePort1
    Set BottomSurfacePort1 = Nothing
       
 ' Construct the Base surface adjoining port 2.
    ' Defining Points for creating Base surface adjoining port 2.
    'Point at Base surface adjoining port 2 - in the Inner half i.e along +Y axis.
    Dim oPointBaseSurfPort2InHalf As AutoMath.DPosition
    Set oPointBaseSurfPort2InHalf = New AutoMath.DPosition
    oPointBaseSurfPort2InHalf.Set dPort2S(3) - dTangentLength2, dPort2S(4), dPort2S(5)
    
    'Point at Base surface adjoining port 2 - in the Outer half i.e along -Y axis.
    Dim oPointBaseSurfPort2OutHalf As AutoMath.DPosition
    Set oPointBaseSurfPort2OutHalf = New AutoMath.DPosition
    oPointBaseSurfPort2OutHalf.Set dPort2S(6) - dTangentLength2, dPort2S(7), dPort2S(8)
    
    dPoints(0) = oPointBaseSurfPort2InHalf.x
    dPoints(1) = oPointBaseSurfPort2InHalf.y
    dPoints(2) = oPointBaseSurfPort2InHalf.z
    
    dPoints(3) = dPort2S(3)
    dPoints(4) = dPort2S(4)
    dPoints(5) = dPort2S(5)
    
    dPoints(6) = dPort2S(6)
    dPoints(7) = dPort2S(7)
    dPoints(8) = dPort2S(8)
    
    dPoints(9) = oPointBaseSurfPort2OutHalf.x
    dPoints(10) = oPointBaseSurfPort2OutHalf.y
    dPoints(11) = oPointBaseSurfPort2OutHalf.z
        
    Dim BottomSurfacePort2 As IngrGeom3D.Plane3d
    Set BottomSurfacePort2 = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)
                            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), BottomSurfacePort2
    Set BottomSurfacePort2 = Nothing
    
' Construct the (Branch Portion) surface at the back.
    ' Defining Points at Branch mouth before the tangential enclosure.
    'Point at left in the Inner half i.e along +Y axis.
    Dim oPointLeftInHalf As AutoMath.DPosition
    Set oPointLeftInHalf = New AutoMath.DPosition
    oPointLeftInHalf.Set dPort3S(0), dPort3S(1), dPort3S(2) + dTangentLength3
    
    'Point at Right in the Inner half i.e along +Y axis.
    Dim oPointRightInHalf As AutoMath.DPosition
    Set oPointRightInHalf = New AutoMath.DPosition
    oPointRightInHalf.Set dPort3S(9), dPort3S(10), dPort3S(11) + dTangentLength3
    
    'Point at left in the Outer half i.e along -Y axis.
    Dim oPointLeftOutHalf As AutoMath.DPosition
    Set oPointLeftOutHalf = New AutoMath.DPosition
    oPointLeftOutHalf.Set dPort3S(3), dPort3S(4), dPort3S(5) + dTangentLength3
    
    'Point at Right in the Outer half i.e along -Y axis.
    Dim oPointRightOutHalf As AutoMath.DPosition
    Set oPointRightOutHalf = New AutoMath.DPosition
    oPointRightOutHalf.Set dPort3S(6), dPort3S(7), dPort3S(8) + dTangentLength3
    
    ReDim dPoints(0 To 23)
    
    dPoints(0) = dPort1S(0)
    dPoints(1) = dPort1S(1)
    dPoints(2) = dPort1S(2)
    
    dPoints(3) = dPort2S(0)
    dPoints(4) = dPort2S(1)
    dPoints(5) = dPort2S(2)
    
    dPoints(6) = dPort2S(3)
    dPoints(7) = dPort2S(4)
    dPoints(8) = dPort2S(5)
    
    dPoints(9) = oPointBaseSurfPort2InHalf.x
    dPoints(10) = oPointBaseSurfPort2InHalf.y
    dPoints(11) = oPointBaseSurfPort2InHalf.z
    
    dPoints(12) = oPointRightInHalf.x
    dPoints(13) = oPointRightInHalf.y
    dPoints(14) = oPointRightInHalf.z
    
    dPoints(15) = oPointLeftInHalf.x
    dPoints(16) = oPointLeftInHalf.y
    dPoints(17) = oPointLeftInHalf.z
    
    dPoints(18) = oPointBaseSurfPort1InHalf.x
    dPoints(19) = oPointBaseSurfPort1InHalf.y
    dPoints(20) = oPointBaseSurfPort1InHalf.z
    
    dPoints(21) = dPort1S(3)
    dPoints(22) = dPort1S(4)
    dPoints(23) = dPort1S(5)
    
    Dim objBranchBase As IngrGeom3D.Plane3d
    Set objBranchBase = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            8, dPoints)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBranchBase
    Set objBranchBase = Nothing
    
'Create Inclined Surface at Left.
    ReDim dPoints(0 To 11)
    
    dPoints(0) = oPointBaseSurfPort1InHalf.x
    dPoints(1) = oPointBaseSurfPort1InHalf.y
    dPoints(2) = oPointBaseSurfPort1InHalf.z
    
    dPoints(3) = oPointBaseSurfPort1OutHalf.x
    dPoints(4) = oPointBaseSurfPort1OutHalf.y
    dPoints(5) = oPointBaseSurfPort1OutHalf.z
    
    dPoints(6) = oPointLeftOutHalf.x
    dPoints(7) = oPointLeftOutHalf.y
    dPoints(8) = oPointLeftOutHalf.z
    
    dPoints(9) = oPointLeftInHalf.x
    dPoints(10) = oPointLeftInHalf.y
    dPoints(11) = oPointLeftInHalf.z
    
    Dim objLeftInclinedSurface As IngrGeom3D.Plane3d
    Set objLeftInclinedSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeftInclinedSurface
    Set objLeftInclinedSurface = Nothing
    Set oPointLeftInHalf = Nothing
    Set oPointBaseSurfPort1InHalf = Nothing
    
'Create Inclined Surface at Right.

    dPoints(0) = oPointBaseSurfPort2InHalf.x
    dPoints(1) = oPointBaseSurfPort2InHalf.y
    dPoints(2) = oPointBaseSurfPort2InHalf.z
    
    dPoints(3) = oPointBaseSurfPort2OutHalf.x
    dPoints(4) = oPointBaseSurfPort2OutHalf.y
    dPoints(5) = oPointBaseSurfPort2OutHalf.z
    
    dPoints(6) = oPointRightOutHalf.x
    dPoints(7) = oPointRightOutHalf.y
    dPoints(8) = oPointRightOutHalf.z
    
    dPoints(9) = oPointRightInHalf.x
    dPoints(10) = oPointRightInHalf.y
    dPoints(11) = oPointRightInHalf.z
    
    Dim objRightInclinedSurface As IngrGeom3D.Plane3d
    Set objRightInclinedSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRightInclinedSurface
    Set objRightInclinedSurface = Nothing
    Set oPointBaseSurfPort2InHalf = Nothing
    Set oPointRightInHalf = Nothing
    
'Create Inclined Surface - front.

    dPoints(0) = oPointBaseSurfPort2OutHalf.x
    dPoints(1) = oPointBaseSurfPort2OutHalf.y
    dPoints(2) = oPointBaseSurfPort2OutHalf.z
    
    dPoints(3) = oPointRightOutHalf.x
    dPoints(4) = oPointRightOutHalf.y
    dPoints(5) = oPointRightOutHalf.z
    
    dPoints(6) = oPointLeftOutHalf.x
    dPoints(7) = oPointLeftOutHalf.y
    dPoints(8) = oPointLeftOutHalf.z
    
    dPoints(9) = oPointBaseSurfPort1OutHalf.x
    dPoints(10) = oPointBaseSurfPort1OutHalf.y
    dPoints(11) = oPointBaseSurfPort1OutHalf.z
    
    Dim objFrontInclinedSurface As IngrGeom3D.Plane3d
    Set objFrontInclinedSurface = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, _
                            4, dPoints)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrontInclinedSurface
    Set objFrontInclinedSurface = Nothing
    Set oPointBaseSurfPort2OutHalf = Nothing
    Set oPointRightOutHalf = Nothing
    Set oPointLeftOutHalf = Nothing
    Set oPointBaseSurfPort1OutHalf = Nothing
    
'   Tangent at port 3

    ReDim Preserve dPort3S(0 To 14)
    dPort3S(12) = dPort3S(0)
    dPort3S(13) = dPort3S(1)
    dPort3S(14) = dPort3S(2)
    
    Dim oProjVector As AutoMath.DVector
    Set oProjVector = New AutoMath.DVector
    oProjVector.Set 0, 0, 1
    
    Dim oLineString As IngrGeom3D.LineString3d
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPort3S)
    
    Dim objTangentPort3 As Object
    Set objTangentPort3 = PlaceProjection(m_OutputColl, oLineString, oProjVector, dTangentLength3, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTangentPort3
    Set objTangentPort3 = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing
    Set oProjVector = Nothing
    
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
    oDir.Set 0, 0, -1
    Set oRadialOrient = New AutoMath.DVector
    Set oPortLocation = New AutoMath.DPosition
    oRadialOrient.Set 0, 1, 0

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
