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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Tuesday 17, Apr 2007
'   Description:
'   This symbol is created for JIS-F7121 Water Filter to support multiple dimensional bases
'   as per CR 118311
'   The symbol supports the following dimensional bases
'   1.)Part Data Basis Value 240:-Water filter, outlet port at 180 degrees
'   2.)Part Data Basis Value 245:-Water filter, outlet port at 90 degrees, clockwise
'   3.)Part Data Basis Value 250:-Water filter, outlet port at 90 degrees, counter-clockwise

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double
    Dim parHeight As Double
    Dim parHeight1 As Double
    Dim parHeight2 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parFacetoCenter = arrayOfInputs(3)
    parHeight = arrayOfInputs(4)
    parHeight1 = arrayOfInputs(5) 'height of the Large vertical cylinder below the center-line.
    parHeight2 = arrayOfInputs(6) 'height of the Large vertical cylinder above the center-line.
    parInsulationThickness = arrayOfInputs(7)
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFacetoCenter As Double
    
    Select Case lPartDataBasis
        Case Is <= 1, 240
            parFacetoFace = arrayOfInputs(2)
            dFacetoCenter = parFacetoFace / 2
            
        Case 245, 250
            parFacetoCenter = arrayOfInputs(3)
            dFacetoCenter = parFacetoCenter
            
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    End Select
    

'   Place Output 1 (Vertical Cylinder)
    Dim objCylinder As Object
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

' Assumption: Length of the horizontal cylinder protrusion on either sides of the Large
'             vertical cylinder is twice of the flange thickness.
    Dim dCylinderDiam As Double
    Dim dFlangeThickness As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dFlangeThickness = flangeThick
' Check for the case where flange thickness is zero
    If CmpDblEqual(flangeThick, 0) Then dFlangeThickness = 0.1 * dFacetoCenter
        
    'Diameter of the larger cylinder = Face to Face - Twice Flange Thickness
    ' - Twice length of the smaller horizontal cylinders.
    dCylinderDiam = 2 * dFacetoCenter - 6 * dFlangeThickness
    
    oStPoint.Set 0, -parHeight1, 0
    oEnPoint.Set 0, parHeight2, 0
    
    Set objCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing

'   Place Output 2(Cylinder Top)
    Dim objCylTop As Object
    
    Dim dCylTopDia As Double
    Dim dTopCylThickness As Double

'   Assumptions
    dCylTopDia = 1.1 * dCylinderDiam
    dTopCylThickness = 0.25 * parHeight2
    
    oStPoint.Set 0, parHeight2, 0
    oEnPoint.Set 0, parHeight2 + dTopCylThickness, 0
    
    Set objCylTop = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylTopDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylTop
    Set objCylTop = Nothing
    
'   Place Output 3(Cover)
    Dim objCover As Object
    
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Dim oEllipseCenter As Object
    Set oEllipseCenter = New AutoMath.DPosition
    
    oEllipseCenter.Set 0, parHeight2 + dTopCylThickness, 0
    
    Dim dMajorRadius As Double
    Dim dMinorRadius As Double
    
    dMajorRadius = dCylTopDia / 2
    dMinorRadius = 0.2 * parHeight2
    'place a quarter ellipse
    Dim oQuarterEllipse As Object
    Set oQuarterEllipse = oGeometryFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
        Nothing, _
        oEllipseCenter.x, oEllipseCenter.y, oEllipseCenter.z, _
        0, 0, 1, _
        dMajorRadius, 0, 0, _
        dMinorRadius / dMajorRadius, _
        0, PI / 2)
        
'   Revolve the quarter ellipse to generate the semi-ellipsoid
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    
    oStPoint.Set 0, 0, 0
    oDir.Set 0, 1, 0
    Set objCover = PlaceRevolution(m_OutputColl, oQuarterEllipse, oDir, oStPoint, PI * 2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCover
    Set objCover = Nothing
    Set oEllipseCenter = Nothing
    Set oQuarterEllipse = Nothing
    Set oGeometryFactory = Nothing
    
'   Place Output 4(CoverTop)
    Dim objCoverTop As Object
    
    Dim dCoverTopDia As Double
    dCoverTopDia = 0.2 * dCylinderDiam
    
    oStPoint.Set 0, parHeight2 + dTopCylThickness, 0
    oEnPoint.Set 0, parHeight - parHeight1, 0
    
    Set objCoverTop = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCoverTopDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCoverTop
    Set objCoverTop = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -dFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dFacetoCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Select Case lPartDataBasis
        Case Is <= 1, 240
        oPlacePoint.Set dFacetoCenter + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
    
        Case 245
        oPlacePoint.Set 0, 0, -dFacetoCenter - sptOffset + depth
        oDir.Set 0, 0, -1
        
        Case 250
        oPlacePoint.Set 0, 0, dFacetoCenter + sptOffset - depth
        oDir.Set 0, 0, 1
    End Select
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dFacetoCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

