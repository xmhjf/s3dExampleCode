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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No.: V41, Page No. D-25 of PDS Piping Component Data Reference Guide.
'   This Symbol is created with Eight Outputs of which one is insualation aspect output.
'   Other seven are physical aspect outputs created as follows:
'    Ellipsoid - By revoling an elliptical arc,
'    BodyCone1, BodyCone2 using 'PlaceCone' function,
'    Arrow1 and Arrow2 using 'PlaceRevolution' function,
'    two ObjNozzle objects  by using 'CreateNozzle' function.'
'   Change History:

'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   01.Aug.2006     KKK       TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'   02.Jan.2006     KKK       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   18.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
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
    Dim ObjEllipsoid As Object
    Dim ObjBodyCone1  As Object
    Dim ObjBodyCone2 As Object
    Dim ObjArrow1 As Object
    Dim ObjArrow2 As Object

    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double

    Dim axis As AutoMath.DVector
    Set axis = New AutoMath.DVector
    Dim centPoint As AutoMath.DPosition
    Set centPoint = New AutoMath.DPosition
    Dim stPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    Dim ellipseMajorX   As Double
    Dim ellipseMajorY   As Double
    Dim ellipseMajorZ   As Double
    Dim ellipseMMRatio  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parFacetoCenter = arrayOfInputs(4)
'    parFace1toCenter = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)
 
    iOutput = 0

' Insert your code for output 3(Ellipsoid)
    Dim ObjEllipticalArc As IngrGeom3D.EllipticalArc3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
   
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): Face-to-face dimension basis
'   (Part data Basis value -10): Face-to-center dimension basis
'   (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    
    Case 10
        parFacetoCenter = arrayOfInputs(4)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parFace1toCenter = arrayOfInputs(5)
        parFace2toCenter = arrayOfInputs(6)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
 
    Case Else
        GoTo ErrorLabel:
        
    End Select

'   Consider Ellipse Majoj Axis is 0.6 times the larger face to center distance /2 and
'   Minor axis as 0.4 times the same center distance.
    ellipseMMRatio = 0.4 / 0.6
    ellipseMajorX = (0.6 * (dFace1toCenter + dFace2toCenter) / 2) / 2
    ellipseMajorY = 0
    ellipseMajorZ = 0
    Set ObjEllipticalArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            Nothing, 0, 0, 0, _
                                            0, 0, 1, ellipseMajorX, ellipseMajorY, ellipseMajorZ, _
                                            ellipseMMRatio, 0, PI)
    centPoint.Set 0, 0, 0
    axis.Set 1, 0, 0
    Set ObjEllipsoid = PlaceRevolution(m_OutputColl, ObjEllipticalArc, axis, _
                        centPoint, 2 * PI, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipsoid
    Set ObjEllipsoid = Nothing

    Set centPoint = Nothing
    Set axis = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = ObjEllipticalArc
    Set ObjEllipticalArc = Nothing
    sp3dElem1.Remove
    Set sp3dElem1 = Nothing

' Insert your code for output 4(Cone - Port1 Side)
    stPoint.Set -dFace1toCenter + flangeThick, 0, 0
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing

 ' Insert your code for output 5(Cone - Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    stPoint.Set dFace2toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Insert your code for output 6( Arrow 1 for Flow Direction)
'   Create a Line String and Revolve it
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dblEndPts(0 To 14) As Double
    Dim revVector As New AutoMath.DVector
    Dim revCenterPt As New AutoMath.DPosition

    Dim ArrowCylDia As Double
    Dim ArrowConeDia  As Double
    Dim ArrowConeHeight As Double
    Dim ArrowtoCenter As Double

    ArrowCylDia = 0.1 * pipeDiam
    ArrowConeDia = 0.2 * pipeDiam
    ArrowConeHeight = 2 * ArrowConeDia
    ArrowtoCenter = 1.25 * pipeDiam

'   Point 1
    dblEndPts(0) = (dFace1toCenter + dFace2toCenter) / 4
    dblEndPts(1) = -ArrowtoCenter
    dblEndPts(2) = 0
'   Point 2
    dblEndPts(3) = (dFace1toCenter + dFace2toCenter) / 4 - ArrowConeHeight
    dblEndPts(4) = -ArrowtoCenter + ArrowConeDia / 2
    dblEndPts(5) = 0
'   Point 3
    dblEndPts(6) = (dFace1toCenter + dFace2toCenter) / 4 - ArrowConeHeight
    dblEndPts(7) = -ArrowtoCenter + ArrowCylDia / 2
    dblEndPts(8) = 0
'   Point 4
    dblEndPts(9) = -(dFace1toCenter + dFace2toCenter) / 4
    dblEndPts(10) = -ArrowtoCenter + ArrowCylDia / 2
    dblEndPts(11) = 0
'   Point 5
    dblEndPts(12) = -(dFace1toCenter + dFace2toCenter) / 4
    dblEndPts(13) = -ArrowtoCenter
    dblEndPts(14) = 0

    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblEndPts)
    revVector.Set 1, 0, 0
    revCenterPt.Set 0, -ArrowtoCenter, 0
    Set ObjArrow1 = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArrow1
    Set ObjArrow1 = Nothing
    Set oLineStr = Nothing

' Insert your code for output 6( Arrow 2 for Flow Direction)
    Dim i As Integer
    For i = 1 To 5
        dblEndPts(3 * i - 2) = -dblEndPts(3 * i - 2)    ' Change the sign of Y-coordinate
    Next i
    Set oLineStr = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblEndPts)
    Set geomFactory = Nothing
    revCenterPt.Set 0, ArrowtoCenter, 0
    Set ObjArrow2 = PlaceRevolution(m_OutputColl, oLineStr, revVector, revCenterPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArrow2
    Set ObjArrow2 = Nothing
    Set oLineStr = Nothing
    Set revVector = Nothing
    Set revCenterPt = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set dFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
End Sub
