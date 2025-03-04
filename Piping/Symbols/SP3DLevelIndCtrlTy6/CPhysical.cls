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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No.: I324FZ, Page No. D-127 of PDS Piping Component Data Reference Guide.
'    This Symbol is created with Six Outputs, all of which are physical aspect outputs.
'    ObjFloatBody, ObjArm1, ObjArm2 using 'PlaceCylinder' function, and ObjIndicator using 'CreateProjection'.
'    Two ObjNozzle object by using 'CreateNozzleWithLength' function
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjFloatBody As Object
    Dim ObjArm1 As Object
    Dim ObjArm2 As Object
    Dim ObjIndicator As Object
    
    Dim parFacetoCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parFloatDiameter As Double
    Dim parFloatToptoCenter As Double
    Dim parFloatBottoCenter As Double
    Dim parArmtoCenter As Double
    Dim parIndicatorHeight1 As Double
    Dim parIndicatorHeight2 As Double
    Dim parAngle As Double
    Dim parArm1Length As Double
    Dim parArm2Length As Double
    Dim parIndicatorWidth As Double
    Dim parIndicatorLength As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'P2
    parOffset1 = arrayOfInputs(3)   'P3
    parOffset2 = arrayOfInputs(4)   'P4
    parFloatDiameter = arrayOfInputs(5) 'P5
    parFloatToptoCenter = arrayOfInputs(6)  'P6
    parFloatBottoCenter = arrayOfInputs(7)  'P7
    parArmtoCenter = arrayOfInputs(8) 'P8
    parIndicatorHeight1 = arrayOfInputs(9)  'P9
    parIndicatorHeight2 = arrayOfInputs(10)
    parAngle = arrayOfInputs(11)    'P11
    parArm1Length = arrayOfInputs(12)   'P12
    parArm2Length = arrayOfInputs(13)   'P13
    parIndicatorWidth = arrayOfInputs(14)   'P14
    parIndicatorLength = arrayOfInputs(15)  'P15

    iOutput = 0

 ' Insert your code for output 1(FloatBody)
     stPoint.Set 0, 0, -parFloatBottoCenter
    enPoint.Set 0, 0, parFloatToptoCenter
    Set ObjFloatBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFloatDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFloatBody
    Set ObjFloatBody = Nothing
    
 ' Insert your code for output 2(Arm1)
    Dim CylinderMeetPoint As Double
    Dim Arm1Dia As Double
    Dim Arm2Dia As Double
    Arm1Dia = 0.05
    Arm2Dia = 0.05

    CylinderMeetPoint = Sqr(((parFloatDiameter / 2) * (parFloatDiameter / 2)) - ((Arm1Dia / 2) * (Arm1Dia / 2)))
    
    stPoint.x = (parArm1Length + Arm2Dia / 2) * Cos(parAngle)
    stPoint.y = -(parArm1Length + Arm2Dia / 2) * Sin(parAngle)
    stPoint.z = parArmtoCenter
    
    enPoint.x = CylinderMeetPoint * Cos(parAngle)
    enPoint.y = -CylinderMeetPoint * Sin(parAngle)
    enPoint.z = parArmtoCenter
    Set ObjArm1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, Arm1Dia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm1
    Set ObjArm1 = Nothing
    
 ' Insert your code for output 3(Arm2)
     Dim Arm2EndPoint As Double
    'Arm2EndPoint refers to the length of the projection of end point of arm2 onto arm1
    Arm2EndPoint = parArm1Length - (parArm2Length - parIndicatorWidth - Arm1Dia / 2) * Tan(parAngle)
    
    stPoint.x = parArm1Length * Cos(parAngle)
    stPoint.y = -parArm1Length * Sin(parAngle)
    stPoint.z = parArmtoCenter
    
    enPoint.x = Arm2EndPoint * Cos(parAngle)
    enPoint.y = -Arm2EndPoint * Sin(parAngle) - ((parArm2Length - parIndicatorWidth - Arm1Dia / 2) / Cos(parAngle))
    enPoint.z = parArmtoCenter
    Set ObjArm2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, Arm2Dia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm2
    Set ObjArm2 = Nothing
    
 ' Insert your code for output 4(Indicator)
    Dim BoxPoint1 As Double
    Dim BoxPoint2 As Double
    Dim BoxPoint3 As Double
    Dim BoxPoint4 As Double
    
    BoxPoint1 = parArm1Length - (parIndicatorLength / 2) - ((parArm2Length - parIndicatorWidth - Arm1Dia / 2) * Tan(parAngle))
    BoxPoint2 = parArm1Length + (parIndicatorLength / 2) - ((parArm2Length - parIndicatorWidth - Arm1Dia / 2) * Tan(parAngle))
    BoxPoint3 = parArm1Length + (parIndicatorLength / 2) - ((parArm2Length - Arm1Dia / 2) * Tan(parAngle))
    BoxPoint4 = parArm1Length - (parIndicatorLength / 2) - ((parArm2Length - Arm1Dia / 2) * Tan(parAngle))
 
     Dim linePts(0 To 14) As Double
    Dim optLine As IngrGeom3D.LineString3d
    Dim projVect As New AutoMath.DVector
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
       
    'Point 1
    linePts(0) = BoxPoint1 * Cos(parAngle)
    linePts(1) = -(BoxPoint1 * Sin(parAngle)) - ((parArm2Length - parIndicatorWidth - Arm1Dia / 2) / Cos(parAngle))
    linePts(2) = parArmtoCenter + parIndicatorHeight1
    'Point2
    linePts(3) = BoxPoint2 * Cos(parAngle)
    linePts(4) = -(BoxPoint2 * Sin(parAngle)) - ((parArm2Length - parIndicatorWidth - Arm1Dia / 2) / Cos(parAngle))
    linePts(5) = parArmtoCenter + parIndicatorHeight1
    'Point3
    linePts(6) = BoxPoint3 * Cos(parAngle)
    linePts(7) = -(BoxPoint3 * Sin(parAngle)) - ((parArm2Length - Arm1Dia / 2) / Cos(parAngle))
    linePts(8) = parArmtoCenter + parIndicatorHeight1
    'Point4
    linePts(9) = BoxPoint4 * Cos(parAngle)
    linePts(10) = -(BoxPoint4 * Sin(parAngle)) - ((parArm2Length - Arm1Dia / 2) / Cos(parAngle))
    linePts(11) = parArmtoCenter + parIndicatorHeight1
    'Point5
    linePts(12) = BoxPoint1 * Cos(parAngle)
    linePts(13) = -(BoxPoint1 * Sin(parAngle)) - ((parArm2Length - parIndicatorWidth - Arm1Dia / 2) / Cos(parAngle))
    linePts(14) = parArmtoCenter + parIndicatorHeight1
        
    projVect.Set 0, 0, -1
    Set optLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, linePts)
    
    Set ObjIndicator = PlaceProjection(m_OutputColl, optLine, projVect, parIndicatorHeight1 + parIndicatorHeight2, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjIndicator
    Set ObjIndicator = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set projVect = Nothing
    Set geomFactory = Nothing
    Set optLine = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim dNozzleLength As Double
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, -parOffset2
    oDir.Set -1, 0, 0
    dNozzleLength = parFacetoCenter
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, parOffset1
    oDir.Set -1, 0, 0
    dNozzleLength = parFacetoCenter
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
       
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub
