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
'    Symbol Model No. is: I324EZ Page No. 369 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjFloatBody As Object
    Dim ObjArm1 As Object
    Dim ObjArm2 As Object
    Dim ObjIndicator As Object
    
    Dim parNozzle3toCenter As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
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
    Dim ArmDiameter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace2toCenter = arrayOfInputs(2)         'P2
    parNozzle3toCenter = arrayOfInputs(3)   'P3
    parFacetoCenter = arrayOfInputs(4)         'P4
    parFace1toCenter = arrayOfInputs(5)         'P5
    parFloatDiameter = arrayOfInputs(6)         'P6
    parFloatToptoCenter = arrayOfInputs(7)      'P7
    parFloatBottoCenter = arrayOfInputs(8)      'P8
    parArmtoCenter = arrayOfInputs(9)           'P9
    parIndicatorHeight1 = arrayOfInputs(10)           'P10
    parIndicatorHeight2 = arrayOfInputs(11)         'P11
    parAngle = arrayOfInputs(12)                    'P12
    parArm1Length = arrayOfInputs(13)           'P13
    parArm2Length = arrayOfInputs(14)           'P14
    parIndicatorWidth = arrayOfInputs(15)       'P15
    parIndicatorLength = arrayOfInputs(16)      'P16
    
    iOutput = 0
    ArmDiameter = 0.05
 ' Insert your code for output 1(FloatBody)
    
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, 0, -parFloatBottoCenter
    endPoint.Set 0, 0, parFloatToptoCenter
    Set ObjFloatBody = PlaceCylinder(m_OutputColl, startPoint, endPoint, parFloatDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFloatBody
    Set ObjFloatBody = Nothing
    
 ' Insert your code for output 2(Arm1)
    Dim Arm1Angle As Double
    Arm1Angle = PI + parAngle
    startPoint.Set 0, 0, parArmtoCenter
    endPoint.Set (parArm1Length + ArmDiameter / 2) * Cos(Arm1Angle), _
                            (parArm1Length + ArmDiameter / 2) * Sin(Arm1Angle), _
                            parArmtoCenter
    Set ObjArm1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, ArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm1
    Set ObjArm1 = Nothing
    
 ' Insert your code for output 3(Arm2)
    Dim Arm2Angle As Double
    Arm2Angle = 1.5 * PI + parAngle
    startPoint.Set parArm1Length * Cos(Arm1Angle), _
                            parArm1Length * Sin(Arm1Angle), _
                            parArmtoCenter
    endPoint.Set startPoint.x + (parArm2Length - parIndicatorWidth - ArmDiameter / 2) * Sin(parAngle), _
                            startPoint.y - (parArm2Length - parIndicatorWidth - ArmDiameter / 2) * Cos(parAngle), _
                            parArmtoCenter
    Set ObjArm2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, ArmDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArm2
    Set ObjArm2 = Nothing
    Set startPoint = Nothing
    
 ' Insert your code for output 4(Indicator)
    Dim objLineStr As IngrGeom3D.LineString3d
    Dim dblLineStrPts(0 To 14) As Double
    Dim indiAngle As Double
    Dim projVect As AutoMath.DVector
    Set projVect = New AutoMath.DVector
    Dim geometryfactory As geometryfactory
    Set geometryfactory = New geometryfactory
    
    indiAngle = 2 * PI + parAngle
    projVect.Set Cos(Arm2Angle), Sin(Arm2Angle), 0
    
    'Point 1
    dblLineStrPts(0) = endPoint.x + Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(1) = endPoint.y + Sin(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(2) = parArmtoCenter + parIndicatorHeight1
    'Point 2
    dblLineStrPts(3) = endPoint.x + Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(4) = endPoint.y + Sin(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(5) = parArmtoCenter - parIndicatorHeight2
    'Point 3
    dblLineStrPts(6) = endPoint.x - Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(7) = endPoint.y - Sin(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(8) = parArmtoCenter - parIndicatorHeight2
    'Point 4
    dblLineStrPts(9) = endPoint.x - Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(10) = endPoint.y - Sin(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(11) = parArmtoCenter + parIndicatorHeight1
    'Point 5
    dblLineStrPts(12) = endPoint.x + Cos(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(13) = endPoint.y + Sin(indiAngle) * parIndicatorLength / 2
    dblLineStrPts(14) = parArmtoCenter + parIndicatorHeight1
    
    Set objLineStr = geometryfactory.LineStrings3d.CreateByPoints(Nothing, 5, dblLineStrPts)
    
    Set ObjIndicator = PlaceProjection(m_OutputColl, objLineStr, projVect, parIndicatorWidth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjIndicator
    Set ObjIndicator = Nothing
    Set projVect = Nothing
    Dim objTmp As IJDObject
    Set objTmp = objLineStr
    Set objLineStr = Nothing
    objTmp.Remove
    Set endPoint = Nothing

 ' Insert your code for output 5(Nozzle 1 withLength)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim NozzleLength As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    NozzleLength = parFacetoCenter - parFloatBottoCenter
    oPlacePoint.Set 0, 0, -parFacetoCenter - sptOffset + depth
    oDir.Set 0, 0, -1
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    NozzleLength = parFace1toCenter - parFloatToptoCenter
    oPlacePoint.Set 0, 0, parFace1toCenter + sptOffset - depth
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    NozzleLength = parFace2toCenter
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, -parNozzle3toCenter
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set geometryfactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub
