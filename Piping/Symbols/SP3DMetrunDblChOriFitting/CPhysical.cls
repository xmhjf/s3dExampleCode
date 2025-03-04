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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, Sep 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMDC2,
'   PDMS Category Name=IFRC). The symbol consists of 11  outputs, nine physical and
'   two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim pipeDiam       As Double
    Dim flangeThick     As Double
    Dim sptOffset      As Double
    Dim flangeDiam      As Double
    Dim depth          As Double
    
    Dim pipeDiam1       As Double
    Dim sptOffset1      As Double
    Dim depth1          As Double
    Dim pipeDiam2       As Double
    Dim sptOffset2      As Double
    Dim depth2          As Double
    
    
    Dim iOutput     As Double
    Dim ObjLHSPipe As Object
    Dim ObjRHSPipe As Object
    Dim ObjFittingBody As Object
    Dim ObjArmCylinder1 As Object
    Dim ObjArmBody2 As Object
    
    Dim parInstrumentLength      As Double
    Dim parInstrumentLength1     As Double
    Dim parInstrumentHeight      As Double
    Dim parInstrumentHeight1     As Double
    Dim parInstrumentLength2     As Double
    Dim parInstrumentLength3     As Double
    Dim parInstrumentWidth       As Double
    Dim parInstrumentWidth1      As Double
    Dim parArmLength             As Double
    Dim parNozzleOffset          As Double
    Dim parNozzleOffset1         As Double
    Dim parArmHeight             As Double
    Dim parArmHeight1            As Double
    Dim parInstrumentLength4     As Double
    Dim parInstrumentLength5     As Double
    Dim parInsulationThickness   As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim centerPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim Angle As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim LineStringPoints(0 To 26) As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)      'D2
    parInstrumentLength1 = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength2 = arrayOfInputs(6)     'D6
    parInstrumentLength3 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parArmLength = arrayOfInputs(10)            'D10
    parNozzleOffset = arrayOfInputs(11)         'D11
    parNozzleOffset1 = arrayOfInputs(12)        'D12
    parArmHeight = arrayOfInputs(13)            'D14
    parArmHeight1 = arrayOfInputs(14)           'D15
    parInstrumentLength4 = arrayOfInputs(15)    'D16
    parInstrumentLength5 = arrayOfInputs(16)    'D17
    parInsulationThickness = arrayOfInputs(17)
    
    iOutput = 0
    
'   Center position is assumed at the intersection of the flow axis and
'   verticle axis through the point at D11 offset from port3 center.

' Insert your code for output 3(LHSPipe)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1

'   Assumption: Instrument Left Hand Side/Right Hand Side flanges (At distances D2 and D3 respectively)
'   diameter is assumed to be 1.2 times and  flange thickness is 0.1 times corresponding pipe diameter.
'   Point 1
    LineStringPoints(0) = -parInstrumentLength - parInstrumentLength4
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = -parInstrumentLength - parInstrumentLength4
    LineStringPoints(4) = 0
    LineStringPoints(5) = pipeDiam1 / 2
    
'   Point 3
    LineStringPoints(6) = -parInstrumentLength - 0.1 * pipeDiam1
    LineStringPoints(7) = 0
    LineStringPoints(8) = pipeDiam1 / 2
    
'   Point 4
    LineStringPoints(9) = -parInstrumentLength - 0.1 * pipeDiam1
    LineStringPoints(10) = 0
    LineStringPoints(11) = 1.2 * pipeDiam1 / 2
    
'   Point 5
    LineStringPoints(12) = -parInstrumentLength
    LineStringPoints(13) = 0
    LineStringPoints(14) = 1.2 * pipeDiam1 / 2
    
'   Point 6
    LineStringPoints(15) = -parInstrumentLength + 0.1 * pipeDiam1
    LineStringPoints(16) = 0
    LineStringPoints(17) = 1.2 * pipeDiam1 / 2
    
'   Point 7
    LineStringPoints(18) = -parInstrumentLength + 0.1 * pipeDiam1
    LineStringPoints(19) = 0
    LineStringPoints(20) = pipeDiam1 / 2
    
'   Point 8
    LineStringPoints(21) = -parInstrumentLength2
    LineStringPoints(22) = 0
    LineStringPoints(23) = pipeDiam1 / 2
    
'   Point 9
    LineStringPoints(24) = -parInstrumentLength2
    LineStringPoints(25) = 0
    LineStringPoints(26) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set -(parInstrumentLength + parInstrumentLength4) / 2, 0, 0
    Angle = 2 * PI
    Set ObjLHSPipe = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, Angle, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSPipe
    Set ObjLHSPipe = Nothing
    
' Insert your code for output 4(RHSPipe)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2

'   Point 1
    LineStringPoints(0) = parInstrumentLength1 + parInstrumentLength5
    LineStringPoints(1) = 0
    LineStringPoints(2) = 0
    
'   Point 2
    LineStringPoints(3) = parInstrumentLength1 + parInstrumentLength5
    LineStringPoints(4) = 0
    LineStringPoints(5) = pipeDiam2 / 2
    
'   Point 3
    LineStringPoints(6) = parInstrumentLength1 + 0.1 * pipeDiam2
    LineStringPoints(7) = 0
    LineStringPoints(8) = pipeDiam2 / 2
    
'   Point 4
    LineStringPoints(9) = parInstrumentLength1 + 0.1 * pipeDiam2
    LineStringPoints(10) = 0
    LineStringPoints(11) = 1.2 * pipeDiam2 / 2
    
'   Point 5
    LineStringPoints(12) = parInstrumentLength1
    LineStringPoints(13) = 0
    LineStringPoints(14) = 1.2 * pipeDiam2 / 2
    
'   Point 6
    LineStringPoints(15) = parInstrumentLength1 - 0.1 * pipeDiam2
    LineStringPoints(16) = 0
    LineStringPoints(17) = 1.2 * pipeDiam2 / 2
    
'   Point 7
    LineStringPoints(18) = parInstrumentLength1 - 0.1 * pipeDiam2
    LineStringPoints(19) = 0
    LineStringPoints(20) = pipeDiam2 / 2
    
'   Point 8
    LineStringPoints(21) = parInstrumentLength3
    LineStringPoints(22) = 0
    LineStringPoints(23) = pipeDiam2 / 2
    
'   Point 9
    LineStringPoints(24) = parInstrumentLength3
    LineStringPoints(25) = 0
    LineStringPoints(26) = 0
    
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, LineStringPoints)
    axisVect.Set 1, 0, 0
    centerPoint.Set (parInstrumentLength1 + parInstrumentLength5) / 2, 0, 0
    Set ObjRHSPipe = PlaceRevolution(m_OutputColl, oLineString, axisVect, centerPoint, Angle, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSPipe
    Set ObjRHSPipe = Nothing
    
' Insert your code for output 5(FittingBody)
    stPoint.Set -parInstrumentLength2, -parInstrumentHeight, parInstrumentWidth
    enPoint.Set parInstrumentLength3, parInstrumentHeight1, -parInstrumentWidth1
    Set ObjFittingBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFittingBody
    Set ObjFittingBody = Nothing
    
' Insert your code for output 6(ArmCylinder1)
'   Assumption: Arm diameter is 0.1 times that of parArmLength
    Dim ArmDiameter As Double
    stPoint.Set 0, parArmHeight, -parInstrumentWidth1
    enPoint.Set 0, parArmHeight, -parArmLength
    ArmDiameter = parArmLength * 0.1
    Set ObjArmCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, ArmDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmCylinder1
    Set ObjArmCylinder1 = Nothing
    
' Insert your code for output 7(ArmBody2)
'   Assuming the Arm to be rectangular cross-section, with width being
'   0.2 times the parInstrumentLength2, and Thickness being 0.1times the parInstrumentLength2
    Dim arm2Width As Double
    Dim arm2Thickness As Double
    arm2Width = 0.6 * parInstrumentLength2
    arm2Thickness = 0.1 * parInstrumentLength2
    
    stPoint.Set arm2Width / 2, _
                        parArmHeight + ArmDiameter / 2 - parArmHeight1, _
                        -(parArmLength - arm2Thickness)

    enPoint.Set -arm2Width / 2, _
                        parArmHeight + ArmDiameter / 2, _
                        -parArmLength
    Set ObjArmBody2 = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmBody2
    Set ObjArmBody2 = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parInstrumentLength - parInstrumentLength4 - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parInstrumentLength1 + parInstrumentLength5 + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set -parNozzleOffset, 0, parInstrumentWidth + sptOffset - depth
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set parNozzleOffset1, 0, parInstrumentWidth + sptOffset - depth
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set centerPoint = Nothing
    Set axisVect = Nothing
 
    Dim ObjRMCurve As IJDObject
    Set ObjRMCurve = oLineString
    ObjRMCurve.Remove
    Set ObjRMCurve = Nothing
    Set oLineString = Nothing
 
    Set geomFactory = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
