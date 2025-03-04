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
'   Creation Date:  Thursday, Sep 9 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol details are taken from PDS Piping Component Data Reference Guide, at Page No. D-93
'   and SN=FS39A.
'   The Symbol is made up of only Physical aspect and consist of a Nozzle with length as outlet, Cylinder as
'   bottom Support, Cylinder as stem or Stand, Cone as a Eye Wash basin and Cylinder as inlet, Cylinder as
'   Eye wash Support, a Cylinder as Shower support and a Cone as Shower
'
'   Change History:
'   dd.mmm.yyyy     who              change description
'   -----------     ---              ------------------
'   10.Jun.2005     svsmylav TR-78430: PlaceCone function used to fail if the base radius is zero.
'                                    Swapped the base and top radii and swapped the base and top center points.
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

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
    
    Dim ObjStem As Object
    Dim ObjShowerBase As Object
    Dim ObjShowerSupLength As Object
    Dim ObjShower As Object
    Dim ObjWashSupLength As Object
    Dim ObjEyeWash As Object
    Dim ObjInlet As Object
    
    Dim parHeight1 As Double
    Dim parHeight2 As Double
    Dim parHeight3 As Double
    Dim parFacetoCenter As Double
    Dim parBaseDiameter As Double
    Dim parBaseThickness As Double
    Dim parShowerSupLength As Double
    Dim parRotation As Double
    Dim parShowerDiameter As Double
    Dim parWashHeight As Double
    Dim parWashSupLength As Double
    Dim parWashDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHeight1 = arrayOfInputs(2) 'D1
    parHeight2 = arrayOfInputs(3)     'D2
    parHeight3 = arrayOfInputs(4)     'D3
    parFacetoCenter = arrayOfInputs(5)           'D4
    parBaseDiameter = arrayOfInputs(6)          'D5
    parBaseThickness = arrayOfInputs(7)         'D6
    parShowerSupLength = arrayOfInputs(8)              'D7
    parRotation = arrayOfInputs(9)                     'D8
    parShowerDiameter = arrayOfInputs(10)   'D9
    parWashHeight = arrayOfInputs(11)        'D10
    parWashSupLength = arrayOfInputs(12)            'D11
    parWashDiameter = arrayOfInputs(13)      'D12
    
    iOutput = 0

 ' Insert your code for output 1(Stem)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    startPt.Set 0, 0, -(parHeight1 - parBaseThickness)
    endPt.Set 0, 0, parHeight2 + parHeight3 + pipeDiam / 2
    Set ObjStem = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 2(ShowerBase)
    startPt.Set 0, 0, -parHeight1
    endPt.Set 0, 0, -(parHeight1 - parBaseThickness)
    Set ObjShowerBase = PlaceCylinder(m_OutputColl, startPt, endPt, parBaseDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerBase
    Set ObjShowerBase = Nothing
    
 ' Insert your code for output 3(ShowerSupLength)
    startPt.Set 0, 0, parHeight2 + parHeight3
    endPt.Set parShowerSupLength, 0, parHeight2 + parHeight3
    Set ObjShowerSupLength = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerSupLength
    Set ObjShowerSupLength = Nothing
    
 ' Insert your code for output 4(Shower)
    Dim basecenter As AutoMath.DPosition
    Dim topcenter As AutoMath.DPosition
    Set basecenter = New AutoMath.DPosition
    Set topcenter = New AutoMath.DPosition
    
    topcenter.Set parShowerSupLength, 0, parHeight2 + parHeight3 - pipeDiam / 2
    basecenter.Set parShowerSupLength, 0, parHeight2 + parHeight3 - pipeDiam / 2 - 0.1524     ' 6" As defined in PDS Manual
    Set ObjShower = PlaceCone(m_OutputColl, basecenter, topcenter, parShowerDiameter / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjShower
    Set ObjShower = Nothing
    
 ' Insert your code for output 5(WashSupLength)
    startPt.Set 0, 0, parWashHeight - parHeight1
    endPt.Set parWashSupLength, 0, parWashHeight - parHeight1
    Set ObjWashSupLength = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWashSupLength
    Set ObjWashSupLength = Nothing
    
 ' Insert your code for output 6(EyeWash)
    startPt.Set parWashSupLength, 0, parWashHeight - parHeight1 + pipeDiam / 2 + 0.1524
    endPt.Set parWashSupLength, 0, parWashHeight - parHeight1 + pipeDiam / 2
    Set ObjEyeWash = PlaceCone(m_OutputColl, startPt, endPt, parWashDiameter / 2, 0, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEyeWash
    Set ObjEyeWash = Nothing
    Set basecenter = Nothing
    Set topcenter = Nothing
    
' Insert your code for output 7(Inlet)
    Dim Rotation As Double
    Rotation = PI + parRotation
    startPt.Set parFacetoCenter * Cos(Rotation), parFacetoCenter * Sin(Rotation), parHeight2
    endPt.Set 0, 0, parHeight2
    Set ObjInlet = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlet
    Set ObjInlet = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set (-parFacetoCenter - sptOffset + depth) * Cos(parRotation), _
                                (-parFacetoCenter - sptOffset + depth) * Sin(parRotation), 0
    oDir.Set -Cos(parRotation), -Sin(parRotation), 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFacetoCenter)
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
