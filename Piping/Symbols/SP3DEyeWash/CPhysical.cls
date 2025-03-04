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
'   ProgID:         SP3DEyeWash.CEyeWash
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 9 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol details are taken from PDS Piping Component Data Reference Guide, at Page No. D-93
'    and SN=FS38A.
'    The Symbol is made up of only Physical aspect and consist of a Nozzle with length as outlet, Cylinder as
'    bottom Support, Cylinder as stem or Stand, Cone as a Eye Wash basin and Cylinder as inlet.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjShowerCone As Object
    Dim ObjStem As Object
    Dim ObjInlet As Object
    Dim ObjOutlet As Object
    Dim ObjBase As Object
    
    Dim parHeight1 As Double
    Dim parHeight2 As Double
    Dim parHeight3 As Double
    Dim parWashDepth As Double
    Dim parBaseDiameter As Double
    Dim parBaseThickness As Double
    Dim parWashDiameter As Double
    Dim parStemDiameter As Double
    Dim parFace1toCenter As Double
    Dim parInletLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHeight1 = arrayOfInputs(2)   'D1
    parHeight2 = arrayOfInputs(3)   'D2
    parHeight3 = arrayOfInputs(4)   'D3
    parWashDepth = arrayOfInputs(5)   'D4
    parBaseDiameter = arrayOfInputs(6)  'D5
    parBaseThickness = arrayOfInputs(7) 'D6
    parWashDiameter = arrayOfInputs(8)    'D7
    parStemDiameter = arrayOfInputs(9)      'D8
    parFace1toCenter = arrayOfInputs(10)        'D9
    parInletLength = arrayOfInputs(11)          'D10
    
    iOutput = 0

 ' Insert your code for output 1(ShowerCone)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim baseCen As AutoMath.DPosition
    Dim topCen As AutoMath.DPosition
    Set baseCen = New AutoMath.DPosition
    Set topCen = New AutoMath.DPosition
    
    baseCen.Set 0, 0, parHeight2 + parHeight3
    topCen.Set 0, 0, parHeight2 + parHeight3 + parWashDepth
    Set ObjShowerCone = PlaceCone(m_OutputColl, baseCen, topCen, parStemDiameter / 2, parWashDiameter / 2, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjShowerCone
    Set ObjShowerCone = Nothing
    Set baseCen = Nothing
    Set topCen = Nothing
    
 ' Insert your code for output 2(Stem)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    startPt.Set 0, 0, -(parHeight1 - parBaseThickness)
    endPt.Set 0, 0, parHeight2 + parHeight3
    Set ObjStem = PlaceCylinder(m_OutputColl, startPt, endPt, parStemDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
 ' Insert your code for output 3(Inlet)
    startPt.Set -parInletLength, 0, parHeight2
    endPt.Set 0, 0, parHeight2
    Set ObjInlet = PlaceCylinder(m_OutputColl, startPt, endPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlet
    Set ObjInlet = Nothing
    
 ' Insert your code for output 5(Base)
    startPt.Set 0, 0, -parHeight1
    endPt.Set 0, 0, -parHeight1 + parBaseThickness
    Set ObjBase = PlaceCylinder(m_OutputColl, startPt, endPt, parBaseDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBase
    Set ObjBase = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)
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
