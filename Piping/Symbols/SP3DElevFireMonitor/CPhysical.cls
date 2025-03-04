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
'   ProgID:         SP3DElevFireMonitor.CElevFireMonitor
'   Author:         MS
'   Creation Date:  Thursday, Sep 9 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: FS2b Page No. D-87 PDS Piping Component Data Reference Guide.
'    Symbol is created with Four Outputs
'    The five physical aspect outputs are created as follows:
'    ObjMonitorBody Using 'PlaceProjection' function,
'    ObjCannonBody , ObjConnBody - Using 'PlaceCylinder' function
'    ObjBaseBody -Using   'PlaceBox' and
'    One ObjNozzle object by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
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
    Dim ObjMonitorBody As Object
    Dim ObjCannonBody As Object
    Dim ObjConnBody As Object
    Dim ObjBaseBody As Object
    
    Dim parFacetoCenter As Double
    Dim parBaseBtoCenter As Double
    Dim parBaseTtoCenter As Double
    Dim parBaseWidth As Double
    Dim parBaseTtoHinge As Double
    Dim parBodyTtoHinge As Double
    Dim parBodyBtoHinge As Double
    Dim parBodyLtoCenter As Double
    Dim parBodyRtoCenter As Double
    Dim parBodyWidth As Double
    Dim parConnBodyDiameter As Double
    Dim parCannonLength As Double
    Dim parCannonDiameter As Double
    Dim parRotation As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)  'D1
    parBaseBtoCenter = arrayOfInputs(3) 'D2
    parBaseTtoCenter = arrayOfInputs(4) 'D3
    parBaseWidth = arrayOfInputs(5) 'D4
    parBaseTtoHinge = arrayOfInputs(6)  'D5
    parBodyTtoHinge = arrayOfInputs(7)  'D6
    parBodyBtoHinge = arrayOfInputs(8)  'D7
    parBodyLtoCenter = arrayOfInputs(9) 'D8
    parBodyRtoCenter = arrayOfInputs(10)    'D9
    parBodyWidth = arrayOfInputs(11)        'D10
    parConnBodyDiameter = arrayOfInputs(12) 'D11
    parCannonLength = arrayOfInputs(13) 'D12
    parCannonDiameter = arrayOfInputs(14) 'D13
    parRotation = arrayOfInputs(15) 'D14
    
    iOutput = 0
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

 ' Insert your code for output 1(Monitor Body)
    Dim linePts(0 To 14) As Double
    Dim optLine As IngrGeom3D.LineString3d
    Dim projVect As New AutoMath.DVector
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
       
    'Point 1
    linePts(0) = parBodyLtoCenter * Cos(parRotation) + (parBodyWidth / 2) * Sin(parRotation)
    linePts(1) = -(parBodyLtoCenter * Sin(parRotation) - (parBodyWidth / 2) * Cos(parRotation))
    linePts(2) = parBaseTtoCenter + parBaseTtoHinge + parBodyTtoHinge
    'Point2
    linePts(3) = parBodyLtoCenter * Cos(parRotation) - (parBodyWidth / 2) * Sin(parRotation)
    linePts(4) = -(parBodyLtoCenter * Sin(parRotation) + (parBodyWidth / 2) * Cos(parRotation))
    linePts(5) = parBaseTtoCenter + parBaseTtoHinge + parBodyTtoHinge
    'Point3
    linePts(6) = -(parBodyRtoCenter * Cos(parRotation) + (parBodyWidth / 2) * Sin(parRotation))
    linePts(7) = parBodyRtoCenter * Sin(parRotation) - (parBodyWidth / 2) * Cos(parRotation)
    linePts(8) = parBaseTtoCenter + parBaseTtoHinge + parBodyTtoHinge
    'Point4
    linePts(9) = -(parBodyRtoCenter * Cos(parRotation) - (parBodyWidth / 2) * Sin(parRotation))
    linePts(10) = parBodyRtoCenter * Sin(parRotation) + (parBodyWidth / 2) * Cos(parRotation)
    linePts(11) = parBaseTtoCenter + parBaseTtoHinge + parBodyTtoHinge
    'Point5
    linePts(12) = parBodyLtoCenter * Cos(parRotation) + (parBodyWidth / 2) * Sin(parRotation)
    linePts(13) = -(parBodyLtoCenter * Sin(parRotation) - (parBodyWidth / 2) * Cos(parRotation))
    linePts(14) = parBaseTtoCenter + parBaseTtoHinge + parBodyTtoHinge
       
    projVect.Set 0, 0, -1
    Set optLine = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, linePts)
    
    Set ObjMonitorBody = PlaceProjection(m_OutputColl, optLine, projVect, parBodyTtoHinge + parBodyBtoHinge, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMonitorBody
    Set ObjMonitorBody = Nothing
    Set projVect = Nothing
    Set geomFactory = Nothing
    
 ' Insert your code for output 2(Cannon Body)
    Dim CannonBodyAngle As Double
    Dim RefAngle As Double
    CannonBodyAngle = PI / 4
    RefAngle = parRotation - (PI / 2)
    
    stPoint.Set 0, 0, parBaseTtoCenter + parBaseTtoHinge
    enPoint.Set parCannonLength * Cos(CannonBodyAngle) * Cos(RefAngle), _
                            -(parCannonLength * Cos(CannonBodyAngle) * Sin(RefAngle)), _
                            parBaseTtoCenter + parBaseTtoHinge + parCannonLength * Sin(CannonBodyAngle)

    Set ObjCannonBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parCannonDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCannonBody
    Set ObjCannonBody = Nothing
    
 ' Insert your code for output 3(Connecting Body)
    stPoint.Set 0, 0, parBaseTtoCenter
    enPoint.Set 0, 0, parBaseTtoCenter + parBaseTtoHinge - parBodyBtoHinge
    
    Set ObjConnBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parConnBodyDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnBody
    Set ObjConnBody = Nothing
    
 ' Insert your code for output 4(Base Body)
 'The width(D4) and Length of the Base Body is assummed to be the same
    stPoint.Set -parBaseWidth / 2, parBaseWidth / 2, parBaseTtoCenter
    enPoint.Set parBaseWidth / 2, -parBaseWidth / 2, -parBaseBtoCenter
        
    Set ObjBaseBody = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBaseBody
    Set ObjBaseBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
        
 ' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0

    Dim NozzleLength As Double

    NozzleLength = parFacetoCenter - parBaseWidth / 2
    If NozzleLength < flangeThick Then NozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

