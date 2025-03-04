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
'   Author:         ACM
'   Creation Date:  Friday, Oct 25 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPR1A),If this is
'    rotated through 180 deg about stem axis (Operator Angle ) then we get (PDS Eden Name = IPL1D) symbol.
'   The symbol consists of Nine outputs,Eight physical outputs
'   and One Insulation Output.
'
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1          As Double
    Dim depth2          As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjStemRod As Object
    Dim ObjValveCylinder As Object
    Dim ObjActuatorRect As Object
    Dim ObjActuatorCyl As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorWidth As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorLength2 As Double
    Dim parInsulationThickness As Double
    Dim parRotation As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) ''D2
    parFacetoFace = arrayOfInputs(3) ''D3''
    parCylHeight = arrayOfInputs(4) ''D4
    parActuatorWidth = arrayOfInputs(5) ''D5
    parActuatorDiameter = arrayOfInputs(6) ''D6
    parActuatorHeight = arrayOfInputs(7) ''D7
    parActuatorHeight1 = arrayOfInputs(8) ''D10
    parActuatorHeight2 = arrayOfInputs(9) ''D11
    parActuatorLength = arrayOfInputs(10) ''D12
    parActuatorLength1 = arrayOfInputs(11) ''D14
    parActuatorWidth1 = arrayOfInputs(12) ''D25
    parActuatorWidth2 = arrayOfInputs(13) ''D26
    parActuatorLength2 = arrayOfInputs(14) ''D28
    parInsulationThickness = arrayOfInputs(15)
    parRotation = arrayOfInputs(16)
    
    iOutput = 0

 ' Insert your code for output 1(Conical Valve Body 1)
   
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
         
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
         
    stPoint.Set -parFace1toCenter + flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 2( Conical Valve Body 2 )
   
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
         
    stPoint.Set parFacetoFace - parFace1toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
        
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Stem Body )
   Dim StemDia As Double
   
   StemDia = 0.05
   stPoint.Set 0, 0, 0
   enPoint.Set 0, parActuatorHeight - parActuatorHeight2, 0
 
   Set ObjStemRod = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemRod
    Set ObjStemRod = Nothing
    
 ' Insert your code for output 4(Valve Cyledrical Body)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValveCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCylinder
    Set ObjValveCylinder = Nothing
    
 ' Insert your code for output 5(Actuator Rectangular box )
    Dim LineStrPts(0 To 14)  As Double
    Dim oLString As IngrGeom3D.LineString3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Dim axisVector As New AutoMath.DVector
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
'Point 1 (Bottom Left)
    LineStrPts(0) = -(parActuatorLength2 + parActuatorLength1)
    LineStrPts(1) = parActuatorHeight + parActuatorHeight1
    LineStrPts(2) = -parActuatorWidth1
    
'Point 2 (Bottom Right)
    LineStrPts(3) = (parActuatorLength - parActuatorLength2)
    LineStrPts(4) = parActuatorHeight + parActuatorHeight1
    LineStrPts(5) = -parActuatorWidth1
    
'Point 3 (Top Right)
    LineStrPts(6) = parActuatorLength - parActuatorLength2
    LineStrPts(7) = parActuatorHeight - parActuatorHeight2
    LineStrPts(8) = -parActuatorWidth1
    
'Point 4 (Top Left)
    LineStrPts(9) = -(parActuatorLength1 + parActuatorLength2)
    LineStrPts(10) = parActuatorHeight - parActuatorHeight2
    LineStrPts(11) = -parActuatorWidth1
    
'Point 5 (Initial Point (Bottom Left))
    LineStrPts(12) = -(parActuatorLength2 + parActuatorLength1)
    LineStrPts(13) = parActuatorHeight + parActuatorHeight1
    LineStrPts(14) = -parActuatorWidth1
    
    
    Dim ProjLength As Double
    ProjLength = parActuatorWidth1 + parActuatorWidth2
    Set oLString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)

    axisVector.Set 0, 0, 1
    Set ObjActuatorRect = PlaceProjection(m_OutputColl, oLString, axisVector, ProjLength, True)

   'Rotation
    Dim oTransMat As IJDT4x4
    Set oTransMat = New AutoMath.DT4x4
    Dim oAboutVec As IJDVector
    If oAboutVec Is Nothing Then Set oAboutVec = New DVector
    oAboutVec.Set 0, 1, 0
    oTransMat.Rotate parRotation, oAboutVec
    ObjActuatorRect.Transform oTransMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorRect
    Set ObjActuatorRect = Nothing
    
 ' Insert your code for output 6(Actuator Cylendrical Body)
  
   stPoint.Set -parActuatorLength2, parActuatorHeight, -parActuatorWidth1
   enPoint.Set -parActuatorLength2, parActuatorHeight, -parActuatorWidth
   
   Set ObjActuatorCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorDiameter, True)

   ObjActuatorCyl.Transform oTransMat
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCyl
    Set ObjActuatorCyl = Nothing
    
' Place Nozzle 1
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
   
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

