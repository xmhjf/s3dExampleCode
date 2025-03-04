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
'   Author       :  PK
'   Creation Date:  Friday, Oct 05 2007
'   Description  :  Straub Flex Type Coupling
'   Change History:
'   dd.mmm.yyyy           who                      change description
'   -----------          -----                      ------------------
'   04.Oct.2007           PK                        CR-126631:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Dim PI As Double

Private Sub Class_Initialize()

     PI = 4 * Atn(1)
     
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
     Dim oPartFclt       As PartFacelets.IJDPart
     Dim pipeDiam        As Double
     Dim flangeThick     As Double
     Dim cptOffset       As Double
     Dim flangeDiam      As Double
     Dim depth           As Double
    
     Dim iOutput                 As Double
     Dim parFacetoFace           As Double
     Dim parDiameter             As Double
     Dim parPorttoPortOffset     As Double
     Dim parBoltEndDimension     As Double
     Dim parInsulationThickness  As Double
   
' Inputs
    
     Set oPartFclt = arrayOfInputs(1)
     parFacetoFace = arrayOfInputs(2)
     parDiameter = arrayOfInputs(3)
     parPorttoPortOffset = arrayOfInputs(4)
     parBoltEndDimension = arrayOfInputs(5)
     parInsulationThickness = arrayOfInputs(6)
     iOutput = 0
     
     'Cylinder code
     Dim oCylinder   As Object
     Dim oStPoint    As AutoMath.DPosition
     Dim oEnpoint    As AutoMath.DPosition
     
     Set oStPoint = New AutoMath.DPosition
     Set oEnpoint = New AutoMath.DPosition
     
     oStPoint.Set -parFacetoFace / 2, 0, 0
     oEnpoint.Set parFacetoFace / 2, 0, 0
     Set oCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, parDiameter, True)
     
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCylinder
     Set oCylinder = Nothing
     
     'strap cylinders
     Dim oStrapCyl1 As Object
     Dim oTransMat  As AutoMath.DT4x4
     Dim oAxis      As AutoMath.DVector
     Dim dAngle     As Double
    
     Set oTransMat = New AutoMath.DT4x4
     Set oAxis = New AutoMath.DVector
     dAngle = PI / 12
     oStPoint.Set -parFacetoFace / 2, 0.58 * parDiameter * Cos(dAngle), -0.58 * parDiameter * Sin(dAngle)
     oEnpoint.Set parFacetoFace / 2, 0.58 * parDiameter * Cos(dAngle), -0.58 * parDiameter * Sin(dAngle)
    
     Set oStrapCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, 0.16 * parDiameter, True)
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oStrapCyl1
     Set oStrapCyl1 = Nothing
     
     Dim oStrapCyl2 As Object
     oStPoint.Set -parFacetoFace / 2, 0.58 * parDiameter * Cos(dAngle), 0.58 * parDiameter * Sin(dAngle)
     oEnpoint.Set parFacetoFace / 2, 0.58 * parDiameter * Cos(dAngle), 0.58 * parDiameter * Sin(dAngle)

     Set oStrapCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, 0.16 * parDiameter, True)
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oStrapCyl2
     Set oStrapCyl2 = Nothing
     Set oAxis = Nothing
     Set oTransMat = Nothing
     
     'Strap Bolts
     Dim oStrapBolt1 As Object
     Dim oStrapBolt2 As Object
     Set oTransMat = New AutoMath.DT4x4
     Set oAxis = New AutoMath.DVector
   
     oStPoint.Set 0.35 * parFacetoFace, 0.58 * parDiameter * Cos(dAngle), 0.25 * parDiameter
     oEnpoint.Set 0.35 * parFacetoFace, 0.58 * parDiameter * Cos(dAngle), -(parBoltEndDimension - parDiameter / 2) * Sin(PI / 8)
     Set oStrapBolt1 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, 0.08 * parDiameter, True)
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oStrapBolt1
     
     oStPoint.Set -0.35 * parFacetoFace, 0.58 * parDiameter * Cos(dAngle), 0.25 * parDiameter
     oEnpoint.Set -0.35 * parFacetoFace, 0.58 * parDiameter * Cos(dAngle), -(parBoltEndDimension - parDiameter / 2) * Sin(PI / 8)
     Set oStrapBolt2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, 0.08 * parDiameter, True)
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oStrapBolt2
     Set oStrapBolt1 = Nothing
     Set oStrapBolt2 = Nothing
     Set oTransMat = Nothing
     Set oAxis = Nothing
     Set oStPoint = Nothing
     Set oEnpoint = Nothing
      
     'Place nozzle1
     Dim oPlacePoint  As AutoMath.DPosition
     Dim oDir         As AutoMath.DVector
     Dim oNozzle1     As GSCADNozzleEntities.IJDNozzle
     
     Set oDir = New AutoMath.DVector
     Set oPlacePoint = New AutoMath.DPosition
     
     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
     
     oPlacePoint.Set -parPorttoPortOffset / 2 - cptOffset + depth, 0, 0
     oDir.Set -1, 0, 0
     Set oNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
     
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle1
     Set oDir = Nothing
     Set oPlacePoint = Nothing
     Set oNozzle1 = Nothing
     
     'place nozzle2
     Dim oNozzle2 As GSCADNozzleEntities.IJDNozzle
     Set oDir = New AutoMath.DVector
     Set oPlacePoint = New AutoMath.DPosition
     
     RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
     
     oPlacePoint.Set parPorttoPortOffset / 2 + cptOffset - depth, 0, 0
     oDir.Set 1, 0, 0
     Set oNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
     
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle2
     Set oDir = Nothing
     Set oPlacePoint = Nothing
     Set oNozzle2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
