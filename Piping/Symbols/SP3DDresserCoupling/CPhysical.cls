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
'   Author:         KKC
'   Creation Date:  Thursday, Aug 23 2007
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The part data is obtained from http://www.dressercouplings.com/internet/pages/documentwarehouse/get.cfm?id=5667&a=1
'
'   Change History:
'   dd.mmm.yyyy           who                      change description
'   -----------         -----                      ------------------
'   10.SEP.2007           KKC            CR-122746  Create realistic Dresser Coupling Symbol
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBody As Object
   
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    Dim parMajorBodyDiameter As Double
    Dim parNoofBolts As Long
    Dim ParPortToPort As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parMajorBodyDiameter = arrayOfInputs(3)
    parNoofBolts = arrayOfInputs(4)
    ParPortToPort = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 1(Body)
     Dim oStPoint   As AutoMath.DPosition
     Set oStPoint = New AutoMath.DPosition
     Dim oEnPoint   As New AutoMath.DPosition
     Set oEnPoint = New AutoMath.DPosition
     oStPoint.Set -parFacetoFace / 2, 0, 0
     oEnPoint.Set parFacetoFace / 2, 0, 0
     
     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
     Set ObjBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.01 * pipeDiam, False)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
'Insert Your Code for Output 2(Right Cylinder)
    Dim dCylinderThick As Double
    dCylinderThick = 0.0254   '1 Inch
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnPoint.Set -parFacetoFace / 2 + dCylinderThick, 0, 0
    Dim ObjRightCylinder As Object
    Set ObjRightCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parMajorBodyDiameter, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCylinder
    Set ObjRightCylinder = Nothing
    
'Insert Your Code for Output 3(Left Cylinder)
    oStPoint.Set parFacetoFace / 2, 0, 0
    oEnPoint.Set parFacetoFace / 2 - dCylinderThick, 0, 0
    Dim ObjLeftCylinder As Object
    Set ObjLeftCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parMajorBodyDiameter, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCylinder
    Set ObjLeftCylinder = Nothing
    
'Insert Your Code for Output 4 (Bolts)
    Dim oTransMatrix  As IJDT4x4
    Set oTransMatrix = New AutoMath.DT4x4
    Dim i As Long
    Dim ObjBolt As Object
    Dim oDir As AutoMath.DVector
    Dim dAngle As Double
    For i = 1 To parNoofBolts
        oStPoint.Set -parFacetoFace / 2 - dCylinderThick, (parMajorBodyDiameter + pipeDiam) / 4, 0
        oEnPoint.Set parFacetoFace / 2 + dCylinderThick, (parMajorBodyDiameter + pipeDiam) / 4, 0
        Set oDir = New AutoMath.DVector
        Set ObjBolt = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.0125, True)
        oTransMatrix.LoadIdentity
        oDir.Set 1, 0, 0
        dAngle = i * (2 * PI / parNoofBolts)
        oTransMatrix.Rotate dAngle, oDir
        ObjBolt.Transform oTransMatrix
        m_OutputColl.AddOutput "ObjBolt_", ObjBolt
        Set ObjBolt = Nothing
    Next i
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
  
    Set oPlacePoint = New AutoMath.DPosition
    oPlacePoint.Set -ParPortToPort / 2, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set ParPortToPort / 2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
