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
'   CSimplePhysical.cls
'   Author:        MA
'   Creation Date: Friday 17, Aug 2007
'   Description:
'    This symbol is created for  CR-123053
'    Source: Rosebox Strainer.doc
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
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
    
    Dim parPipingSpecialtyLength As Double
    Dim parPipingSpecialtyHeight As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPipingSpecialtyLength = arrayOfInputs(2)
    parPipingSpecialtyHeight = arrayOfInputs(3)

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    iOutput = 0
        
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        
 ' Insert your code for output 1(Cylinder)
    Dim ObjCylinder As Object
    'Assumption: Cylinder length is assumed to be 0.2 times parPipingSpecialtyLength
    
    oStPoint.Set -0.2 * parPipingSpecialtyLength, 0, 0
    oEnPoint.Set 0, 0, 0
    
    Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.2 * pipeDiam, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
        
 ' Insert your code for output 2(Box1)
    Dim ObjBox1 As Object
    
    oStPoint.Set 0, -0.5 * parPipingSpecialtyHeight, 0.5 * parPipingSpecialtyHeight
    oEnPoint.Set parPipingSpecialtyLength, 0.5 * parPipingSpecialtyHeight, -0.5 * parPipingSpecialtyHeight
    
    Set ObjBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
 ' Insert your code for output 3(Box2)
    Dim ObjBox2 As Object
    
    oStPoint.Set -0.2 * parPipingSpecialtyLength, -0.05 * parPipingSpecialtyHeight, _
                               0.6 * pipeDiam + 0.2 * parPipingSpecialtyHeight
    oEnPoint.Set 0, 0.05 * parPipingSpecialtyHeight, 0.6 * pipeDiam
    
    Set ObjBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing
    
 ' Insert your code for output 4(Box3)
    Dim ObjBox3 As Object
    
    oStPoint.Set -0.2 * parPipingSpecialtyLength, -0.05 * parPipingSpecialtyHeight, -0.6 * pipeDiam
    oEnPoint.Set 0, 0.05 * parPipingSpecialtyHeight, _
                         -0.6 * pipeDiam - 0.2 * parPipingSpecialtyHeight
    
    Set ObjBox3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing
    
 ' Insert your code for output 5(Box4)
    Dim ObjBox4 As Object
    
    oStPoint.Set 0, -0.05 * parPipingSpecialtyHeight, 0.5 * parPipingSpecialtyHeight
    oEnPoint.Set parPipingSpecialtyLength, 0.05 * parPipingSpecialtyHeight, 0.6 * parPipingSpecialtyHeight
    
    Set ObjBox4 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox4
    Set ObjBox4 = Nothing
    
 ' Insert your code for output 6(Box5)
    Dim ObjBox5 As Object
    
    oStPoint.Set 0, -0.05 * parPipingSpecialtyHeight, -0.5 * parPipingSpecialtyHeight
    oEnPoint.Set parPipingSpecialtyLength, 0.05 * parPipingSpecialtyHeight, -0.6 * parPipingSpecialtyHeight
    
    Set ObjBox5 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox5
    Set ObjBox5 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

