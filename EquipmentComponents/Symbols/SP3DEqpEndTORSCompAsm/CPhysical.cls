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
'   Author: VRK
'   Creation Date:  Wednesday, Mar 21 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Five Outputs:
'      In which Planes, Circles, and Points are created at bottom end of the Head.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     PI = Atn(1) * 4
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim parVesselDiameter As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parInsulationThickness As Double
    Dim CenterPos As IJDPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parEndHeadSphericalRadius = arrayOfInputs(3)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(4)          'P51
    parInsulationThickness = arrayOfInputs(5)
    
    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    Set CenterPos = New DPosition
    CenterPos.Set 0, 0, 0
    iOutput = 0

    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim dbldishhalfAngle As Double
    Dim dblConeHeight As Double
    Dim knuckleangle As Double

    a = (0.5 * parVesselDiameter) - parEndHeadKnuckleRadius
    b = parEndHeadSphericalRadius - parEndHeadKnuckleRadius
    c = Abs(Sqr(b * b - a * a))
    
    'Compute dish Half angle value
    dbldishhalfAngle = ASin(a / b)
    knuckleangle = PI / 2 - dbldishhalfAngle
    
    Dim objCurve  As IngrGeom3D.Arc3d
    Dim objRevolution  As IngrGeom3D.Revolution3d
    Dim geomFactory  As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    'Create Tori Spherical head by creating two revolutions
    Set objCurve = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                                            a, 0, 0, parVesselDiameter / 2, 0, 0, parEndHeadKnuckleRadius * Cos(knuckleangle), 0, -parEndHeadKnuckleRadius * Sin(knuckleangle))

    
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCurve, _
                                                     0, 0, 1, _
                                                     0, 0, 0, _
                                                     2 * PI, False)
                                                    
                                                   
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objCurve = Nothing
    Set objRevolution = Nothing

    Set objCurve = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                                             0, 0, c, a + parEndHeadKnuckleRadius * Cos(knuckleangle), 0, -parEndHeadKnuckleRadius * Sin(knuckleangle), 0, 0, -parEndHeadSphericalRadius + c)
    
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCurve, _
                                                    0, 0, 1, _
                                                    0, 0, c, _
                                                     2 * PI, False)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objCurve = Nothing
    Set objRevolution = Nothing

    Dim oCircle3d As Object
    Dim oPlane As Object
    Dim oPoint3d As Object

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                   0, 0, 0)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPoint3d
    Set oPoint3d = Nothing

    Set oCircle3d = geomFactory.Circles3d.CreateByCenterNormalRadius(m_outputColl.ResourceManager, _
                                        0, 0, 0, 0, 0, 1, parVesselDiameter / 2)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oCircle3d

    Set oPlane = geomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, oCircle3d)

    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPlane

    Set oCircle3d = Nothing
    Set oPlane = Nothing
    Set geomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub






