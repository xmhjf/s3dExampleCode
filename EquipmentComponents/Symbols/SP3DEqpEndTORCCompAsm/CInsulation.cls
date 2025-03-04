VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author: VRK
'   Creation Date:  Wednesday, Mar 21 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Two Outputs:
'           Insulated TORI Conical head
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages
    
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjInsTORCHeadbody As Object
    Dim iOutput     As Double
    Dim parVesselDiameter As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parInsulationThickness As Double

    
    Dim CenterPos As IJDPosition
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim e As Double
    Dim Offset As Double
    Dim dblConeLargeDiameter As Double
    Dim dblConeRadius As Double
    Dim dblKnuckleAngle As Double
    Dim dblConeHeight As Double
    Dim dblHcenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parEndHeadConeTopDiameter = arrayOfInputs(3)    'P52
    parEndHeadConeHeight = arrayOfInputs(4)         'P50
    parEndHeadKnuckleRadius = arrayOfInputs(5)      'P51
    parInsulationThickness = arrayOfInputs(6)

    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    Set CenterPos = New DPosition
    CenterPos.Set 0, 0, 0
    iOutput = 0
    
    parVesselDiameter = parVesselDiameter + 2 * parInsulationThickness
    parEndHeadConeTopDiameter = parEndHeadConeTopDiameter + 2 * parInsulationThickness
    parEndHeadKnuckleRadius = parEndHeadKnuckleRadius + parInsulationThickness
    
    dblConeRadius = 0.5 * parVesselDiameter
    a = dblConeRadius - (0.5 * parEndHeadConeTopDiameter) - parEndHeadKnuckleRadius
    b = parEndHeadKnuckleRadius * parEndHeadConeHeight
    c = a * a * a * a + (a * a) * (parEndHeadConeHeight * parEndHeadConeHeight) - (a * a) * (parEndHeadKnuckleRadius * parEndHeadKnuckleRadius)
    d = parEndHeadConeHeight * parEndHeadConeHeight + a * a
    e = (b + Abs(Sqr(c))) / d
    
    'Compute knuckle arc angle value
    dblKnuckleAngle = ASin(e)
    dblHcenter = parEndHeadKnuckleRadius * (Sin(dblKnuckleAngle))
    dblConeHeight = parEndHeadConeHeight - dblHcenter
    
    Offset = parEndHeadKnuckleRadius * (1 - Cos(dblKnuckleAngle))
    dblConeLargeDiameter = 2 * (parVesselDiameter / 2 - parEndHeadKnuckleRadius + parEndHeadKnuckleRadius * Cos(dblKnuckleAngle))

   
    Dim objCurve As IngrGeom3D.Arc3d
    Dim objRevolution  As IngrGeom3D.Revolution3d
    Dim geomFactory    As IngrGeom3D.GeometryFactory

    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objCurve = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                                               parVesselDiameter / 2 - parEndHeadKnuckleRadius, 0, 0, parVesselDiameter / 2, 0, 0, dblConeLargeDiameter / 2, 0, -dblHcenter)
                                                         
    Set objRevolution = geomFactory.Revolutions3d.CreateByCurve( _
                                                    m_outputColl.ResourceManager, _
                                                    objCurve, _
                                                    0, 0, 1, _
                                                    0, 0, 0, _
                                                    2 * PI, True)
                                                    
                                                   
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objRevolution
    Set objCurve = Nothing
    Set objRevolution = Nothing

    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    stPoint.Set 0, 0, -parEndHeadConeHeight
    enPoint.Set 0, 0, -parEndHeadConeHeight + dblConeHeight

    Set ObjInsTORCHeadbody = PlaceCone(m_outputColl, stPoint, enPoint, parEndHeadConeTopDiameter / 2, dblConeLargeDiameter / 2, True)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTORCHeadbody
    Set ObjInsTORCHeadbody = Nothing
    Set geomFactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub



   
   
    





