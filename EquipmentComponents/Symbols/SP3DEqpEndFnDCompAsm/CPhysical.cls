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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: VRK
'   Creation Date:  Friday, Oct 25 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Four Outputs:
'      In which Planes, Circles,and Points are created at bottom end of the Head.
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
    Dim ObjFnDHeadbody As Object
    Dim iOutput As Double
    
    Dim parVesselDiameter As Double
    Dim parEndHeadHeight As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)       'P52
    parInsulationThickness = arrayOfInputs(3)

    m_oSymGeomHelper.OutputCollection = m_outputColl
    parEndHeadHeight = 0.19 * parVesselDiameter
    iOutput = 0
    
    'ObjFnDHeadbody
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Dim oSemiEllipsoid As IngrGeom3D.Revolution3d
    Dim oCenterPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector
    Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
    Dim dEllipseMajorAxis As Double
    Dim dMinorToMajorRatio As Double
    dEllipseMajorAxis = parVesselDiameter / 2
    dMinorToMajorRatio = parEndHeadHeight / dEllipseMajorAxis
    
    oCenterPoint.Set 0, 0, 0
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oEllipticalArc = oGeometryFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                         oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dEllipseMajorAxis, 0, 0, _
                         dMinorToMajorRatio, PI, PI / 2)
  
    oAxisVec.Set 0, 0, 1
    Set ObjFnDHeadbody = PlaceRevolution(m_outputColl, oEllipticalArc, oAxisVec, oCenterPoint, 2 * PI, True)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFnDHeadbody
    Set ObjFnDHeadbody = Nothing
    Set oCenterPoint = Nothing
    
    Dim oCircle3d As Object
    Dim oPlane As Object
    Dim oPoint3d As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory

    Set oPoint3d = oGeometryFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                   0, 0, 0)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPoint3d
    Set oPoint3d = Nothing
      
    Set oCircle3d = oGeometryFactory.Circles3d.CreateByCenterNormalRadius(m_outputColl.ResourceManager, _
                                        0, 0, 0, 0, 0, 1, dEllipseMajorAxis)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oCircle3d
    
    Set oPlane = oGeometryFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, oCircle3d)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPlane
   
    Set oCircle3d = Nothing
    Set oPlane = Nothing
    Set oGeometryFactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub



