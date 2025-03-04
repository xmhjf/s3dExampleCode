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
'   This class module has Seven Outputs:
'      In which Planes, Circles, and Points are created at both ends of the Head.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'******************************************************************************

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     
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
    Dim ObjConeHeadbody As Object
    Dim iOutput As Double
    Dim parEndHeadConeTopDiameter As Double
    Dim parVesselDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parInsulationThickness As Double

    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)                'P52
    parEndHeadConeTopDiameter = arrayOfInputs(3)        'P51
    parEndHeadConeHeight = arrayOfInputs(4)             'P50
    parInsulationThickness = arrayOfInputs(5)
    
    m_oSymGeomHelper.OutputCollection = m_outputColl
    iOutput = 0
    
    stPoint.Set 0, 0, parEndHeadConeHeight
    enPoint.Set 0, 0, 0
    
    Set ObjConeHeadbody = PlaceCone(m_outputColl, stPoint, enPoint, parEndHeadConeTopDiameter / 2, parVesselDiameter / 2, True)
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConeHeadbody
    Set ObjConeHeadbody = Nothing

    Dim oCircle3d As Object
    Dim oPlane As Object
    Dim oPoint3d As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                   0, 0, 0)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPoint3d
    Set oPoint3d = Nothing
    
    Set oPoint3d = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                    0, 0, parEndHeadConeHeight)
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

    Set oCircle3d = geomFactory.Circles3d.CreateByCenterNormalRadius(m_outputColl.ResourceManager, _
                           0, 0, parEndHeadConeHeight, 0, 0, 1, parEndHeadConeTopDiameter / 2)
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



