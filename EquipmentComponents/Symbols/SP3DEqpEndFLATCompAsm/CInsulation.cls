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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author: VRK
'   Creation Date:  Friday, Oct 25 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Three Outputs:
'      In which Insulated Planes, Circles, and Points are created
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
    Dim parFrontEndHeadDia As Double
    Dim iOutput     As Double
    Dim parInsulationThickness As Double
    Dim parVesselDiameter As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)        'P52
    parInsulationThickness = arrayOfInputs(3)

    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    iOutput = 0
    Dim oCircle3d As Object
    Dim oPlane As Object
    Dim oPoint3d As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory

    Set oPoint3d = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                   0, 0, 0)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oPoint3d
    Set oPoint3d = Nothing
    
    parVesselDiameter = parVesselDiameter + 2 * parInsulationThickness
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





