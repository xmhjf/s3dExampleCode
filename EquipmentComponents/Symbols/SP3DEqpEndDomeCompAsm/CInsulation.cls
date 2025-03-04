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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author: VRK
'   Creation Date:  Friday, Oct 25 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           Insulated Dome head
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     ---                  ------------------
'   11.04.2008      VRK                 TR-139002:When user modifies Domehead radii of E210SimVertCylEqpSkirt, shape is incorrect
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
    Dim ObjDomeInsHeadbody As Object
    
    Dim iOutput As Double
    Dim parVesselDiameter As Double
    Dim dEndHeadHeight As Double
    Dim parEndHeadDomeradius As Double
    Dim parInsulationThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)            'P52
    parEndHeadDomeradius = arrayOfInputs(3)         'P50
    parInsulationThickness = arrayOfInputs(4)
    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    iOutput = 0
    'Insert your code for Dome
    Dim oGeomfactory    As IngrGeom3D.GeometryFactory
    Set oGeomfactory = New IngrGeom3D.GeometryFactory
    Dim dVesselInsDia As Double
    Dim dEndHeadDomeRadius As Double
    dVesselInsDia = parVesselDiameter + 2 * parInsulationThickness
    dEndHeadDomeRadius = parEndHeadDomeradius + parInsulationThickness
    dEndHeadHeight = Sqr((dEndHeadDomeRadius * dEndHeadDomeRadius) - (dVesselInsDia * dVesselInsDia / 4))

    Dim objCurve As IngrGeom3D.Arc3d
    Dim objRevolution  As IngrGeom3D.Revolution3d
    Set objCurve = oGeomfactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                     0, 0, -dEndHeadHeight, _
                                                     -dVesselInsDia / 2, 0, 0, _
                                                       0, 0, dEndHeadDomeRadius - dEndHeadHeight)
                                                      
    Set objRevolution = oGeomfactory.Revolutions3d.CreateByCurve( _
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
    Set oGeomfactory = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
