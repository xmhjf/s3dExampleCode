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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Saturday, Jul 21 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy           who                         change description
'   -----------         -----                         ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006      KKC                     DI-95670: Replace names with initials in
'                                           all revision history sheets and symbols
'  21.Feb.2007      RRK                     TR-113129 The Eccentric Reducer Symbol Insulation Aspect
'                                           should define the Insulation to follow the physical body
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1     As Double
    Dim depth1           As Double
    Dim pipeInsideDiam1 As Double
    Dim pipeInsideDiam2 As Double
    
    Dim oCenterPos       As Object
    Set oCenterPos = New AutoMath.DPosition

    oCenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parEccentricOffsetBasis As Long

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parEccentricOffsetBasis = arrayOfInputs(4)
   
    iOutput = 0

 ' Insert your code for output 2(Insulated Body)
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim CptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double

  ' Create Circle1

    RetrieveParametersWithInsidePipeDiameter 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1, pipeInsideDiam1
    
    RetrieveParametersWithInsidePipeDiameter 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, CptOffset2, depth2, pipeInsideDiam2
    
    Dim dOffsetBetweemCentreLines As Double
    
    If parEccentricOffsetBasis = 2 Then 'based on inside diameter
        dOffsetBetweemCentreLines = (pipeInsideDiam1 - pipeInsideDiam2) / 2
    Else 'considering parEccentricOffsetBasis = 1 case based on outside diameter
        dOffsetBetweemCentreLines = (pipeDiam1 - pipeDiam2) / 2
    End If
    
    
    Dim dCirNormalX As Double
    Dim dCirNormalY As Double
    Dim dCirNormalZ As Double
    
    dCirNormalX = 1
    dCirNormalY = 0
    dCirNormalZ = 0
    
    Dim objCircle1 As IngrGeom3D.Circle3d
    Dim dInsulPort1Diam As Double
    Dim dInsulPort2Diam As Double
    
    dInsulPort1Diam = BodyOD(pipeDiam1, flangeDiam1) + 2 * parInsulationThickness
    dInsulPort2Diam = BodyOD(pipeDiam2, flangeDiam2) + 2 * parInsulationThickness

    Dim dCir1CenterX As Double
    Dim dCir1CenterY As Double
    Dim dCir1CenterZ As Double
    
    dCir1CenterX = -parFacetoFace / 2
    dCir1CenterY = dOffsetBetweemCentreLines / 2
    dCir1CenterZ = 0
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, dCir1CenterX, _
                                                dCir1CenterY, dCir1CenterZ, dCirNormalX, _
                                                dCirNormalY, dCirNormalZ, dInsulPort1Diam / 2)

'Create Circle2
    Dim objCircle2 As IngrGeom3D.Circle3d
    Dim dCir2CenterX As Double
    Dim dCir2CenterY As Double
    Dim dCir2CenterZ As Double
    
    dCir2CenterX = parFacetoFace / 2
    dCir2CenterY = -dOffsetBetweemCentreLines / 2
    dCir2CenterZ = 0
        
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, dCir2CenterX, _
                                                dCir2CenterY, dCir2CenterZ, dCirNormalX, _
                                                dCirNormalY, dCirNormalZ, dInsulPort2Diam / 2)
    
    
    'Create Ruled Surface for body
    Dim ObjInsulatedBody As IngrGeom3D.RuledSurface3d
    
    Set ObjInsulatedBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, objCircle2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
    Set oCenterPos = Nothing
    Set geomFactory = Nothing
    Set objCircle1 = Nothing
    Set objCircle2 = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


