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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Friday 16, Jul 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol is Tri-Clamp 2FMP-14 90 degree symbol, taken from page no. 2 of TRICLAMP.pdf catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Dim m_oGeomHelper As IJSymbolGeometryHelper
Private PI       As Double

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
      PI = 4 * Atn(1)
     Set m_oGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub


Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
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
    Dim objInsulatedBody As Object

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

    m_oGeomHelper.OutputCollection = m_OutputColl

' Insert your code for output 1(Insulated Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim dElbowRadius As Double
    Dim dInsulationDiameter As Double
    Dim objCircleI   As IngrGeom3D.Circle3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    dElbowRadius = parFace2toCenter - flangeThick

    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set objCircleI = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -dElbowRadius, 0, 0, _
                        1, 0, 0, _
                        dInsulationDiameter / 2)

    Set oGeomFactory = Nothing
'   Revolve the circle for 90 degree about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector   'rotation vector for rotation
    CP.Set -dElbowRadius, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set objInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, PI / 2, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
    Set objInsulatedBody = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircleI
    Set objCircleI = Nothing
    sp3dElem1.Remove
    Set sp3dElem1 = Nothing

' Insert your code for output 2(Insulated Port1)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set stPoint.x + flangeThick + parInsulationThickness, stPoint.y, stPoint.z

    If flangeDiam > pipeDiam Then
        dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If

'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
                         dInsulationDiameter

' Insert your code for output 5(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    stPoint.Set 0, parFace2toCenter, 0
    enPoint.Set stPoint.x, stPoint.y - flangeThick - parInsulationThickness, stPoint.z

    If flangeDiam > pipeDiam Then
        dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
                         dInsulationDiameter
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
