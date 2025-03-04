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
'   Creation Date:  Monday, Jan 22 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     -----                ------------------
'   09.Jul.2003     SymbolTeam(India)    Copyright Information, Header  is added.
'   08.SEP.2006     KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'   19.Sep.2007     VRG                  CR-127655 Enhanced the existing check valve to support sanitary vertical and horizontal check valves
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:"    'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput As Integer
    Dim oInsulatedBody As Object

    Dim parFacetoFace As Double
    Dim parBodyDiameter As Double
    Dim parPortVOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    'parHandwheelAngle = arrayOfInputs(4) (optional parameter)
    'parBodyDiameter = arrayOfInputs(5) (optional parameter)
    'parPortVOffset = arrayOfInputs(6)  (optional parameter)
    If UBound(arrayOfInputs) >= 4 Then
        parHandwheelAngle = arrayOfInputs(4)
    End If
    
    iOutput = 0
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    If (lPartDataBasis = 455 Or lPartDataBasis = 456) Then
        parBodyDiameter = arrayOfInputs(5)
        parPortVOffset = arrayOfInputs(6)
        If (lPartDataBasis = 456) Then parPortVOffset = 0
        'Insert your code for InsulatedBody(Output 4)
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition

        oStPoint.Set -parFacetoFace / 2 - parInsulationThickness, -0.1 * parBodyDiameter, 0
        oEnPoint.Set parFacetoFace / 2 + parInsulationThickness, -0.1 * parBodyDiameter, 0
        Set oInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBodyDiameter + 2 * parInsulationThickness, True)

    ElseIf (lPartDataBasis <= 1) Then
        'Insert your code for output 2(Check Valve Insulated Body)
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        oStPoint.Set -parFacetoFace / 2, 0, 0
        oEnPoint.Set parFacetoFace / 2, 0, 0

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        Dim parFlangeDiameter As Double
        parFlangeDiameter = BodyOD(flangeDiam, pipeDiam)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        Dim parFlangeDiameter2 As Double
        parFlangeDiameter2 = BodyOD(flangeDiam, pipeDiam)

        Dim parInsulationDiameter As Double
        parInsulationDiameter = InsulationDiameter(parFlangeDiameter, parFlangeDiameter2, parInsulationThickness)

        Set oInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
    End If
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsulatedBody

    Set oInsulatedBody = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
