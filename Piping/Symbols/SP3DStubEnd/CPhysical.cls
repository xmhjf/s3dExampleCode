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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          RRK
'   Creation Date:   Wednusday, Mar 26 2008
'   Description:

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'  26.Mar.2008     RRK      CR-83045:Created new symbol for stub end
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI           As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim flangeDiam As Double
    Dim cptOffset As Double
    Dim depth As Double

    Dim iOutput As Double

    Dim parStubLength As Double
    Dim parLapThickness As Double
    Dim parLapDiameter As Double
    Dim parLapRadius As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStubLength = arrayOfInputs(2)
    parLapThickness = arrayOfInputs(3)
    parLapDiameter = arrayOfInputs(4)
    parLapRadius = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0

    'Create Output 1 (Cylinder for Hub)
    Dim objHub As Object

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition

    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    oStPoint.Set 0, 0, 0
    oEnPoint.Set parLapThickness, 0, 0

    Set objHub = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parLapDiameter, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHub
    Set objHub = Nothing

    ' Create Output 2(Stub End Cylinder)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    Dim objStubEnd As Object

    oStPoint.Set parLapThickness, 0, 0
    oEnPoint.Set parStubLength, 0, 0

    Set objStubEnd = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objStubEnd
    Set objStubEnd = Nothing

    ' Create Output 3(Curved Portion)
    'Creating the curved portion only when lap radius is provided
    If Not CmpDblEqual(parLapRadius, 0) Then
        Dim objCurvePortion As Object
        Dim oArc As IngrGeom3D.Arc3d
        Set oArc = New IngrGeom3D.Arc3d

        Dim oAxis As AutoMath.DVector
        Set oAxis = New AutoMath.DVector

        Dim oCentPos As AutoMath.DPosition
        Set oCentPos = New AutoMath.DPosition

        'Creating arc which will be revolved about X-axis
        oAxis.Set 0, 0, 1
        oCentPos.Set parLapThickness + parLapRadius, (pipeDiam / 2) + parLapRadius, 0

        oStPoint.Set parLapThickness, (pipeDiam / 2) + parLapRadius, 0
        oEnPoint.Set parLapThickness + parLapRadius, pipeDiam / 2, 0

        Set oArc = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCentPos, oAxis)

        'Revolving the arc about X-axis
        oAxis.Set 1, 0, 0
        oCentPos.Set 0, 0, 0
        Set objCurvePortion = PlaceRevolution(m_OutputColl, oArc, oAxis, oCentPos, 2 * PI, True)

        ' Set the output
        m_OutputColl.AddOutput "CurvePortion", objCurvePortion
        Set objCurvePortion = Nothing
        Set oAxis = Nothing
        Set oCentPos = Nothing
    End If

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    ' Place Nozzle 1
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As New AutoMath.DPosition

    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    oPlacePoint.Set 0, 0, 0

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    ' Place Nozzle 2
    Dim dPort2Location As Double

    oPlacePoint.Set parStubLength + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
