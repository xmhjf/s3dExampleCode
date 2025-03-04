VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'  "SimplePhysical:".cls
'  Author:        Botao.Huo
'  Creation Date:  November 2007
'  Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      TODO - fill in header description information
'
'  Change History:
'  dd.mmm.yyyy    who    change description
'  -----------    ---    ------------------
'  22.08.2007     Huo Botao(China)

Option Explicit
Private Const MODULE = "CSimplePhysical"    'Used for error messages

Private Const Type_A = 5
Private Const E_FAIL = &H80004005

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize"
    On Error GoTo Errx

    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo Errx

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    iOutput = 0

    ' Declare variables for Inputs and Outputs

    'Inputs
    Dim parPlateThickness As Double
    Dim parPlateCorner01X As Double
    Dim parPlateCorner01Y As Double
    Dim parPlateCorner02X As Double
    Dim parPlateCorner02Y As Double
    Dim parPlateCorner03X As Double
    Dim parPlateCorner03Y As Double
    Dim parPlateCorner04X As Double
    Dim parPlateCorner04Y As Double
    Dim parNumberOfBolts As Long
    Dim parBoltDiameter As Double
    Dim parBoltStandout As Double
    Dim parBolt01X As Double
    Dim parBolt01Y As Double
    Dim parBolt02X As Double
    Dim parBolt02Y As Double
    Dim parBolt03X As Double
    Dim parBolt03Y As Double
    Dim parBolt04X As Double
    Dim parBolt04Y As Double

    'Outputs
    Dim objWearPlate As Object
    Dim objBolt1 As Object

    'Insert your code for inputs
    Set oPartFclt = arrayOfInputs(1)
    parPlateThickness = arrayOfInputs(2)
    parPlateCorner01X = arrayOfInputs(3)
    parPlateCorner01Y = arrayOfInputs(4)
    parPlateCorner02X = arrayOfInputs(5)
    parPlateCorner02Y = arrayOfInputs(6)
    parPlateCorner03X = arrayOfInputs(7)
    parPlateCorner03Y = arrayOfInputs(8)
    parPlateCorner04X = arrayOfInputs(9)
    parPlateCorner04Y = arrayOfInputs(10)
    parNumberOfBolts = arrayOfInputs(11)
    parBoltDiameter = arrayOfInputs(12)
    parBoltStandout = arrayOfInputs(13)
    parBolt01X = arrayOfInputs(14)
    parBolt01Y = arrayOfInputs(15)
    parBolt02X = arrayOfInputs(16)
    parBolt02Y = arrayOfInputs(17)
    parBolt03X = arrayOfInputs(18)
    parBolt03Y = arrayOfInputs(19)
    parBolt04X = arrayOfInputs(20)
    parBolt04Y = arrayOfInputs(21)


    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim iBoltNum As Integer

    'Insert your code for output1 (Wear Plate)

    Dim arrPoints(11) As Double
    Dim BoxX(8) As Double
    Dim BoxY(8) As Double
    Dim BoxZ(8) As Double

    BoxX(1) = parPlateCorner01X         'Corner 01
    BoxY(1) = -parPlateCorner01Y
    BoxZ(1) = 0
    BoxX(2) = parPlateCorner02X         'Corner 02
    BoxY(2) = parPlateCorner02Y
    BoxZ(2) = 0
    BoxX(3) = -parPlateCorner03X        'Corner 03
    BoxY(3) = parPlateCorner03Y
    BoxZ(3) = 0
    BoxX(4) = -parPlateCorner04X        'Corner 04
    BoxY(4) = -parPlateCorner04Y
    BoxZ(4) = 0
    BoxX(5) = parPlateCorner01X         'Corner 05
    BoxY(5) = -parPlateCorner01Y
    BoxZ(5) = parPlateThickness
    BoxX(6) = parPlateCorner02X         'Corner 06
    BoxY(6) = parPlateCorner02Y
    BoxZ(6) = parPlateThickness
    BoxX(7) = -parPlateCorner03X        'Corner 07
    BoxY(7) = parPlateCorner03Y
    BoxZ(7) = parPlateThickness
    BoxX(8) = -parPlateCorner04X        'Corner 08
    BoxY(8) = -parPlateCorner04Y
    BoxZ(8) = parPlateThickness

    'Create 1st surface.
    arrPoints(0) = BoxX(1)
    arrPoints(1) = BoxY(1)
    arrPoints(2) = BoxZ(1)
    arrPoints(3) = BoxX(4)
    arrPoints(4) = BoxY(4)
    arrPoints(5) = BoxZ(4)
    arrPoints(6) = BoxX(3)
    arrPoints(7) = BoxY(3)
    arrPoints(8) = BoxZ(3)
    arrPoints(9) = BoxX(2)
    arrPoints(10) = BoxY(2)
    arrPoints(11) = BoxZ(2)

    iOutput = iOutput + 1
    Set objWearPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWearPlate

    'Create 2nd surface.
    arrPoints(0) = BoxX(5)
    arrPoints(1) = BoxY(5)
    arrPoints(2) = BoxZ(5)
    arrPoints(3) = BoxX(6)
    arrPoints(4) = BoxY(6)
    arrPoints(5) = BoxZ(6)
    arrPoints(6) = BoxX(7)
    arrPoints(7) = BoxY(7)
    arrPoints(8) = BoxZ(7)
    arrPoints(9) = BoxX(8)
    arrPoints(10) = BoxY(8)
    arrPoints(11) = BoxZ(8)

    iOutput = iOutput + 1
    Set objWearPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWearPlate

    'Create 3rd surface.
    arrPoints(0) = BoxX(1)
    arrPoints(1) = BoxY(1)
    arrPoints(2) = BoxZ(1)
    arrPoints(3) = BoxX(2)
    arrPoints(4) = BoxY(2)
    arrPoints(5) = BoxZ(2)
    arrPoints(6) = BoxX(6)
    arrPoints(7) = BoxY(6)
    arrPoints(8) = BoxZ(6)
    arrPoints(9) = BoxX(5)
    arrPoints(10) = BoxY(5)
    arrPoints(11) = BoxZ(5)

    iOutput = iOutput + 1
    Set objWearPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWearPlate

    'Create 4th surface.
    arrPoints(0) = BoxX(1)
    arrPoints(1) = BoxY(1)
    arrPoints(2) = BoxZ(1)
    arrPoints(3) = BoxX(5)
    arrPoints(4) = BoxY(5)
    arrPoints(5) = BoxZ(5)
    arrPoints(6) = BoxX(8)
    arrPoints(7) = BoxY(8)
    arrPoints(8) = BoxZ(8)
    arrPoints(9) = BoxX(4)
    arrPoints(10) = BoxY(4)
    arrPoints(11) = BoxZ(4)

    iOutput = iOutput + 1
    Set objWearPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWearPlate

    'Create 5th surface.
    arrPoints(0) = BoxX(2)
    arrPoints(1) = BoxY(2)
    arrPoints(2) = BoxZ(2)
    arrPoints(3) = BoxX(3)
    arrPoints(4) = BoxY(3)
    arrPoints(5) = BoxZ(3)
    arrPoints(6) = BoxX(7)
    arrPoints(7) = BoxY(7)
    arrPoints(8) = BoxZ(7)
    arrPoints(9) = BoxX(6)
    arrPoints(10) = BoxY(6)
    arrPoints(11) = BoxZ(6)

    iOutput = iOutput + 1
    Set objWearPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWearPlate

    'Create 6th surface.
    arrPoints(0) = BoxX(3)
    arrPoints(1) = BoxY(3)
    arrPoints(2) = BoxZ(3)
    arrPoints(3) = BoxX(4)
    arrPoints(4) = BoxY(4)
    arrPoints(5) = BoxZ(4)
    arrPoints(6) = BoxX(8)
    arrPoints(7) = BoxY(8)
    arrPoints(8) = BoxZ(8)
    arrPoints(9) = BoxX(7)
    arrPoints(10) = BoxY(7)
    arrPoints(11) = BoxZ(7)

    iOutput = iOutput + 1
    Set objWearPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objWearPlate

    Dim stPos As IJDPosition
    Dim enPos As IJDPosition
    Set stPos = New DPosition
    Set enPos = New DPosition

    parNumberOfBolts = Int(parNumberOfBolts)
    If parNumberOfBolts < 1 Then parNumberOfBolts = 1
    If parNumberOfBolts > 4 Then parNumberOfBolts = 4

    'Insert your code for output2 (Bolt1)
    stPos.Set parBolt01X, parBolt01Y, 0
    enPos.Set parBolt01X, parBolt01Y, -parBoltStandout

    iOutput = iOutput + 1
    Set objBolt1 = PlaceCylinder(m_OutputColl, stPos, enPos, parBoltDiameter, True)
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBolt1

    'Insert your code for output2 (Bolt2)
    If parNumberOfBolts >= 2 Then
        Dim objBolt2 As Object
        stPos.Set parBolt02X, parBolt02Y, 0
        enPos.Set parBolt02X, parBolt02Y, -parBoltStandout

        iOutput = iOutput + 1
        Set objBolt2 = PlaceCylinder(m_OutputColl, stPos, enPos, parBoltDiameter, True)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBolt2
        Set objBolt2 = Nothing
    End If

    'Insert your code for output3 (Bolt3)
    If parNumberOfBolts >= 3 Then
        Dim objBolt3 As Object
        stPos.Set parBolt03X, parBolt03Y, 0
        enPos.Set parBolt03X, parBolt03Y, -parBoltStandout

        iOutput = iOutput + 1
        Set objBolt3 = PlaceCylinder(m_OutputColl, stPos, enPos, parBoltDiameter, True)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBolt3
        Set objBolt3 = Nothing
    End If

    'Insert your code for output4 (Bolt4)
    If parNumberOfBolts >= 4 Then
        Dim objBolt4 As Object
        stPos.Set parBolt04X, parBolt04Y, 0
        enPos.Set parBolt04X, parBolt04Y, -parBoltStandout

        iOutput = iOutput + 1
        Set objBolt4 = PlaceCylinder(m_OutputColl, stPos, enPos, parBoltDiameter, True)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBolt4
        Set objBolt4 = Nothing
    End If

    ''Release BO 's
    Set oGeomFactory = Nothing
    Set objWearPlate = Nothing
    Set objBolt1 = Nothing
    Set stPos = Nothing
    Set enPos = Nothing

    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
              Err.HelpFile, Err.HelpContext
End Sub

