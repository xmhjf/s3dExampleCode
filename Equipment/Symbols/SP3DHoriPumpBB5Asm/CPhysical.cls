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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Tuesday, April 10 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Twenty Five Outputs:
'      In which Default Surface is created at the bottom of the Base Plate.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   05.10.2007      RUK     CP 128473 Replaced the Inputs "MotorDiameter", "PumpTopSupportWidth"
'                           with "MotorHeight", "MotorWidth". Created the Pump and Motor by using these parameters
'   05.10.2007      RUK     CP 128584 Added new attributes JunBoxHeight, JunBoxWidth, JunBoxLength
'                               and used to create the Junction Box of Motor.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double
Private Const dInch = 15 * 2.54 / 100

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parMotorHeight As Double
    Dim parCenterToNozzFace As Double
    Dim parDischargeToSuction As Double
    Dim parImpShaftFrontEnd As Double
    Dim parImpShaftRearEnd As Double
    Dim parDistBetShaftEnds As Double
    Dim parMotorShaftToEnd As Double
    Dim parMotorWidth As Double
    Dim parDismantlingLength As Double
    Dim parBasePlateBottom As Double
    Dim parCenToFouPortStart As Double
    Dim parBasePlateWidth As Double
    Dim parBasePlateLength As Double
    Dim parBasePlateHeight As Double
    Dim parJunBoxHeight As Double
    Dim parJunBoxWidth As Double
    Dim parJunBoxLength As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parMotorHeight = arrayOfInputs(2)               'D
    parCenterToNozzFace = arrayOfInputs(3)          'NH
    parDischargeToSuction = arrayOfInputs(4)        'X
    parImpShaftFrontEnd = arrayOfInputs(5)          'Y
    parImpShaftRearEnd = arrayOfInputs(6)           'W
    parDistBetShaftEnds = arrayOfInputs(7)          'DBSE
    parMotorShaftToEnd = arrayOfInputs(8)           'CMAX
    parMotorWidth = arrayOfInputs(9)                'PMAX
    parDismantlingLength = arrayOfInputs(10)        'EX
    parBasePlateBottom = arrayOfInputs(11)          'HD
    parCenToFouPortStart = arrayOfInputs(12)        'HR
    parBasePlateWidth = arrayOfInputs(13)           'HA
    parBasePlateLength = arrayOfInputs(14)          'HB
    parBasePlateHeight = arrayOfInputs(15)          'BD
    parJunBoxHeight = arrayOfInputs(16)             'JH
    parJunBoxWidth = arrayOfInputs(17)              'JW
    parJunBoxLength = arrayOfInputs(18)             'JL
    parXBoltHole = arrayOfInputs(19)                'BX
    parYBoltHole = arrayOfInputs(20)                'HF
    
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim dImpellerDia As Double
    dImpellerDia = 0.7 * parMotorWidth
    
    'Create the Cylinder 1 at Left most End (Output 1)
    oStPoint.Set -parImpShaftFrontEnd, 0, 0
    oEnPoint.Set oStPoint.x + 0.25 * parImpShaftFrontEnd, 0, 0
    'Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.35 * dImpellerDia
    
    'Create the Cylinder 2 After the Cylinder 1 (Output 2)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.1 * parImpShaftFrontEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.5 * dImpellerDia
    
    'Create the Flange at Front End Of the Pump (Output 3)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oEnPoint.x + 0.05 * parImpShaftFrontEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.8 * dImpellerDia
    
    'Create the Impeller part 1(Cylinder) (Output 4)
    parImpShaftRearEnd = parImpShaftRearEnd - (0.05 * parImpShaftFrontEnd + 0.1 * parImpShaftFrontEnd + 0.15 * parImpShaftFrontEnd)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set 0.25 * parImpShaftRearEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dImpellerDia
    
    'Create the Impeller part 2(Cone) (Output 5)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.05 * parImpShaftRearEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCone arrayOfOutputs(iOutput), oStPoint, oEnPoint, dImpellerDia, 0.9 * dImpellerDia
    
    'Create the Impeller part 3 (Cylinder) (Output 6)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.55 * parImpShaftRearEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.9 * dImpellerDia
    
    'Create the Flange at Rear End Of the Pump (Output 7)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.05 * parImpShaftFrontEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.8 * dImpellerDia
    
    'Create the Cylinder After the Flange at Rear End (Output 8)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.1 * parImpShaftFrontEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.5 * dImpellerDia
    
    'Create the Cylindet at Rear End of the Impeller (Output 9)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.15 * parImpShaftFrontEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.35 * dImpellerDia
    
    'Create the Coupling Cylinder 1 (Output 10)
    oStPoint.Set oEnPoint.x + 0.05 * parImpShaftRearEnd, 0, 0
    oEnPoint.Set oStPoint.x + 0.1 * parImpShaftRearEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.15 * dImpellerDia
    
    'Create the Coupling (Output 11)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + parDistBetShaftEnds, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.2 * dImpellerDia
    
    'Create the Coupling Cylinder 2 (Output 12)
    oStPoint.Set oEnPoint.x, 0, 0
    oEnPoint.Set oStPoint.x + 0.1 * parImpShaftRearEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.15 * dImpellerDia
    
    'Create the Motor Flange (Output 13)
    parMotorShaftToEnd = parMotorShaftToEnd - 0.1 * parImpShaftRearEnd
    oStPoint.Set oEnPoint.x + 0.1 * parMotorShaftToEnd, 0, 0
    oEnPoint.Set oStPoint.x + 0.025 * parMotorShaftToEnd, 0, 0
    Dim dShaftEnd As Double
    dShaftEnd = oStPoint.x
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.15 * dImpellerDia
    
    'Create the Motor (Output 14)
    Dim objMotor As Object
    oStPoint.Set oEnPoint.x, parMotorWidth / 2, parMotorHeight / 2
    oEnPoint.Set oStPoint.x + 0.875 * parMotorShaftToEnd, -oStPoint.y, -parMotorHeight / 2
    Set objMotor = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMotor
    Set objMotor = Nothing
    
    'Create the Support for Motor (Output 15)
    Dim ObjMotorSup As Object
    oStPoint.Set oStPoint.x + 0.1 * parMotorShaftToEnd, 0.4 * parMotorWidth, -parMotorHeight / 2
    oEnPoint.Set oEnPoint.x - 0.1 * parMotorShaftToEnd, -0.4 * parMotorWidth, _
                                -(parBasePlateBottom - parBasePlateHeight)
    Set ObjMotorSup = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotorSup
    Set ObjMotorSup = Nothing
    
    'Create the Junction Box for Motor (Output 16)
    Dim objMotorBox As Object
    Dim dJunctionBoxHt As Double
    Dim dJunctionBoxWidth As Double
    Dim dJunctionBoxLength As Double
    
    dJunctionBoxHt = parJunBoxHeight
    dJunctionBoxWidth = parJunBoxWidth
    dJunctionBoxLength = parJunBoxLength
    
    oStPoint.Set oStPoint.x + 0.1 * parMotorShaftToEnd, parMotorWidth / 2, dJunctionBoxHt / 2
    oEnPoint.Set oStPoint.x + dJunctionBoxLength, oStPoint.y + dJunctionBoxWidth, -dJunctionBoxHt / 2
    
    Set objMotorBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMotorBox
    Set objMotorBox = Nothing
    
    'Create the Shaft (Output 17)
    oStPoint.Set -parImpShaftFrontEnd, 0, 0
    oEnPoint.Set dShaftEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.1 * dImpellerDia
    
    'Create the Supports for Pump (Output 18)
    Dim ObjPumpSupColl As New Collection
    Dim oTopSurPts(0 To 3) As IJDPosition
    Dim oBotSurPts(0 To 3) As IJDPosition
    Dim iCount As Integer
    Dim jCount As Integer
    Dim dTempDistY As Double
    
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = New DPosition
        Set oBotSurPts(iCount) = New DPosition
    Next iCount
        
    dTempDistY = parMotorWidth
    For iCount = 0 To 1
        oTopSurPts(0).Set -0.55 * parImpShaftFrontEnd, 0.7 * dTempDistY / 2, 0
        oTopSurPts(1).Set -0.4 * parImpShaftFrontEnd, 0.7 * dTempDistY / 2, 0
        oTopSurPts(2).Set -0.4 * parImpShaftFrontEnd, dTempDistY / 2, 0
        oTopSurPts(3).Set -0.55 * parImpShaftFrontEnd, dTempDistY / 2, 0

        oBotSurPts(0).Set -0.55 * parImpShaftFrontEnd, 0.7 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)
        oBotSurPts(1).Set -0.4 * parImpShaftFrontEnd, 0.7 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)
        oBotSurPts(2).Set -0.4 * parImpShaftFrontEnd, 1.2 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)
        oBotSurPts(3).Set -0.55 * parImpShaftFrontEnd, 1.2 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)

        Set ObjPumpSupColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        'Set the Output
        If iCount = 0 Then
            iOutput = iOutput + 1
        End If
        For jCount = 1 To ObjPumpSupColl.Count
            m_OutputColl.AddOutput "PumpSupport_", ObjPumpSupColl(jCount)
        Next jCount
        Set ObjPumpSupColl = Nothing
        dTempDistY = -dTempDistY
    Next iCount
    
    dTempDistY = 0.9 * parMotorWidth
    For iCount = 0 To 1
        oTopSurPts(0).Set 0.75 * parImpShaftRearEnd - 0.15 * parImpShaftFrontEnd, 0.7 * dTempDistY / 2, 0
        oTopSurPts(1).Set 0.75 * parImpShaftRearEnd, 0.7 * dTempDistY / 2, 0
        oTopSurPts(2).Set 0.75 * parImpShaftRearEnd, dTempDistY / 2, 0
        oTopSurPts(3).Set 0.75 * parImpShaftRearEnd - 0.15 * parImpShaftFrontEnd, dTempDistY / 2, 0

        oBotSurPts(0).Set 0.75 * parImpShaftRearEnd - 0.15 * parImpShaftFrontEnd, 0.7 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)
        oBotSurPts(1).Set 0.75 * parImpShaftRearEnd, 0.7 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)
        oBotSurPts(2).Set 0.75 * parImpShaftRearEnd, 1.2 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)
        oBotSurPts(3).Set 0.75 * parImpShaftRearEnd - 0.15 * parImpShaftFrontEnd, 1.2 * dTempDistY / 2, -(parBasePlateBottom - parBasePlateHeight)

        Set ObjPumpSupColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        'Set the Output
        For jCount = 1 To ObjPumpSupColl.Count
            m_OutputColl.AddOutput "PumpSupport_", ObjPumpSupColl(jCount)
        Next jCount
        Set ObjPumpSupColl = Nothing
        dTempDistY = -dTempDistY
    Next iCount
    
    'Create Base Plate (Output 19)
    If CmpDblGreaterthan(parBasePlateHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBasePlateLength, LINEAR_TOLERANCE) _
                And CmpDblGreaterthan(parBasePlateWidth, LINEAR_TOLERANCE) Then
        Dim ObjBasePlateColl As Collection
        
        oTopSurPts(0).Set -parCenToFouPortStart - dInch, parBasePlateWidth / 2, -(parBasePlateBottom - parBasePlateHeight)
        oTopSurPts(1).Set oTopSurPts(0).x + parBasePlateLength, oTopSurPts(0).y, oTopSurPts(0).z
        oTopSurPts(2).Set oTopSurPts(1).x, -parBasePlateWidth / 2, oTopSurPts(0).z
        oTopSurPts(3).Set oTopSurPts(0).x, oTopSurPts(2).y, oTopSurPts(0).z
        
        oBotSurPts(0).Set oTopSurPts(0).x, oTopSurPts(0).y, -parBasePlateBottom
        oBotSurPts(1).Set oTopSurPts(1).x, oTopSurPts(1).y, -parBasePlateBottom
        oBotSurPts(2).Set oTopSurPts(2).x, oTopSurPts(2).y, -parBasePlateBottom
        oBotSurPts(3).Set oTopSurPts(3).x, oTopSurPts(3).y, -parBasePlateBottom
        
        Set ObjBasePlateColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To ObjBasePlateColl.Count
            m_OutputColl.AddOutput "BasePlate_", ObjBasePlateColl(iCount)
        Next iCount
        
        'Create Default Surface (Output 20)
        Dim ObjDefSurface As Object
        Set ObjDefSurface = ObjBasePlateColl(2)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
        Set ObjDefSurface = Nothing
        Set ObjBasePlateColl = Nothing
        
        'Create the Edges and Points for Base Plate (Output 21)
        Dim ObjEdgeColl As New Collection
        
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                    oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                    oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                    oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                    oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                    oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                    oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                    oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                    oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                    oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                    oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                    oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                                    oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                    oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To ObjEdgeColl.Count
            m_OutputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
        Next iCount
        Set ObjEdgeColl = Nothing
        
        'Creating the Points on each Face of the Base Plate (Output 22)
        Dim ObjPointColl As New Collection
        
        'Points on Right and Left surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                (oTopSurPts(1).x + oBotSurPts(2).x) / 2, (oTopSurPts(1).y + oBotSurPts(2).y) / 2, (oTopSurPts(1).z + oBotSurPts(2).z) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                (oTopSurPts(0).x + oBotSurPts(3).x) / 2, (oTopSurPts(0).y + oBotSurPts(3).y) / 2, (oTopSurPts(0).z + oBotSurPts(3).z) / 2)
        'Points on Front and Back surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                (oTopSurPts(3).x + oBotSurPts(2).x) / 2, (oTopSurPts(3).y + oBotSurPts(2).y) / 2, (oTopSurPts(3).z + oBotSurPts(2).z) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                (oTopSurPts(0).x + oBotSurPts(1).x) / 2, (oTopSurPts(0).y + oBotSurPts(1).y) / 2, (oTopSurPts(0).z + oBotSurPts(1).z) / 2)
        'Points on Top and Bottom surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                (oTopSurPts(0).x + oTopSurPts(2).x) / 2, (oTopSurPts(0).y + oTopSurPts(2).y) / 2, (oTopSurPts(0).z + oTopSurPts(2).z) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                                (oBotSurPts(0).x + oBotSurPts(2).x) / 2, (oBotSurPts(0).y + oBotSurPts(2).y) / 2, (oBotSurPts(0).z + oBotSurPts(2).z) / 2)
        'Set the output
        iOutput = iOutput + 1
        For iCount = 1 To ObjPointColl.Count
            m_OutputColl.AddOutput "Points_", ObjPointColl(iCount)
        Next iCount
        Set ObjPointColl = Nothing
    End If
   
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = Nothing
        Set oBotSurPts(iCount) = Nothing
    Next iCount
    Set oGeomFactory = Nothing
   
    'Create Equipment Foundation Port at the Bottom of Base Plate (Output 23)
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |   X   |
    '   |                  |   ^   |
    '   |                  |   |   |
    '   -----> X           |   |   |Port CS
    '   Symbol CS          |   |----> Y
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |-------|
    
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = parYBoltHole - parCenToFouPortStart
    dOrigin(1) = 0
    dOrigin(2) = -parBasePlateBottom
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())

    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call ObjFoundationPort.SetHoles(holes)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Insert your code for Point at origin (Output 24)
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    '==========================================
    '    Construction of  Control Point
    '==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, _
                                                        0, 0, 0, 0.02, , , False)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub



