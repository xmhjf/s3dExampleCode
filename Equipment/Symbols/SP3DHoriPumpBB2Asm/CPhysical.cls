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
'   Creation Date:  Tuesday, April 3 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Twenty Five Outputs:
'      In which Default Surface is created at the bottom of the Base Plate.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double
Private Const Inch = 15 * 2.54 / 100

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
    
    Dim parImpellerDiameter As Double
    Dim parDischargeNozLoc As Double
    Dim parSuctionNozLoc As Double
    Dim parCenterToNozzFace As Double
    Dim parShaftFrontEnd As Double
    Dim parCouplingStart As Double
    Dim parCoupEndToMotorEnd As Double
    Dim parCouplingLength As Double
    Dim parSupportBottomLoc As Double
    Dim parFounPortStartPoint As Double
    Dim parSupportHeight As Double
    Dim parSupportWidth As Double
    Dim parSupportLength As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parImpellerDiameter = arrayOfInputs(2)          'D
    parDischargeNozLoc = arrayOfInputs(3)           'Z
    parSuctionNozLoc = arrayOfInputs(4)             'S
    parCenterToNozzFace = arrayOfInputs(5)          'X
    parShaftFrontEnd = arrayOfInputs(6)             'Y
    parCouplingStart = arrayOfInputs(7)             'W
    parCoupEndToMotorEnd = arrayOfInputs(8)         'CMAX
    parCouplingLength = arrayOfInputs(9)            'DBSE
    parSupportBottomLoc = arrayOfInputs(10)         'HC
    parFounPortStartPoint = arrayOfInputs(11)       'HR
    parSupportHeight = arrayOfInputs(12)            'HG
    parSupportWidth = arrayOfInputs(13)             'HA
    parSupportLength = arrayOfInputs(14)            'HB
    parXBoltHole = arrayOfInputs(15)                'HF
    parYBoltHole = arrayOfInputs(16)                'HE
    
    'Origin is at DP1/PP1
    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim OstPoint As New AutoMath.DPosition
    Dim OenPoint As New AutoMath.DPosition
    
    'Create the Pump Cylinder 1 (Output 1)
    OstPoint.Set -parShaftFrontEnd, 0, 0
    OenPoint.Set -0.7 * parShaftFrontEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.5 * parImpellerDiameter
    
    'Create Pump Cylinder 2 (Output 2)
    OstPoint.Set -0.6 * parShaftFrontEnd, 0, 0
    OenPoint.Set -0.55 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.2 * parImpellerDiameter
    
    'Create Pump Cylinder 3 (Output 3)
    OstPoint.Set -0.55 * parShaftFrontEnd, 0, 0
    OenPoint.Set 0.55 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.5 * parImpellerDiameter
    
    'Create the Pump Cylinder 4 (Output 4)
    OstPoint.Set -0.45 * parShaftFrontEnd, 0, 0
    OenPoint.Set 0.45 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, parImpellerDiameter
    
    'Create the Pump Cylinder 5 (Output 5)
    OstPoint.Set -0.4 * parShaftFrontEnd, 0, 0
    OenPoint.Set -0.35 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.25 * parImpellerDiameter
    
    'Create the PumpCylinder 6 (Output 6)
    OstPoint.Set -0.35 * parShaftFrontEnd, 0, 0
    OenPoint.Set 0.35 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.25 * parImpellerDiameter
    
    'Create Pump Cylinder 7 (Output 7)
    OstPoint.Set 0.35 * parShaftFrontEnd, 0, 0
    OenPoint.Set 0.4 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.25 * parImpellerDiameter
    
    'Create the Pump Cylinder 8 (Output 8)
    OstPoint.Set 0.55 * parShaftFrontEnd, 0, 0
    OenPoint.Set 0.6 * parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.2 * parImpellerDiameter
    
    'Create the Pump Cylinder 9 (Output 9)
    OstPoint.Set 0.7 * parShaftFrontEnd, 0, 0
    OenPoint.Set parShaftFrontEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.5 * parImpellerDiameter
    
    'Create the Coupling Cylinder 10 (Output 10)
    OstPoint.Set 0.9 * parCouplingStart, 0, 0
    OenPoint.Set parCouplingStart, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.15 * parImpellerDiameter
    
    'Create the Coupling (Output 11)
    OstPoint.Set parCouplingStart, 0, 0
    OenPoint.Set parCouplingStart + parCouplingLength, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.2 * parImpellerDiameter
    
    'Create the Coupling Cylinder (Output 12)
    OstPoint.Set parCouplingStart + parCouplingLength, 0, 0
    OenPoint.Set 1.1 * parCouplingStart + parCouplingLength, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.15 * parImpellerDiameter
    
    'Create Motor Flange (Output 13)
    parCoupEndToMotorEnd = parCoupEndToMotorEnd - 0.1 * parCouplingStart
    OstPoint.Set 1.1 * parCouplingStart + parCouplingLength + 0.1 * parCoupEndToMotorEnd, 0, 0
    OenPoint.Set 1.1 * parCouplingStart + parCouplingLength + 0.125 * parCoupEndToMotorEnd, 0, 0
    'Set the OutPut
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.2 * parImpellerDiameter
    
    'Create Motor Front Head (Output 14)
    OstPoint.Set OenPoint.x, 0, 0
    OenPoint.Set OstPoint.x + 0.15 * parCoupEndToMotorEnd, 0, 0
    Dim dMotorSuppStart As Double
    dMotorSuppStart = OenPoint.x
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.9 * parImpellerDiameter
    
    'Create Motor (Output 15)
    OstPoint.Set OenPoint.x, 0, 0
    OenPoint.Set OstPoint.x + 0.425 * parCoupEndToMotorEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, parImpellerDiameter
    
    'Create Motor Rear Head (Output 16)
    OstPoint.Set OenPoint.x, 0, 0
    OenPoint.Set OstPoint.x + 0.3 * parCoupEndToMotorEnd, 0, 0
    Dim dMotorSuppEnd As Double
    dMotorSuppEnd = (OstPoint.x + OenPoint.x) / 2
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, parImpellerDiameter
    
    'Create Shaft (Output 17)
    OstPoint.Set -0.7 * parShaftFrontEnd, 0, 0
    OenPoint.Set 1.1 * parCouplingStart + parCouplingLength + 0.1 * parCoupEndToMotorEnd, 0, 0
    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), OstPoint, OenPoint, 0.1 * parImpellerDiameter
        
    'Create Motor Support (Output 18)
    Dim dCenToBasePlateStart As Double
    dCenToBasePlateStart = parSupportBottomLoc - parSupportHeight
    OstPoint.Set dMotorSuppStart, parImpellerDiameter / 3, -dCenToBasePlateStart
    OenPoint.Set dMotorSuppEnd, -parImpellerDiameter / 3, -0.5 * parImpellerDiameter
    Dim ObjMotorSup As Object
    Set ObjMotorSup = PlaceBox(m_OutputColl, OstPoint, OenPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMotorSup
    Set ObjMotorSup = Nothing
    
    'Create the Supports for Pump (Output 19)
    Dim ObjPumpSupColl As New Collection
    Dim oTopSurPts(0 To 3) As IJDPosition
    Dim oBotSurPts(0 To 3) As IJDPosition
    Dim iCount As Integer
    Dim jCount As Integer
    Dim dTempDist As Double
    
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = New DPosition
        Set oBotSurPts(iCount) = New DPosition
    Next iCount
    
    dTempDist = parImpellerDiameter
    For iCount = 0 To 1
        oTopSurPts(0).Set -0.55 * parShaftFrontEnd, 0.5 * dTempDist, 0
        oTopSurPts(1).Set 0.55 * parShaftFrontEnd, 0.5 * dTempDist, 0
        oTopSurPts(2).Set 0.55 * parShaftFrontEnd, 0.6 * dTempDist, 0
        oTopSurPts(3).Set -0.55 * parShaftFrontEnd, 0.6 * dTempDist, 0
        
        oBotSurPts(0).Set -0.55 * parShaftFrontEnd, 0.5 * dTempDist, -dCenToBasePlateStart
        oBotSurPts(1).Set 0.55 * parShaftFrontEnd, 0.5 * dTempDist, -dCenToBasePlateStart
        oBotSurPts(2).Set 0.55 * parShaftFrontEnd, 0.7 * dTempDist, -dCenToBasePlateStart
        oBotSurPts(3).Set -0.55 * parShaftFrontEnd, 0.7 * dTempDist, -dCenToBasePlateStart
    
        Set ObjPumpSupColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        'Set the Output
        If iCount = 0 Then
            iOutput = iOutput + 1
        End If
        For jCount = 1 To ObjPumpSupColl.Count
            m_OutputColl.AddOutput "PumpSupport_", ObjPumpSupColl(jCount)
        Next jCount
        Set ObjPumpSupColl = Nothing
        dTempDist = -parImpellerDiameter
    Next iCount

    'Create Base Plate (Output 20)
    If CmpDblGreaterthan(parSupportHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) _
                And CmpDblGreaterthan(parSupportWidth, LINEAR_TOLERANCE) Then
        Dim ObjBasePlateColl As Collection
        
        oTopSurPts(0).Set -parFounPortStartPoint - Inch, parSupportWidth / 2, -(parSupportBottomLoc - parSupportHeight)
        oTopSurPts(1).Set oTopSurPts(0).x + parSupportLength, oTopSurPts(0).y, oTopSurPts(0).z
        oTopSurPts(2).Set oTopSurPts(1).x, -parSupportWidth / 2, oTopSurPts(0).z
        oTopSurPts(3).Set oTopSurPts(0).x, oTopSurPts(2).y, oTopSurPts(0).z
        
        oBotSurPts(0).Set oTopSurPts(0).x, oTopSurPts(0).y, -parSupportBottomLoc
        oBotSurPts(1).Set oTopSurPts(1).x, oTopSurPts(1).y, -parSupportBottomLoc
        oBotSurPts(2).Set oTopSurPts(2).x, oTopSurPts(2).y, -parSupportBottomLoc
        oBotSurPts(3).Set oTopSurPts(3).x, oTopSurPts(3).y, -parSupportBottomLoc
        
        Set ObjBasePlateColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopSurPts, oBotSurPts)
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To ObjBasePlateColl.Count
            m_OutputColl.AddOutput "BasePlate_", ObjBasePlateColl(iCount)
        Next iCount
        
        'Create Default Surface (Output 21)
        Dim ObjDefSurface As Object
        Set ObjDefSurface = ObjBasePlateColl(2)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
        Set ObjDefSurface = Nothing
        Set ObjBasePlateColl = Nothing
        
        'Create the Edges and Points for Base Plate (Output 22)
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
        
        'Creating the Points on each Face of the Base Plate (Output 23)
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
    Set OstPoint = Nothing
    Set OenPoint = Nothing
    For iCount = 0 To 3
        Set oTopSurPts(iCount) = Nothing
        Set oBotSurPts(iCount) = Nothing
    Next iCount
    Set oGeomFactory = Nothing
   
    'Create Equipment Foundation Port at the Bottom of Base Plate (Output 24)
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
    dOrigin(0) = parYBoltHole - parFounPortStartPoint
    dOrigin(1) = 0
    dOrigin(2) = -parSupportBottomLoc
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
    
    'Insert your code for Point at origin (Output 25)
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

