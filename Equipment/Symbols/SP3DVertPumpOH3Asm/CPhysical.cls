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
'++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: VRK
'   Creation Date:  Wednesday, Jan 23 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy                 who                     change description
'   -----------                 ---                     ------------------
'   16.Oct.2007                 VRK                     TR-129143:Unable to bulkload foundation port data.
'*****************************************************************************

Option Explicit
Private PI As Double
Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "CPhysical:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    Set m_oGeomHelper = New SymbolServices
     PI = Atn(1) * 4
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim Center As IJDPosition
    Dim iOutput     As Double
    Dim parMotorHeight As Double
    Dim parBaseplatebottomtoMotorFace As Double
    Dim parBaseplatebottomtoDischargeCL As Double
    Dim parCLofPumptoSuctionFace As Double
    Dim parCLofPumptoDischargeFace As Double
    Dim parBaseplateLength As Double
    Dim parBaseplateWidth As Double
    Dim parBaseplateHeight As Double
    Dim parMotorDiameter As Double
    Dim parCLofDischargetoMotorface As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    Dim dBaseCenter As New AutoMath.DPosition

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
  
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parMotorHeight = arrayOfInputs(2)                        'AG
    parMotorDiameter = arrayOfInputs(3)                      'd
    parBaseplatebottomtoMotorFace = arrayOfInputs(4)         'C
    parBaseplatebottomtoDischargeCL = arrayOfInputs(5)       'B
    parCLofPumptoSuctionFace = arrayOfInputs(6)              'S
    parCLofPumptoDischargeFace = arrayOfInputs(7)            'D
    parBaseplateHeight = arrayOfInputs(8)
    parBaseplateWidth = arrayOfInputs(9)                      'E
    parBaseplateLength = arrayOfInputs(10)                    'L
    parXboltHole = arrayOfInputs(11)                          'G
    parYboltHole = arrayOfInputs(12)                          'F
    
    m_oGeomHelper.OutputCollection = m_outputColl

    Set Center = New DPosition
    Center.Set 0, 0, 0
    iOutput = 0
    parCLofDischargetoMotorface = parBaseplatebottomtoMotorFace - parBaseplatebottomtoDischargeCL
    
    'Insert your code for output 1
    Dim dCylinderDiameter1 As Double
    dCylinderDiameter1 = 0.8 * parMotorDiameter
    stPoint.Set 0, 0, parCLofDischargetoMotorface
    enPoint.Set 0, 0, parCLofDischargetoMotorface + 0.02 * parMotorHeight
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter1
    
    'Insert your code for output 2
    Dim dCylinderDiameter2 As Double
    dCylinderDiameter2 = 0.75 * parMotorDiameter
    stPoint.Set 0, 0, parCLofDischargetoMotorface + 0.02 * parMotorHeight
    enPoint.Set 0, 0, parCLofDischargetoMotorface + 0.05 * parMotorHeight
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter2
    
    'Insert your code for output 3(MotorBody)
    'Motor is assumed with elliptical heads. Motor consist of a cylinder and two elliptical caps on both sides.
    stPoint.Set 0, 0, parCLofDischargetoMotorface + 0.2 * parMotorHeight
    enPoint.Set 0, 0, parCLofDischargetoMotorface + 0.85 * parMotorHeight
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parMotorDiameter

    'Insert your code for output 4 (Motor Elliptical Rear Cap)
    Dim oEllipse As IngrGeom3D.EllipticalArc3d
    Dim centerX As Double
    Dim centerY As Double
    Dim centerZ As Double
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim revAxis As AutoMath.DVector
    Dim ObjEllipticalCap As Object
    Dim CenterPos As AutoMath.DPosition

    Set revAxis = New AutoMath.DVector
    Set CenterPos = New AutoMath.DPosition
    
    centerX = 0
    centerY = 0
    centerZ = parCLofDischargetoMotorface + 0.2 * parMotorHeight
    
    normalX = 0
    normalY = 1
    normalZ = 0
    
    MajorX = parMotorDiameter / 2
    MajorY = 0
    MajorZ = 0
    
    mMRatio = (0.15 * parMotorHeight) / (parMotorDiameter / 2)
    Set oEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                    centerX, centerY, centerZ, normalX, normalY, normalZ, _
                                    MajorX, MajorY, MajorZ, mMRatio, 0, PI / 2)
    revAxis.Set 0, 0, 1
    CenterPos.Set 0, centerY, 0
    Set ObjEllipticalCap = PlaceRevolution(m_outputColl, oEllipse, revAxis, CenterPos, PI * 2, True)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipticalCap
    Set ObjEllipticalCap = Nothing
    
    'Remove the Ellipse
    Dim ObjTemp As IJDObject
    Set ObjTemp = oEllipse
    Set oEllipse = Nothing
    ObjTemp.Remove
    Set ObjTemp = Nothing

    'Insert your code for output 5 (Motor Elliptical Near Cap)
    centerX = 0
    centerY = 0
    centerZ = parCLofDischargetoMotorface + 0.85 * parMotorHeight
    
    normalX = 0
    normalY = -1
    normalZ = 0
    
    MajorX = parMotorDiameter / 2
    MajorY = 0
    MajorZ = 0
    
    mMRatio = (0.15 * parMotorHeight) / (parMotorDiameter / 2)
    Set oEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                    centerX, centerY, centerZ, normalX, normalY, normalZ, _
                                    MajorX, MajorY, MajorZ, mMRatio, PI / 2, PI / 2)
    revAxis.Set 0, 0, 1
    CenterPos.Set 0, centerY, 0
    Set ObjEllipticalCap = PlaceRevolution(m_outputColl, oEllipse, revAxis, CenterPos, PI * 2, True)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipticalCap
    Set ObjEllipticalCap = Nothing
    
    'Remove the Ellipse
    Set ObjTemp = oEllipse
    Set oEllipse = Nothing
    ObjTemp.Remove
    Set ObjTemp = Nothing
    Set revAxis = Nothing
    Set CenterPos = Nothing
    
    'Insert your code for output 6 to 11 (Baseplate)
    Dim topSurPts(3) As IJDPosition
    Dim botSurPts(3) As IJDPosition
    Dim iCount As Integer

    For iCount = 0 To 3
        Set topSurPts(iCount) = New DPosition
        Set botSurPts(iCount) = New DPosition
    Next iCount
    
    dBaseCenter.Set 0, 0, 0.95 * parCLofDischargetoMotorface

    'Point 1
    botSurPts(0).x = -parBaseplateLength / 2
    botSurPts(0).y = -parBaseplateWidth / 2
    botSurPts(0).z = dBaseCenter.z
    
    'Point 2
    botSurPts(1).x = parBaseplateLength / 2
    botSurPts(1).y = -parBaseplateWidth / 2
    botSurPts(1).z = dBaseCenter.z
    
    'Point3
    botSurPts(2).x = parBaseplateLength / 2
    botSurPts(2).y = parBaseplateWidth / 2
    botSurPts(2).z = dBaseCenter.z
    
    'Point 4
    botSurPts(3).x = -parBaseplateLength / 2
    botSurPts(3).y = parBaseplateWidth / 2
    botSurPts(3).z = dBaseCenter.z
   
    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, parCLofDischargetoMotorface
    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, parCLofDischargetoMotorface
    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, parCLofDischargetoMotorface
    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, parCLofDischargetoMotorface
    
    Dim ObjCollection As Collection

    Set ObjCollection = PlaceTrapezoidWithPlanes(m_outputColl, topSurPts, botSurPts)
    For iCount = 1 To 6
         iOutput = iOutput + 1
         m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount

    Set ObjCollection = Nothing
    Set dBaseCenter = Nothing
    Set topSurPts(3) = Nothing
    Set botSurPts(3) = Nothing
    
    'Insert your code for output 12 to 15
    'Insert four cylinders onthe baseplate
    Dim dCylinderDiameter4 As Double
    dCylinderDiameter4 = 0.15 * parBaseplateWidth

    stPoint.Set -parBaseplateLength / 2 + dCylinderDiameter4 / 2, -parBaseplateWidth / 2 + dCylinderDiameter4 / 2, -parBaseplatebottomtoDischargeCL + parBaseplateHeight
    enPoint.Set -parBaseplateLength / 2 + dCylinderDiameter4 / 2, -parBaseplateWidth / 2 + dCylinderDiameter4 / 2, 0.95 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter4
    
    stPoint.Set -parBaseplateLength / 2 + dCylinderDiameter4 / 2, parBaseplateWidth / 2 - dCylinderDiameter4 / 2, -parBaseplatebottomtoDischargeCL + parBaseplateHeight
    enPoint.Set -parBaseplateLength / 2 + dCylinderDiameter4 / 2, parBaseplateWidth / 2 - dCylinderDiameter4 / 2, 0.95 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter4

    stPoint.Set parBaseplateLength / 2 - dCylinderDiameter4 / 2, parBaseplateWidth / 2 - dCylinderDiameter4 / 2, -parBaseplatebottomtoDischargeCL + parBaseplateHeight
    enPoint.Set parBaseplateLength / 2 - dCylinderDiameter4 / 2, parBaseplateWidth / 2 - dCylinderDiameter4 / 2, 0.95 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter4

    stPoint.Set parBaseplateLength / 2 - dCylinderDiameter4 / 2, -parBaseplateWidth / 2 + dCylinderDiameter4 / 2, -parBaseplatebottomtoDischargeCL + parBaseplateHeight
    enPoint.Set parBaseplateLength / 2 - dCylinderDiameter4 / 2, -parBaseplateWidth / 2 + dCylinderDiameter4 / 2, 0.95 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter4
    
    'Insert your code for output 16
    Dim dCylinderDiameter3 As Double
    dCylinderDiameter3 = 0.8 * parBaseplateWidth
 
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, 0.05 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter3
    
    'Insert your code for output 17 to 19
    'InsideCylinders
    'The cylinder is assumed to be 0.6 times the baseplate width
    Dim dCylinderDiameter5 As Double
    dCylinderDiameter5 = 0.6 * parBaseplateWidth
    stPoint.Set 0, 0, 0.05 * parCLofDischargetoMotorface
    enPoint.Set 0, 0, 0.3 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter5
    
    'The cylinder is assumed to be 0.5 times the baseplate width
    Dim dCylinderDiameter6 As Double
    dCylinderDiameter6 = 0.5 * parBaseplateWidth
    stPoint.Set 0, 0, 0.3 * parCLofDischargetoMotorface
    enPoint.Set 0, 0, 0.6 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter6
    
    'The cylinder is assumed to be 0.55 times the baseplate width
    Dim dCylinderDiameter7 As Double
    dCylinderDiameter7 = 0.55 * parBaseplateWidth
    stPoint.Set 0, 0, 0.6 * parCLofDischargetoMotorface
    enPoint.Set 0, 0, 0.95 * parCLofDischargetoMotorface
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter7
       
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    'Insert your code for output 20 to 25 (Place Base)
    dBaseCenter.Set parBaseplateLength / 2, parBaseplateWidth / 2, -parBaseplatebottomtoDischargeCL

    For iCount = 0 To 3
        Set topSurPts(iCount) = New DPosition
        Set botSurPts(iCount) = New DPosition
    Next iCount
    
    
    'Point 1
    botSurPts(0).x = -parBaseplateLength / 2
    botSurPts(0).y = -parBaseplateWidth / 2
    botSurPts(0).z = dBaseCenter.z
    
    'Point 2
    botSurPts(1).x = parBaseplateLength / 2
    botSurPts(1).y = -parBaseplateWidth / 2
    botSurPts(1).z = dBaseCenter.z
    
    'Point3
    botSurPts(2).x = parBaseplateLength / 2
    botSurPts(2).y = parBaseplateWidth / 2
    botSurPts(2).z = dBaseCenter.z
    
    'Point 4
    botSurPts(3).x = -parBaseplateLength / 2
    botSurPts(3).y = parBaseplateWidth / 2
    botSurPts(3).z = dBaseCenter.z
    
    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parBaseplateHeight
    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, botSurPts(1).z + parBaseplateHeight
    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, botSurPts(2).z + parBaseplateHeight
    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, botSurPts(3).z + parBaseplateHeight
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_outputColl, topSurPts, botSurPts)
    For iCount = 1 To 6
         iOutput = iOutput + 1
         m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount

    Set ObjCollection = Nothing
    Set dBaseCenter = Nothing
    
    'Insert your code for output 28
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Coordinate sysytem is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |        |
    '   ^                  |        |
    '   |                  |        |
    '   |                  |        |
    '   |                  |        |
    '   -----> X           |        |
    '   Symbol CS          |        |
    '                      |        |
    '                       |   X   |
    '                       |   ^   |
    '                       |   |   |
    '                       |   |   |
    '                       |---|---|----> Y
    '                   Port CS
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = -parBaseplatebottomtoDischargeCL

    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_outputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())
                                                              
    holes(0, 1) = parYboltHole / 2
    holes(0, 2) = parXboltHole / 2
    holes(1, 1) = -parYboltHole / 2
    holes(1, 2) = parXboltHole / 2
    holes(2, 1) = -parYboltHole / 2
    holes(2, 2) = -parXboltHole / 2
    holes(3, 1) = parYboltHole / 2
    holes(3, 2) = -parXboltHole / 2
                                                   
                                                   
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
        
    Call ObjFoundationPort.SetHoles(holes)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Insert your code for output 29
    '=====================================================================
    'Construction of  Control Point at Pump Origin as sphere zero diameter
    '=====================================================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Center.Set 0, 0, 0
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_outputColl.ResourceManager, Center.x, Center.y, Center.z, 0, , , True)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    Set Center = Nothing
    
    'Adding LinePoints and edges on the Baseplate
     'Create line by two pints:
     Dim LinePoints(0 To 23) As Double
     Dim objPoint As Object
     Dim line As Object
   
     LinePoints(0) = -parBaseplateLength / 2
     LinePoints(1) = -parBaseplateWidth / 2
     LinePoints(2) = -parBaseplatebottomtoDischargeCL
     
     LinePoints(3) = parBaseplateLength / 2
     LinePoints(4) = -parBaseplateWidth / 2
     LinePoints(5) = -parBaseplatebottomtoDischargeCL
     
     LinePoints(6) = parBaseplateLength / 2
     LinePoints(7) = parBaseplateWidth / 2
     LinePoints(8) = -parBaseplatebottomtoDischargeCL
     
     LinePoints(9) = -parBaseplateLength / 2
     LinePoints(10) = parBaseplateWidth / 2
     LinePoints(11) = -parBaseplatebottomtoDischargeCL
     
     LinePoints(12) = -parBaseplateLength / 2
     LinePoints(13) = -parBaseplateWidth / 2
     LinePoints(14) = -parBaseplatebottomtoDischargeCL + parBaseplateHeight
     
     LinePoints(15) = parBaseplateLength / 2
     LinePoints(16) = -parBaseplateWidth / 2
     LinePoints(17) = -parBaseplatebottomtoDischargeCL + parBaseplateHeight
     
     LinePoints(18) = parBaseplateLength / 2
     LinePoints(19) = parBaseplateWidth / 2
     LinePoints(20) = -parBaseplatebottomtoDischargeCL + parBaseplateHeight
     
     LinePoints(21) = -parBaseplateLength / 2
     LinePoints(22) = parBaseplateWidth / 2
     LinePoints(23) = -parBaseplatebottomtoDischargeCL + parBaseplateHeight
     
    'Create salient points on each face of the body
    Dim x, y, z As Double
    x = (LinePoints(0) + LinePoints(6)) / 2
    y = (LinePoints(1) + LinePoints(7)) / 2
    z = (LinePoints(2) + LinePoints(8)) / 2
    'Bottom Point
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    x = (LinePoints(12) + LinePoints(18)) / 2
    y = (LinePoints(13) + LinePoints(19)) / 2
    z = (LinePoints(14) + LinePoints(20)) / 2
    'Top Point
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    x = (LinePoints(0) + LinePoints(15)) / 2
    y = (LinePoints(1) + LinePoints(16)) / 2
    z = (LinePoints(2) + LinePoints(17)) / 2
    'Left Face Point
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    x = (LinePoints(6) + LinePoints(21)) / 2
    y = (LinePoints(7) + LinePoints(22)) / 2
    z = (LinePoints(8) + LinePoints(23)) / 2
    'Right Face Point
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    x = (LinePoints(0) + LinePoints(21)) / 2
    y = (LinePoints(1) + LinePoints(22)) / 2
    z = (LinePoints(2) + LinePoints(23)) / 2
    'Front Face Point
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    x = (LinePoints(3) + LinePoints(18)) / 2
    y = (LinePoints(4) + LinePoints(19)) / 2
    z = (LinePoints(5) + LinePoints(20)) / 2
    'Back Face Point
    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPoint
    Set objPoint = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(0), LinePoints(1), LinePoints(2), LinePoints(3), LinePoints(4), LinePoints(5))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(3), LinePoints(4), LinePoints(5), LinePoints(6), LinePoints(7), LinePoints(8))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(6), LinePoints(7), LinePoints(8), LinePoints(9), LinePoints(10), LinePoints(11))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(9), LinePoints(10), LinePoints(11), LinePoints(0), LinePoints(1), LinePoints(2))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(0), LinePoints(1), LinePoints(2), LinePoints(12), LinePoints(13), LinePoints(14))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(12), LinePoints(13), LinePoints(14), LinePoints(21), LinePoints(22), LinePoints(23))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(21), LinePoints(22), LinePoints(23), LinePoints(9), LinePoints(10), LinePoints(11))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(18), LinePoints(19), LinePoints(20), LinePoints(6), LinePoints(7), LinePoints(8))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
        
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(18), LinePoints(19), LinePoints(20), LinePoints(15), LinePoints(16), LinePoints(17))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(3), LinePoints(4), LinePoints(5), LinePoints(15), LinePoints(16), LinePoints(17))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(15), LinePoints(16), LinePoints(17), LinePoints(12), LinePoints(13), LinePoints(14))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, LinePoints(21), LinePoints(22), LinePoints(23), LinePoints(18), LinePoints(19), LinePoints(20))
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), line
    Set line = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

