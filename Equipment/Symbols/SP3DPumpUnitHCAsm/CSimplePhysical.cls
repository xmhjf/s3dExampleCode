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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:21 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          MS
'   Creation Date:  Tuesday, 22 April 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003    SymbolTeam(India)     Copyright Information, Header  is added.
'   23.Jul.2003    svsmylav/ACM          TR 44842: 1.Removed Default surface from Reference Aspect and added it to
'                                                               Physical Aspect.
'                                                               2. Made pump base which was a Ruled surface to 6 planes.
'                                                               3. For Trepizoidal base the planes on the base are inclined,
'                                                               Added two more Referance planes in X-Z and Y-Z planes at center
'                                                               of the pump base for constraining in those directions.
'   22.Aug.2003    SymbolTeam(India)     TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                        Modified port index logic.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "SimplePhysical:" 'Used for error messages

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

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
    
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim ObjSupport  As Object
    Dim ObjEllipticalCap As Object
    
    Dim iOutput     As Double
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parDisNozzletoFace      As Double
    Dim parDisNozzletoChFace As Double
    Dim parDischargetoFaceSuction       As Double
    Dim parBaseplatetoSuction       As Double
    Dim parPumptoFaceDisNozzle      As Double
    Dim parDisNozzletoFaceBaseplate         As Double
    Dim parMotorLength      As Double
    Dim parCouplingLength       As Double
    Dim parSupportLength        As Double
    Dim parSupportHeight        As Double
    Dim parSupportWidth         As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    
    Dim parBaseType As String

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parDisNozzletoFace = arrayOfInputs(2)
    parDisNozzletoChFace = arrayOfInputs(3)
    parDischargetoFaceSuction = arrayOfInputs(4)
    parBaseplatetoSuction = arrayOfInputs(5)
    parPumptoFaceDisNozzle = arrayOfInputs(6)
    parDisNozzletoFaceBaseplate = arrayOfInputs(7)
    parMotorLength = arrayOfInputs(8)
    parCouplingLength = arrayOfInputs(9)
    parSupportLength = arrayOfInputs(10)
    parSupportHeight = arrayOfInputs(11)
    parSupportWidth = arrayOfInputs(12)
    parXboltHole = arrayOfInputs(13)
    parYboltHole = arrayOfInputs(14)
    parBaseType = arrayOfInputs(15)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
     iOutput = 0
    
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
' Insert your code for output 1(Motor Body)
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    'Motor is assumed with elliptical heads. Motor consist of a cylinder and two elliptical caps on both sides.
    'The motor diameter is assumed to be 1.6 times the parBaseplatetoSuction
    Dim dMotorDiameter As Double
    dMotorDiameter = 1.6 * parBaseplatetoSuction
    
    'The Motor length includes the small shaft portion. So elliptical head is taken as 0.15*parMotorLength and shaftlength as 0.1* parMotorLength
    stPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.25 * parMotorLength, 0
    enPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.85 * parMotorLength, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dMotorDiameter
    
' Insert your code for output 2 (Motor Elliptical Rear Cap)
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
    Dim centerPos As AutoMath.DPosition
    
    Set revAxis = New AutoMath.DVector
    Set centerPos = New AutoMath.DPosition
    
    centerX = 0
    centerY = parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.85 * parMotorLength
    centerZ = 0
    
    normalX = 0
    normalY = 0
    normalZ = 1
    
    MajorX = dMotorDiameter / 2
    MajorY = 0
    MajorZ = 0
    
    mMRatio = (0.15 * parMotorLength) / (dMotorDiameter / 2)
    Set oEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                    centerX, centerY, centerZ, normalX, normalY, normalZ, _
                                    MajorX, MajorY, MajorZ, mMRatio, 0, PI / 2)
    revAxis.Set 0, 1, 0
    centerPos.Set 0, centerY, 0
    Set ObjEllipticalCap = PlaceRevolution(m_OutputColl, oEllipse, revAxis, centerPos, PI * 2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipticalCap
    Set ObjEllipticalCap = Nothing
'   Remove the Ellipse
    Dim ObjTemp As IJDObject
    Set ObjTemp = oEllipse
    Set oEllipse = Nothing
    ObjTemp.Remove
    Set ObjTemp = Nothing

' Insert your code for output 3 (Motor Elliptical Near Cap)
    centerX = 0
    centerY = parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.25 * parMotorLength
    centerZ = 0
    
    normalX = 0
    normalY = 0
    normalZ = -1
    
    MajorX = dMotorDiameter / 2
    MajorY = 0
    MajorZ = 0
    
    mMRatio = (0.15 * parMotorLength) / (dMotorDiameter / 2)
    Set oEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                    centerX, centerY, centerZ, normalX, normalY, normalZ, _
                                    MajorX, MajorY, MajorZ, mMRatio, PI / 2, PI / 2)
    revAxis.Set 0, 1, 0
    centerPos.Set 0, centerY, 0
    Set ObjEllipticalCap = PlaceRevolution(m_OutputColl, oEllipse, revAxis, centerPos, PI * 2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEllipticalCap
    Set ObjEllipticalCap = Nothing
'   Remove the Ellipse
    Set ObjTemp = oEllipse
    Set oEllipse = Nothing
    ObjTemp.Remove
    Set ObjTemp = Nothing
    Set revAxis = Nothing
    Set centerPos = Nothing

' Insert your code for output 4(Shaft Coupling)
'The shaft diameter is assumed to be 0.075 times the dMotorDiameter and coupling diameter as twice the shaft diameter
    Dim dShaftDiameter As Double
    dShaftDiameter = 0.075 * dMotorDiameter
    stPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength, 0
    enPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction, 0

    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * dShaftDiameter
                                
' Insert your code for output 5(Motor Shaft)
    stPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.1 * parMotorLength, 0
    enPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction - 0.1 * parMotorLength, 0

   iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dShaftDiameter
    
' Insert your code for output 6(Pump Cylinder 1)
    'The cylinder 1 diameter is assumed to be 1.2 times the parBaseplatetoSuction
    Dim dCylinderDiameter As Double
    dCylinderDiameter = 1.2 * parBaseplatetoSuction
    
    stPoint.Set 0, parDisNozzletoFace + parDischargetoFaceSuction - 0.1 * parMotorLength, -0.2 * parBaseplatetoSuction
    enPoint.Set 0, parDischargetoFaceSuction + parDisNozzletoChFace, -0.2 * parBaseplatetoSuction

   iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter
    
' Insert your code for output 7(Pump Cylinder 2)
    'The cylinder 2 diameter is assumed to be 1.6 times the parBaseplatetoSuction
    Dim dCylinderDiameter1 As Double
    dCylinderDiameter1 = 1.6 * parBaseplatetoSuction
    'The face of the flange is assumed to be 0.3 times the parDisNozzletoChFace
    stPoint.Set 0, parDischargetoFaceSuction + parDisNozzletoChFace, 0
    enPoint.Set 0, parDischargetoFaceSuction + 0.3 * parDisNozzletoChFace, 0

   iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter1
                                
' Insert your code for output 8(Pump Cylinder 3)As a Cylinder
    'The cylinder 3 diameter is assumed to be 1.95 times the parBaseplatetoSuction
    'The thickness of the flange is assumed to be 0.075 times the parDisNozzletoChFace
    Dim dCylinderDiameter2 As Double
    dCylinderDiameter2 = 1.95 * parBaseplatetoSuction
    stPoint.Set 0, parDischargetoFaceSuction + 0.3 * parDisNozzletoChFace, 0
    enPoint.Set 0, parDischargetoFaceSuction + 0.225 * parDisNozzletoChFace, 0

    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylinderDiameter2
                                
' Insert your code for output 9(Pump Cylinder 4)
    'The cylinder 4 diameter is assumed to be 1.8 times the smallest between parPumptoFaceDisNozzle and parBaseplatetoSuction
    'The cylinder 4  is assumed to go 0.5 * parDischargetoFaceSuction away from the discharge centreline
    Dim dPumpDiameter As Double
    dPumpDiameter = 1.8 * SmallerDim(parPumptoFaceDisNozzle, parBaseplatetoSuction)
    stPoint.Set 0, parDischargetoFaceSuction + 0.225 * parDisNozzletoChFace, 0
    enPoint.Set 0, 0.5 * parDischargetoFaceSuction, 0

   iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dPumpDiameter

' Insert your code for output 10 to 15(Support Body as 6 planes)
    ' The support under the pump is taken as 0.1 *parSupportHeight
    Dim dSpaceBelowDisNozz As Double
    dSpaceBelowDisNozz = 0.1 * parSupportHeight
    
    Dim dBaseCenter As New AutoMath.DPosition
    Dim ObjTrapSupportBody As Object
    Dim dBLength As Double
    Dim dBWidth As Double
    Dim dTLength As Double
    Dim dTWidth As Double
    Dim dHeight As Double
    dBaseCenter.Set 0, parDischargetoFaceSuction + (parSupportLength / 2 - parDisNozzletoFaceBaseplate), -parBaseplatetoSuction - parSupportHeight
    dBLength = parSupportLength
    dBWidth = parSupportWidth
    'The PolyShield support base is trapezoidal
    'The top length of the trapezoidal support is taken as  0.9 * parSupportLength
    'The top width of the trapezoidal support is taken as  0.9 * parSupportWidth
    
    If parBaseType = "PolyShield" Then
        dTLength = 0.9 * parSupportLength
        dTWidth = 0.9 * parSupportWidth
    Else
        dTLength = parSupportLength
        dTWidth = parSupportWidth
    End If
    dHeight = parSupportHeight - dSpaceBelowDisNozz
    Dim ObjColl As New Collection
    Dim oTopPts(0 To 3) As IJDPosition
    Dim oBotPts(0 To 3) As IJDPosition

    Dim iCount As Integer
    For iCount = 0 To 3
        Set oTopPts(iCount) = New DPosition
        Set oBotPts(iCount) = New DPosition
    Next iCount
    
    oTopPts(0).Set dBaseCenter.x - dTWidth / 2, dBaseCenter.y + dTLength / 2, dBaseCenter.z + dHeight
    oTopPts(1).Set oTopPts(0).x, oTopPts(0).y - dTLength, oTopPts(0).z
    oTopPts(2).Set oTopPts(1).x + dTWidth, oTopPts(1).y, oTopPts(0).z
    oTopPts(3).Set oTopPts(2).x, oTopPts(2).y + dTLength, oTopPts(0).z
        
    oBotPts(0).Set dBaseCenter.x - dBWidth / 2, dBaseCenter.y + dBLength / 2, dBaseCenter.z
    oBotPts(1).Set oBotPts(0).x, oBotPts(0).y - dBLength, oBotPts(0).z
    oBotPts(2).Set oBotPts(1).x + dBWidth, oBotPts(1).y, oBotPts(0).z
    oBotPts(3).Set oBotPts(2).x, oBotPts(2).y + dBLength, oBotPts(0).z
        
    Set ObjColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopPts, oBotPts)

' Set the output
    For iCount = 1 To ObjColl.Count
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjColl(iCount)
    Next iCount
    Set ObjColl = Nothing
    
    For iCount = 0 To 3
        Set oTopPts(iCount) = Nothing
        Set oBotPts(iCount) = Nothing
    Next iCount

' Insert your code for output 16(Reference Plane 1)
    Dim ObjRefPlane As Object
    Dim dPlanePoints(0 To 11) As Double
    
    dPlanePoints(0) = dBaseCenter.x
    dPlanePoints(1) = dBaseCenter.y + dBLength / 4
    dPlanePoints(2) = dBaseCenter.z
    
    dPlanePoints(3) = dPlanePoints(0)
    dPlanePoints(4) = dBaseCenter.y - dBLength / 4
    dPlanePoints(5) = dPlanePoints(2)
    
    dPlanePoints(6) = dPlanePoints(0)
    dPlanePoints(7) = dPlanePoints(4)
    dPlanePoints(8) = dBaseCenter.z + dHeight / 2
    
    dPlanePoints(9) = dPlanePoints(0)
    dPlanePoints(10) = dPlanePoints(1)
    dPlanePoints(11) = dPlanePoints(8)
    
    Set ObjRefPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dPlanePoints)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRefPlane
    Set ObjRefPlane = Nothing
    

' Insert your code for output 17(Reference Plane 2)
    
    dPlanePoints(0) = dBaseCenter.x - dBWidth / 4
    dPlanePoints(1) = dBaseCenter.y
    dPlanePoints(2) = dBaseCenter.z
    
    dPlanePoints(3) = dBaseCenter.x + dBWidth / 4
    dPlanePoints(4) = dPlanePoints(1)
    dPlanePoints(5) = dPlanePoints(2)
    
    dPlanePoints(6) = dPlanePoints(3)
    dPlanePoints(7) = dPlanePoints(1)
    dPlanePoints(8) = dBaseCenter.z + dHeight / 2
    
    dPlanePoints(9) = dPlanePoints(0)
    dPlanePoints(10) = dPlanePoints(1)
    dPlanePoints(11) = dPlanePoints(8)
    
    Set ObjRefPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dPlanePoints)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRefPlane
    Set ObjRefPlane = Nothing

    
' Insert your code for output 18(support under the pump discharge nozzle)
'All supports width under pump and motor are assumed to be 0.8 times their respective diameters
    Dim ObjSuppBox As Object
    stPoint.Set -dPumpDiameter * 0.4, parDischargetoFaceSuction + 0.1 * parDisNozzletoChFace, -parBaseplatetoSuction - dSpaceBelowDisNozz
    enPoint.Set dPumpDiameter * 0.4, 0.9 * parDischargetoFaceSuction, -dPumpDiameter / 2
    Set ObjSuppBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSuppBox
    Set ObjSuppBox = Nothing
    
    
' Insert your code for output 19(support under the intermediate chamber)
    Dim ObjSuppBox1 As Object
    stPoint.Set -dCylinderDiameter * 0.4, parBaseplatetoSuction + parDisNozzletoChFace, -parBaseplatetoSuction - dSpaceBelowDisNozz
    enPoint.Set dCylinderDiameter * 0.4, parBaseplatetoSuction + 0.9 * parDisNozzletoChFace, -dCylinderDiameter1 / 2
    
    Set ObjSuppBox1 = PlaceBox(m_OutputColl, stPoint, enPoint)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSuppBox1
    Set ObjSuppBox1 = Nothing
                                 
' Insert your code for output 20(support under the motor)
    Dim ObjSuppBox2 As Object
    stPoint.Set -dMotorDiameter * 0.4, parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.25 * parMotorLength, -parBaseplatetoSuction - dSpaceBelowDisNozz
    enPoint.Set dMotorDiameter * 0.4, parDisNozzletoFace + parDischargetoFaceSuction + parCouplingLength + 0.75 * parMotorLength, -dMotorDiameter / 2
    Set ObjSuppBox2 = PlaceBox(m_OutputColl, stPoint, enPoint)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSuppBox2
    Set ObjSuppBox2 = Nothing
    
        
'==========================================
'Construction of Equipment Foundatiopn Port
'==========================================
      'Coordinate sysytem is at the bottom of support.
    '
    '                       |-------|
    '   Y                  |        |
    '   ^                  |        |
    '   |                  |        |
    '   |                  |        |
    '   |                  |        |
    '   -----> X        |        |
    '   Symbol CS  |        |
    '                       |        |
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
    dOrigin(1) = parDischargetoFaceSuction + (parSupportLength / 2 - parDisNozzletoFaceBaseplate)
    dOrigin(2) = -parBaseplatetoSuction - parSupportHeight

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

    holes(0, 1) = -parYboltHole / 2
    holes(0, 2) = parXboltHole
    holes(1, 1) = parYboltHole / 2
    holes(1, 2) = parXboltHole
    holes(2, 1) = parYboltHole / 2
    holes(2, 2) = -parXboltHole
    holes(3, 1) = -parYboltHole / 2
    holes(3, 2) = -parXboltHole
                                                        
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
        
    Call ObjFoundationPort.SetHoles(holes)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
' Insert your code for output 27(Bend of the drain nozzle)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim ObjBend As Object
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim radPosition3 As Double
    'The BendRadius is assumed to be 0.3 * parDischargetoFaceSuction
    Dim dBendRadius As Double
    dBendRadius = 0.3 * parDischargetoFaceSuction
    radPosition3 = 0.6 * parBaseplatetoSuction
    
    ' Construct a circle that will be used to create the revolution
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         0, dBendRadius, -radPosition3, _
                        0, 1, 0, _
                      pipeDiam / 2)
    ' Revolve the circle for 90 degree about the Z axis
    Dim revcentPoint As New AutoMath.DPosition 'revolution center point
    Dim revVector As New AutoMath.DVector 'rotation vector for rotation
    revcentPoint.Set dBendRadius, dBendRadius, -radPosition3
    revVector.Set 0, 0, 1
    
    Set ObjBend = PlaceRevolution(m_OutputColl, objCircle, _
                    revVector, revcentPoint, PI / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBend
    Set ObjBend = Nothing
    Set geomFactory = Nothing
    Set revcentPoint = Nothing
    Set revVector = Nothing
    Set objCircle = Nothing

' Insert your code for output 28(Cylinder of the Drain Nozzle)
    Dim ObjBendCyl As Object
    stPoint.Set 0, 0.5 * parDischargetoFaceSuction, -radPosition3
    enPoint.Set 0, dBendRadius, -radPosition3
    
    Set ObjBendCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBendCyl
    Set ObjBendCyl = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
