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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: Veena
'   Creation Date:  Monday, Jan 08 2006
'
'   Description:
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Twenty Five Outputs:
'      In which Default Surface is created at the bottom of the Base Plate.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   14.Jan-2008     VRK             TR-128504:Model Motor with Cylinder
'******************************************************************************

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
    Dim parSuctiontoDischargeDistance As Double
    Dim parBaseplateBottomtoDischargeFace As Double
    Dim parDischargeCLtoSuctionFace As Double
    Dim parPumpSuctionFacetoCoupling As Double
    Dim parCouplingtoMotorEnd As Double
    Dim parBaseplateLength As Double
    Dim parBaseplateHeight As Double
    Dim parBaseplateWidth As Double
    Dim parOffsetfromSuctiontoCLofBaseplate As Double
    Dim parCouplingLength As Double
    Dim parBaseplateBottomtoSuctionCL As Double
    Dim parBaseplatetoSuction As Double
    Dim parMotorDia As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    Dim dBaseCenter As New AutoMath.DPosition

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
  
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition

    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSuctiontoDischargeDistance = arrayOfInputs(2)         'Z
    parBaseplateBottomtoDischargeFace = arrayOfInputs(3)     'H
    parDischargeCLtoSuctionFace = arrayOfInputs(4)           'e
    parPumpSuctionFacetoCoupling = arrayOfInputs(5)          'P
    parCouplingtoMotorEnd = arrayOfInputs(6)                 'K
    parBaseplateLength = arrayOfInputs(7)                    'L
    parBaseplateHeight = arrayOfInputs(8)                    'BU
    parBaseplateWidth = arrayOfInputs(9)                     'B
    parOffsetfromSuctiontoCLofBaseplate = arrayOfInputs(10)  'O
    parCouplingLength = arrayOfInputs(11)                    'y
    parBaseplateBottomtoSuctionCL = arrayOfInputs(12)        'h1
    parMotorDia = arrayOfInputs(13)                          'D
    parXboltHole = arrayOfInputs(14)
    parYboltHole = arrayOfInputs(15)
    parBaseplatetoSuction = parBaseplateBottomtoSuctionCL - parBaseplateHeight  '(h1-BU)
    
    m_oGeomHelper.OutputCollection = m_outputColl

    Set Center = New DPosition
    Center.Set 0, 0, 0
    iOutput = 0

    'Insert your code for output 1 to 6 (Place Base)
    dBaseCenter.Set 0, parBaseplateLength / 2, -parBaseplateBottomtoSuctionCL

    Dim oTopSurPts(3) As IJDPosition
    Dim oBotSurPts(3) As IJDPosition
    Dim iCount As Integer

    For iCount = 0 To 3
         Set oTopSurPts(iCount) = New DPosition
         Set oBotSurPts(iCount) = New DPosition
    Next iCount
    
    'Point 1
    oBotSurPts(0).x = dBaseCenter.x - parBaseplateWidth / 2
    oBotSurPts(0).y = dBaseCenter.y - parBaseplateLength / 2
    oBotSurPts(0).z = dBaseCenter.z

    'Point 2
    oBotSurPts(1).x = dBaseCenter.x + parBaseplateWidth / 2
    oBotSurPts(1).y = dBaseCenter.y - parBaseplateLength / 2
    oBotSurPts(1).z = dBaseCenter.z
    
    'Point3
    oBotSurPts(2).x = dBaseCenter.x + parBaseplateWidth / 2
    oBotSurPts(2).y = dBaseCenter.y + parBaseplateLength / 2
    oBotSurPts(2).z = dBaseCenter.z
    
    'Point 4
    oBotSurPts(3).x = dBaseCenter.x - parBaseplateWidth / 2
    oBotSurPts(3).y = dBaseCenter.y + parBaseplateLength / 2
    oBotSurPts(3).z = dBaseCenter.z
    
    oTopSurPts(0).Set oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z + parBaseplateHeight
    oTopSurPts(1).Set oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z + parBaseplateHeight
    oTopSurPts(2).Set oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z + parBaseplateHeight
    oTopSurPts(3).Set oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z + parBaseplateHeight
    
    Dim ObjCollection As Collection
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopSurPts, oBotSurPts)
    
    For iCount = 1 To 6
         iOutput = iOutput + 1
         m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount
    
    'Create Default Surface (Output 7)
    Dim ObjDefSurface As Object
    Set ObjDefSurface = ObjCollection(2)
    'Set the Output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing
    Set ObjCollection = Nothing
    Set dBaseCenter = Nothing
    
    'Insert your code for output 8(Pump Cylinder 1)
    'The cylinder 1 diameter is assumed to be  the parBaseplatetoSuction
    Dim dCylinderDiameter As Double
    dCylinderDiameter = parBaseplatetoSuction
    
    oStPoint.Set 0, 0.1 * parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter
    
    'Insert your code for output 9(Pump Cylinder 2)
    'The cylinder 1 diameter is assumed to be 0.8 times the parBaseplatetoSuction
    Dim dCylinderDiameter1 As Double
    dCylinderDiameter1 = 0.8 * parBaseplatetoSuction
    
    oStPoint.Set 0, 0.2 * parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0.1 * parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter1
    
    'Insert your code for output 10(Pump Cylinder 3)
    'The cylinder 2 diameter is assumed to be 0.3 times the parBaseplatetoSuction
    Dim dCylinderDiameter2 As Double
    dCylinderDiameter2 = 0.3 * parBaseplatetoSuction
    oStPoint.Set 0, 0.35 * parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0.2 * parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter2
                
    'Insert your code for output 11(Pump Cylinder 4)As a Cylinder
    'The cylinder 3 diameter is assumed to be 0.4 times the parBaseplatetoSuction
    Dim dCylinderDiameter3 As Double
    dCylinderDiameter3 = 0.4 * parBaseplatetoSuction
    oStPoint.Set 0, 0.45 * parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0.2 * parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter3
                
    'Insert your code for output 12(Pump Cylinder 5)
    'The cylinder 4 diameter is assumed to be 0.7 times the parBaseplatetoSuction
    Dim dCylinderDiameter4 As Double
    dCylinderDiameter4 = 0.7 * parBaseplatetoSuction
    oStPoint.Set 0, 0.5 * parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0.2 * parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter4
    
    'Insert your code for output 13(Pump Cylinder 6)
    'The cylinder 5 diameter is assumed to be 0.6 times parBaseplatetoSuction
    Dim dCylinderDiameter5 As Double
    dCylinderDiameter5 = 0.6 * parBaseplatetoSuction
    oStPoint.Set 0, 0.8 * parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0.5 * parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter5
    
    'Insert your code for output 14(Pump Cylinder7)
    'The cylinder 6 diameter is assumed to be 0.1 times  parBaseplatetoSuction
    Dim dCylinderDiameter6 As Double
    dCylinderDiameter6 = 0.1 * parBaseplatetoSuction
    oStPoint.Set 0, parPumpSuctionFacetoCoupling, 0
    oEnPoint.Set 0, 0.35 * parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter6
    
    'Insert your code for output 15(Motor Body)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Motor is assumed as Cylinder
    Dim dMotorDiameter As Double
    dMotorDiameter = parMotorDia
    
    oStPoint.Set 0, parPumpSuctionFacetoCoupling + parCouplingLength + 0.1 * parCouplingtoMotorEnd, 0
    oEnPoint.Set 0, parPumpSuctionFacetoCoupling + parCouplingLength + parCouplingtoMotorEnd, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dMotorDiameter
        
    'Insert your code for output 16(Shaft Coupling)
    'The shaft diameter is assumed to be 0.25 times the dMotorDiameter
    Dim dShaftDiameter As Double
    dShaftDiameter = 0.25 * dMotorDiameter
    oStPoint.Set 0, parPumpSuctionFacetoCoupling + parCouplingLength, 0
    oEnPoint.Set 0, parPumpSuctionFacetoCoupling, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dShaftDiameter
    
    'Insert your code for output 17(Shaft Motor)
    oStPoint.Set 0, parPumpSuctionFacetoCoupling + parCouplingLength + 0.1 * parCouplingtoMotorEnd, 0
    oEnPoint.Set 0, parPumpSuctionFacetoCoupling + parCouplingLength, 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter6
    
    'Insert your code for output 18(support under the motor)
    Dim ObjMotorSuppBox As Object
    oStPoint.Set -parBaseplateWidth * 0.4, parPumpSuctionFacetoCoupling + parCouplingLength + 0.25 * parCouplingtoMotorEnd, -parBaseplatetoSuction
    oEnPoint.Set parBaseplateWidth * 0.4, parPumpSuctionFacetoCoupling + parCouplingLength + 0.85 * parCouplingtoMotorEnd, -dMotorDiameter / 2

    Set ObjMotorSuppBox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjMotorSuppBox
    Set ObjMotorSuppBox = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Insert your code for output 19 to 30(supports under the Pump)
    'Support1
    Dim dSupportCenter As New AutoMath.DPosition
    dSupportCenter.Set -parBaseplateWidth / 2, 0.04 * parPumpSuctionFacetoCoupling, -parBaseplatetoSuction
    
    Dim oTopPoints(3) As IJDPosition
    Dim oBottomPoints(3) As IJDPosition
    
    For iCount = 0 To 3
        Set oTopPoints(iCount) = New DPosition
        Set oBottomPoints(iCount) = New DPosition
    Next iCount
    
    'Point 1
    oBottomPoints(0).x = dSupportCenter.x
    oBottomPoints(0).y = dSupportCenter.y
    oBottomPoints(0).z = dSupportCenter.z
    
    'Point 2
    oBottomPoints(1).x = dSupportCenter.x + parBaseplateWidth / 4
    oBottomPoints(1).y = dSupportCenter.y
    oBottomPoints(1).z = dSupportCenter.z
    
    'Point3
    oBottomPoints(2).x = dSupportCenter.x + parBaseplateWidth / 4
    oBottomPoints(2).y = dSupportCenter.y + 0.1 * parPumpSuctionFacetoCoupling
    oBottomPoints(2).z = dSupportCenter.z
    
    'Point 4
    oBottomPoints(3).x = dSupportCenter.x
    oBottomPoints(3).y = dSupportCenter.y + 0.1 * parPumpSuctionFacetoCoupling
    oBottomPoints(3).z = dSupportCenter.z
    
    oTopPoints(0).Set oBottomPoints(0).x, oBottomPoints(0).y - 0.01 * parPumpSuctionFacetoCoupling, oBottomPoints(0).z + parBaseplatetoSuction
    oTopPoints(1).Set oBottomPoints(1).x, oBottomPoints(1).y - 0.01 * parPumpSuctionFacetoCoupling, oBottomPoints(1).z + parBaseplatetoSuction
    oTopPoints(2).Set oBottomPoints(2).x, oBottomPoints(2).y - 0.01 * parPumpSuctionFacetoCoupling, oBottomPoints(2).z + parBaseplatetoSuction
    oTopPoints(3).Set oBottomPoints(3).x, oBottomPoints(3).y - 0.01 * parPumpSuctionFacetoCoupling, oBottomPoints(3).z + parBaseplatetoSuction
    
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopPoints, oBottomPoints)
    For iCount = 1 To 6
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount
    
    Set ObjCollection = Nothing
    Set dSupportCenter = Nothing
    
    'Support2
    dSupportCenter.Set parBaseplateWidth / 2, 0.04 * parPumpSuctionFacetoCoupling, -parBaseplatetoSuction
    
    'Point 1
    oBottomPoints(0).x = dSupportCenter.x
    oBottomPoints(0).y = dSupportCenter.y
    oBottomPoints(0).z = dSupportCenter.z
    
    'Point 2
    oBottomPoints(1).x = dSupportCenter.x - parBaseplateWidth / 4
    oBottomPoints(1).y = dSupportCenter.y
    oBottomPoints(1).z = dSupportCenter.z
    
    'Point3
    oBottomPoints(2).x = dSupportCenter.x - parBaseplateWidth / 4
    oBottomPoints(2).y = dSupportCenter.y + 0.1 * parPumpSuctionFacetoCoupling
    oBottomPoints(2).z = dSupportCenter.z
    
    'Point 4
    oBottomPoints(3).x = dSupportCenter.x
    oBottomPoints(3).y = dSupportCenter.y + 0.1 * parPumpSuctionFacetoCoupling
    oBottomPoints(3).z = dSupportCenter.z
    
    oTopPoints(0).Set oBottomPoints(0).x, oBottomPoints(0).y - 0.02 * parPumpSuctionFacetoCoupling, oBottomPoints(0).z + parBaseplatetoSuction
    oTopPoints(1).Set oBottomPoints(1).x, oBottomPoints(1).y - 0.02 * parPumpSuctionFacetoCoupling, oBottomPoints(1).z + parBaseplatetoSuction
    oTopPoints(2).Set oBottomPoints(2).x, oBottomPoints(2).y - 0.02 * parPumpSuctionFacetoCoupling, oBottomPoints(2).z + parBaseplatetoSuction
    oTopPoints(3).Set oBottomPoints(3).x, oBottomPoints(3).y - 0.02 * parPumpSuctionFacetoCoupling, oBottomPoints(3).z + parBaseplatetoSuction
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopPoints, oBottomPoints)
    
    For iCount = 1 To 6
         iOutput = iOutput + 1
         m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount
    
    Set ObjCollection = Nothing
    
    '==========================================
    'Construction of Equipment Foundatiopn Port 31
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
    dOrigin(1) = parBaseplateLength / 2
    dOrigin(2) = -parBaseplateBottomtoSuctionCL

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

    holes(0, 1) = -(parBaseplateLength / 2 - parXboltHole)
    holes(0, 2) = -(parBaseplateWidth / 2 - parYboltHole)
    holes(1, 1) = parBaseplateLength / 2 - parXboltHole
    holes(1, 2) = -(parBaseplateWidth / 2 - parYboltHole)
    holes(2, 1) = parBaseplateLength / 2 - parXboltHole
    holes(2, 2) = parBaseplateWidth / 2 - parYboltHole
    holes(3, 1) = -(parBaseplateLength / 2 - parXboltHole)
    holes(3, 2) = parBaseplateWidth / 2 - parYboltHole
                                                        
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
        
    Call ObjFoundationPort.SetHoles(holes)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Insert your code for output 32
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
     'Create line by two points:
     Dim LinePoints(0 To 23) As Double
     LinePoints(0) = -parBaseplateWidth / 2
     LinePoints(1) = 0
     LinePoints(2) = -parBaseplateBottomtoSuctionCL
     
     LinePoints(3) = parBaseplateWidth / 2
     LinePoints(4) = 0
     LinePoints(5) = -parBaseplateBottomtoSuctionCL
     
     LinePoints(6) = parBaseplateWidth / 2
     LinePoints(7) = parBaseplateLength
     LinePoints(8) = -parBaseplateBottomtoSuctionCL
     
     LinePoints(9) = -parBaseplateWidth / 2
     LinePoints(10) = parBaseplateLength
     LinePoints(11) = -parBaseplateBottomtoSuctionCL
     
     LinePoints(12) = -parBaseplateWidth / 2
     LinePoints(13) = 0
     LinePoints(14) = -parBaseplateBottomtoSuctionCL + parBaseplateHeight
     
     LinePoints(15) = parBaseplateWidth / 2
     LinePoints(16) = 0
     LinePoints(17) = -parBaseplateBottomtoSuctionCL + parBaseplateHeight
     
     LinePoints(18) = parBaseplateWidth / 2
     LinePoints(19) = parBaseplateLength
     LinePoints(20) = -parBaseplateBottomtoSuctionCL + parBaseplateHeight
     
     LinePoints(21) = -parBaseplateWidth / 2
     LinePoints(22) = parBaseplateLength
     LinePoints(23) = -parBaseplateBottomtoSuctionCL + parBaseplateHeight
     
     'Create the Edges and Points for Base Plate (Output 33)
    If CmpDblGreaterthan(parBaseplateHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBaseplateLength, LINEAR_TOLERANCE) _
                And CmpDblGreaterthan(parBaseplateWidth, LINEAR_TOLERANCE) Then

     Dim ObjEdgeColl As New Collection
        
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(0), LinePoints(1), LinePoints(2), _
                                        LinePoints(3), LinePoints(4), LinePoints(5))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(3), LinePoints(4), LinePoints(5), _
                                        LinePoints(6), LinePoints(7), LinePoints(8))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(6), LinePoints(7), LinePoints(8), _
                                        LinePoints(9), LinePoints(10), LinePoints(11))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(9), LinePoints(10), LinePoints(11), _
                                        LinePoints(0), LinePoints(1), LinePoints(2))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(0), LinePoints(1), LinePoints(2), _
                                        LinePoints(12), LinePoints(13), LinePoints(14))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(12), LinePoints(13), LinePoints(14), _
                                        LinePoints(21), LinePoints(22), LinePoints(23))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(21), LinePoints(22), LinePoints(23), _
                                        LinePoints(9), LinePoints(10), LinePoints(11))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(18), LinePoints(19), LinePoints(20), _
                                        LinePoints(6), LinePoints(7), LinePoints(8))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(18), LinePoints(19), LinePoints(20), _
                                        LinePoints(15), LinePoints(16), LinePoints(17))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(3), LinePoints(4), LinePoints(5), _
                                        LinePoints(15), LinePoints(16), LinePoints(17))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(15), LinePoints(16), LinePoints(17), _
                                        LinePoints(12), LinePoints(13), LinePoints(14))
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                        LinePoints(21), LinePoints(22), LinePoints(23), _
                                        LinePoints(18), LinePoints(19), LinePoints(20))
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To ObjEdgeColl.Count
            m_outputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
        Next iCount
        Set ObjEdgeColl = Nothing
        
        'Creating the Points on each Face of the Base Plate (Output 34)
        Dim ObjPointColl As New Collection
        
        'Points on Bottom and Top surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                              (LinePoints(0) + LinePoints(6)) / 2, (LinePoints(1) + LinePoints(7)) / 2, (LinePoints(2) + LinePoints(8)) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                 (LinePoints(12) + LinePoints(18)) / 2, (LinePoints(13) + LinePoints(19)) / 2, (LinePoints(14) + LinePoints(20)) / 2)
        'Points on Left and Right surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (LinePoints(0) + LinePoints(15)) / 2, (LinePoints(1) + LinePoints(16)) / 2, (LinePoints(2) + LinePoints(17)) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (LinePoints(6) + LinePoints(21)) / 2, (LinePoints(7) + LinePoints(22)) / 2, (LinePoints(8) + LinePoints(23)) / 2)
        'Points on Front and Back surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (LinePoints(0) + LinePoints(21)) / 2, (LinePoints(1) + LinePoints(22)) / 2, (LinePoints(2) + LinePoints(23)) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                               (LinePoints(3) + LinePoints(18)) / 2, (LinePoints(4) + LinePoints(19)) / 2, (LinePoints(5) + LinePoints(20)) / 2)
        'Set the output
        iOutput = iOutput + 1
        For iCount = 1 To ObjPointColl.Count
            m_outputColl.AddOutput "Points_", ObjPointColl(iCount)
        Next iCount
        Set ObjPointColl = Nothing
    End If
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

