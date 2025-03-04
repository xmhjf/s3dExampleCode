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
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         VRK
'   Creation Date:  Tuesday, April 3 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Thirty Outputs:
'      In which Default Surface is created at the bottom of the Base Plate.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI As Double
Private m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim iOutput  As Double
    Dim parPumpFacetoDischargeCL As Double
    Dim parPumpCouplingFacetoDischargeCL As Double
    Dim parCouplingtoMotorEnd As Double
    Dim parPumpCLtoDischargeFace As Double
    Dim parPumpCLtoSuctionFace As Double
    Dim parPumpCLtoDischargeCL As Double
    Dim parPumpCLtoSuctionCL As Double
    Dim parImpellerDiameter As Double
    Dim parDistbetweenShaftEnds As Double
    Dim parBaseplateLength As Double
    Dim parBaseplateHeight As Double
    Dim parBaseplateWidth As Double
    Dim parBaseplateBottomtoPumpCL As Double
    Dim parFoundationPortStartPoint As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    Dim oBaseCenter As New AutoMath.DPosition
    Dim iCount As Integer
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPumpFacetoDischargeCL = arrayOfInputs(2)            'Y
    parPumpCouplingFacetoDischargeCL = arrayOfInputs(3)    'W
    parCouplingtoMotorEnd = arrayOfInputs(4)               'Cmax
    parPumpCLtoDischargeFace = arrayOfInputs(5)            'T
    parPumpCLtoSuctionFace = arrayOfInputs(6)              'U
    parPumpCLtoDischargeCL = arrayOfInputs(7)              'z
    parPumpCLtoSuctionCL = arrayOfInputs(8)                's
    parImpellerDiameter = arrayOfInputs(9)                 'D
    parDistbetweenShaftEnds = arrayOfInputs(10)            'HT
    parBaseplateLength = arrayOfInputs(11)                 'HB
    parBaseplateWidth = arrayOfInputs(12)                  'HA
    parBaseplateHeight = arrayOfInputs(13)
    parBaseplateBottomtoPumpCL = arrayOfInputs(14)         'HD
    parFoundationPortStartPoint = arrayOfInputs(15)             'HR
    parXboltHole = arrayOfInputs(16)                       'HE
    parYboltHole = arrayOfInputs(17)                       'HF
    
    m_oGeomHelper.OutputCollection = m_outputColl
    
    Set Center = New DPosition
    Center.Set 0, 0, 0
    iOutput = 0
    
    'Insert your code for Motor Body(Output 1)
    oStPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.23 * parCouplingtoMotorEnd, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + parCouplingtoMotorEnd, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parImpellerDiameter
     
    'Insert your code for Shaft Ends(Output 2)
    'The Small Cylinder diameter is assumed to be 0.15 times the MotorDiameter
    Dim dSmallCylinder As Double
    dSmallCylinder = 0.15 * parImpellerDiameter
    oStPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * (parPumpCouplingFacetoDischargeCL - 0.9 * parPumpFacetoDischargeCL), 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dSmallCylinder

    'Insert your code for Shaft Ends(Output 3)
    oStPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + (0.2 * parCouplingtoMotorEnd - 0.2 * 0.25 * parCouplingtoMotorEnd), 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dSmallCylinder
    
    'Insert your code for Shaft Ends(Output 4)
    oStPoint.Set 0.9 * parPumpFacetoDischargeCL + 0.8 * (parPumpCouplingFacetoDischargeCL - 0.9 * parPumpFacetoDischargeCL), 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dSmallCylinder
   
    'Insert your code for Coupling (Output 5)
    'The CouplingDiameter  is assumed to be 0.2 times the MotorDiameter
    Dim dCouplingDiameter As Double
    dCouplingDiameter = 0.2 * parImpellerDiameter
    oStPoint.Set parPumpCouplingFacetoDischargeCL, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCouplingDiameter
       
    'Insert your code for Impeller(Output 6)
    oStPoint.Set -0.1 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.1 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parImpellerDiameter
    
    'Insert your code for Pump Cylinder1(Output 7)
    'The cylinder1 diameter is assumed to be 0.6 times the ImpellerDiameter
    Dim dCylinderDiameter1 As Double
    dCylinderDiameter1 = 0.6 * parImpellerDiameter
    oStPoint.Set -0.1 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set -0.35 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter1
    
    'Insert your code for Pump Cylinder2(Output 8)
    'The cylinder 2 diameter is assumed to be 0.4 times the ImpellerDiameter
    Dim dCylinderDiameter2 As Double
    dCylinderDiameter2 = 0.4 * parImpellerDiameter
    oStPoint.Set -0.35 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set -0.55 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter2
    
    'Insert your code for Pump Cylinder3(Output 9)
    'The cylinder3 diameter is assumed to be 0.3 times the ImpellerDiameter
    Dim dCylinderDiameter3 As Double
    dCylinderDiameter3 = 0.3 * parImpellerDiameter
    oStPoint.Set -0.55 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set -0.7 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter3
    
    'Insert your code for Pump Cylinder4(Output 10)
    'The cylinder4 diameter is assumed to be 0.5 times the ImpellerDiameter
    Dim dCylinderDiameter4 As Double
    dCylinderDiameter4 = 0.5 * parImpellerDiameter
    oStPoint.Set -0.75 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set -parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter4
  
    'Insert your code for Pump Cylinder1(Output 11)
    oStPoint.Set 0.1 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.35 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter1
    
    'Insert your code for Pump Cylinder2(Output 12)
    oStPoint.Set 0.35 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.55 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter2
    
    'Insert your code for Pump Cylinder3(Output 13)
    oStPoint.Set 0.55 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.7 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter3
    
    'Insert your code for Pump Cylinder4(Output 14)
    oStPoint.Set 0.75 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.9 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter4

    'Insert your code for CouplingShaft(Output 15)
    oStPoint.Set -parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * 0.75 * parCouplingtoMotorEnd, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.1 * parImpellerDiameter
    
    Set oEnPoint = Nothing
    Set oStPoint = Nothing
    
    'Insert your code for Supports Under Pump (Output 16)
    Dim oPumpSupportCen As New AutoMath.DPosition
    oPumpSupportCen.Set -0.8 * parPumpFacetoDischargeCL, 0, -parBaseplateBottomtoPumpCL + parBaseplateHeight
    Dim oTopPoints(3) As IJDPosition
    Dim oBottomPoints(3) As IJDPosition
    Dim objCollection As Collection

    For iCount = 0 To 3
       Set oTopPoints(iCount) = New DPosition
       Set oBottomPoints(iCount) = New DPosition
    Next iCount
    
    'Point 1
    oBottomPoints(0).x = oPumpSupportCen.x
    oBottomPoints(0).y = oPumpSupportCen.y - 0.3 * parBaseplateWidth
    oBottomPoints(0).z = oPumpSupportCen.z
    
    'Point 2
    oBottomPoints(1).x = oPumpSupportCen.x
    oBottomPoints(1).y = oPumpSupportCen.y + 0.3 * parBaseplateWidth
    oBottomPoints(1).z = oPumpSupportCen.z
    
    'Point3
    oBottomPoints(2).x = oPumpSupportCen.x + 0.25 * parPumpFacetoDischargeCL
    oBottomPoints(2).y = oPumpSupportCen.y + 0.3 * parBaseplateWidth
    oBottomPoints(2).z = oPumpSupportCen.z
    
    'Point 4
    oBottomPoints(3).x = oPumpSupportCen.x + 0.25 * parPumpFacetoDischargeCL
    oBottomPoints(3).y = oPumpSupportCen.y - 0.3 * parBaseplateWidth
    oBottomPoints(3).z = oPumpSupportCen.z
    
    oTopPoints(0).Set oBottomPoints(0).x + 0.1 * parPumpFacetoDischargeCL, oBottomPoints(0).y, oBottomPoints(0).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    oTopPoints(1).Set oBottomPoints(1).x + 0.1 * parPumpFacetoDischargeCL, oBottomPoints(1).y, oBottomPoints(1).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    oTopPoints(2).Set oBottomPoints(2).x, oBottomPoints(2).y, oBottomPoints(2).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    oTopPoints(3).Set oBottomPoints(3).x, oBottomPoints(3).y, oBottomPoints(3).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    
    Set objCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopPoints, oBottomPoints)
    iOutput = iOutput + 1
    For iCount = 1 To objCollection.Count
            m_outputColl.AddOutput "PumpSupport1_", objCollection(iCount)
    Next iCount
    Set objCollection = Nothing
    Set oPumpSupportCen = Nothing
    
    For iCount = 0 To 3
       Set oTopPoints(iCount) = Nothing
       Set oBottomPoints(iCount) = Nothing
    Next iCount
    
    'PumpSupport2 (Output 17)
    oPumpSupportCen.Set 0, 0, -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
    For iCount = 0 To 3
       Set oTopPoints(iCount) = New DPosition
       Set oBottomPoints(iCount) = New DPosition
    Next iCount
    
    'Point 1
    oBottomPoints(0).x = oPumpSupportCen.x + 0.55 * parPumpFacetoDischargeCL
    oBottomPoints(0).y = oPumpSupportCen.y - 0.3 * parBaseplateWidth
    oBottomPoints(0).z = oPumpSupportCen.z
    
    'Point 2
    oBottomPoints(1).x = oPumpSupportCen.x + 0.55 * parPumpFacetoDischargeCL
    oBottomPoints(1).y = oPumpSupportCen.y + 0.3 * parBaseplateWidth
    oBottomPoints(1).z = oPumpSupportCen.z
    
    'Point3
    oBottomPoints(2).x = oPumpSupportCen.x + 0.8 * parPumpFacetoDischargeCL
    oBottomPoints(2).y = oPumpSupportCen.y + 0.3 * parBaseplateWidth
    oBottomPoints(2).z = oPumpSupportCen.z
    
    'Point 4
    oBottomPoints(3).x = oPumpSupportCen.x + 0.8 * parPumpFacetoDischargeCL
    oBottomPoints(3).y = oPumpSupportCen.y - 0.3 * parBaseplateWidth
    oBottomPoints(3).z = oPumpSupportCen.z
    
    oTopPoints(0).Set oBottomPoints(0).x, oBottomPoints(0).y, oBottomPoints(0).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    oTopPoints(1).Set oBottomPoints(1).x, oBottomPoints(1).y, oBottomPoints(1).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    oTopPoints(2).Set oBottomPoints(2).x - 0.1 * parPumpFacetoDischargeCL, oBottomPoints(2).y, oBottomPoints(2).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3
    oTopPoints(3).Set oBottomPoints(3).x - 0.1 * parPumpFacetoDischargeCL, oBottomPoints(3).y, oBottomPoints(3).z - parBaseplateHeight + parBaseplateBottomtoPumpCL - 0.5 * dCylinderDiameter3

    Set objCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopPoints, oBottomPoints)
    iOutput = iOutput + 1
    For iCount = 1 To objCollection.Count
            m_outputColl.AddOutput "PumpSupport2_", objCollection(iCount)
    Next iCount
    Set objCollection = Nothing
    Set oPumpSupportCen = Nothing

    For iCount = 0 To 3
       Set oTopPoints(iCount) = Nothing
       Set oBottomPoints(iCount) = Nothing
    Next iCount
    
    'Insert your code for supports under the motor
    'Support1 (Output 18)
    Dim oMotorSupportCen As New AutoMath.DPosition
    oMotorSupportCen.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.23 * parCouplingtoMotorEnd, 0, -parBaseplateBottomtoPumpCL + parBaseplateHeight
                   
    For iCount = 0 To 3
       Set oTopPoints(iCount) = New DPosition
       Set oBottomPoints(iCount) = New DPosition
    Next iCount
    
    'Point 1
    oBottomPoints(0).x = oMotorSupportCen.x
    oBottomPoints(0).y = oMotorSupportCen.y - 0.6 * parImpellerDiameter
    oBottomPoints(0).z = oMotorSupportCen.z
    
    'Point 2
    oBottomPoints(1).x = oMotorSupportCen.x
    oBottomPoints(1).y = oMotorSupportCen.y - 0.3 * parImpellerDiameter
    oBottomPoints(1).z = oMotorSupportCen.z
    
    'Point3
    oBottomPoints(2).x = oMotorSupportCen.x + 0.63 * parCouplingtoMotorEnd
    oBottomPoints(2).y = oMotorSupportCen.y - 0.3 * parImpellerDiameter
    oBottomPoints(2).z = oMotorSupportCen.z
    
    'Point 4
    oBottomPoints(3).x = oMotorSupportCen.x + 0.63 * parCouplingtoMotorEnd
    oBottomPoints(3).y = oMotorSupportCen.y - 0.6 * parImpellerDiameter
    oBottomPoints(3).z = oMotorSupportCen.z
    
    oTopPoints(0).Set oBottomPoints(0).x, oBottomPoints(0).y, oBottomPoints(0).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    oTopPoints(1).Set oBottomPoints(1).x, oBottomPoints(1).y - 0.1 * parImpellerDiameter, oBottomPoints(1).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    oTopPoints(2).Set oBottomPoints(2).x, oBottomPoints(2).y - 0.1 * parImpellerDiameter, oBottomPoints(2).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    oTopPoints(3).Set oBottomPoints(3).x, oBottomPoints(3).y, oBottomPoints(0).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    
    Set objCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopPoints, oBottomPoints)
    iOutput = iOutput + 1
    For iCount = 1 To objCollection.Count
            m_outputColl.AddOutput "MotorSupport1_", objCollection(iCount)
    Next iCount
    Set objCollection = Nothing
    
    For iCount = 0 To 3
       Set oTopPoints(iCount) = Nothing
       Set oBottomPoints(iCount) = Nothing
    Next iCount
    
    'Support2 (Output 19)
    For iCount = 0 To 3
       Set oTopPoints(iCount) = New DPosition
       Set oBottomPoints(iCount) = New DPosition
    Next iCount
    
    'Point 1
    oBottomPoints(0).x = oMotorSupportCen.x
    oBottomPoints(0).y = oMotorSupportCen.y + 0.3 * parImpellerDiameter
    oBottomPoints(0).z = oMotorSupportCen.z
    
    'Point 2
    oBottomPoints(1).x = oMotorSupportCen.x
    oBottomPoints(1).y = oMotorSupportCen.y + 0.6 * parImpellerDiameter
    oBottomPoints(1).z = oMotorSupportCen.z
    
    'Point3
    oBottomPoints(2).x = oMotorSupportCen.x + 0.63 * parCouplingtoMotorEnd
    oBottomPoints(2).y = oMotorSupportCen.y + 0.6 * parImpellerDiameter
    oBottomPoints(2).z = oMotorSupportCen.z
    
    'Point 4
    oBottomPoints(3).x = oMotorSupportCen.x + 0.63 * parCouplingtoMotorEnd
    oBottomPoints(3).y = oMotorSupportCen.y + 0.3 * parImpellerDiameter
    oBottomPoints(3).z = oMotorSupportCen.z
    
    oTopPoints(0).Set oBottomPoints(0).x, oBottomPoints(0).y + 0.1 * parImpellerDiameter, oBottomPoints(0).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    oTopPoints(1).Set oBottomPoints(1).x, oBottomPoints(1).y, oBottomPoints(0).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    oTopPoints(2).Set oBottomPoints(2).x, oBottomPoints(2).y, oBottomPoints(0).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    oTopPoints(3).Set oBottomPoints(3).x, oBottomPoints(3).y + 0.1 * parImpellerDiameter, oBottomPoints(0).z + parBaseplateBottomtoPumpCL - parBaseplateHeight - 0.5 * parImpellerDiameter
    
    Set objCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopPoints, oBottomPoints)
    iOutput = iOutput + 1
    For iCount = 1 To objCollection.Count
            m_outputColl.AddOutput "MotorSupport2_", objCollection(iCount)
    Next iCount
    Set objCollection = Nothing
    
    For iCount = 0 To 3
       Set oTopPoints(iCount) = Nothing
       Set oBottomPoints(iCount) = Nothing
    Next iCount
   
    'Insert your code for MotorSidePlate1(Output 20)
    Dim objMotorSidePlate As Object
    Dim oPlateCorner1 As New AutoMath.DPosition
    Dim oPlateCorner2 As New AutoMath.DPosition
    
    oPlateCorner1.Set oMotorSupportCen.x - 0.03 * parCouplingtoMotorEnd, oMotorSupportCen.y - 0.6 * parImpellerDiameter, oMotorSupportCen.z - 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    oPlateCorner2.Set oMotorSupportCen.x, oMotorSupportCen.y + 0.6 * parImpellerDiameter, oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    Set objMotorSidePlate = PlaceBox(m_outputColl, oPlateCorner1, oPlateCorner2)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objMotorSidePlate
    Set objMotorSidePlate = Nothing
    Set oPlateCorner1 = Nothing
    Set oPlateCorner2 = Nothing
     
    'Insert your code for MotorSidePlate2(Output 21)
    oPlateCorner1.Set oMotorSupportCen.x + 0.6 * parCouplingtoMotorEnd, oMotorSupportCen.y - 0.6 * parImpellerDiameter, oMotorSupportCen.z - 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    oPlateCorner2.Set oMotorSupportCen.x + 0.63 * parCouplingtoMotorEnd, oMotorSupportCen.y + 0.6 * parImpellerDiameter, oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    Set objMotorSidePlate = PlaceBox(m_outputColl, oPlateCorner1, oPlateCorner2)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objMotorSidePlate
    Set objMotorSidePlate = Nothing
    Set oPlateCorner1 = Nothing
    Set oPlateCorner2 = Nothing
    
    'Insert your code for MotorSidePlateTopDomeArc(Output 22)
    Dim objArc As Object
    Dim oBoxTopArc As IngrGeom3D.Arc3d
    Dim oLineArc As IngrGeom3D.Line3d
    Dim CenterX, CenterY, CenterZ As Double
    Dim StartX, StartY, StartZ As Double
    Dim EndX, EndY, EndZ As Double
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New DVector
    Set oStPoint = New DPosition
    Set objCollection = New Collection

    CenterX = parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd
    CenterY = 0
    CenterZ = oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight + 0.05
    
    StartX = parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd
    StartY = -0.6 * parImpellerDiameter
    StartZ = oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    
    EndX = parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd
    EndY = 0.6 * parImpellerDiameter
    EndZ = oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    
    Set oBoxTopArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                                                   StartX, StartY, StartZ, _
                                                                  CenterX, CenterY, CenterZ, EndX, EndY, EndZ)
    
    objCollection.Add oBoxTopArc
    Set oBoxTopArc = Nothing
    
    Set oLineArc = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                 EndX, EndY, EndZ, StartX, StartY, StartZ)
    
    objCollection.Add oLineArc
    Set oLineArc = Nothing
    
    oStPoint.Set StartX, StartY, StartZ
    Set oAxisVec = New DVector
    oAxisVec.Set 1, 0, 0
    Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
    Set objArc = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, 0.03 * parCouplingtoMotorEnd, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objArc
    Set objArc = Nothing
    Set oBoxTopArc = Nothing
    Set oAxisVec = Nothing
    Dim Count As Integer
    For Count = 1 To objCollection.Count
        objCollection.Remove 1
        oComplexStr.RemoveCurve True
    Next Count
    Set objCollection = Nothing
    
    'Insert your code for MotorSidePlateTopDomeArc(Output 23)
    CenterX = parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.83 * parCouplingtoMotorEnd
    CenterY = 0
    CenterZ = oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight + 0.05
    
    StartX = parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.83 * parCouplingtoMotorEnd
    StartY = -0.6 * parImpellerDiameter
    StartZ = oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    
    EndX = parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.83 * parCouplingtoMotorEnd
    EndY = 0.6 * parImpellerDiameter
    EndZ = oMotorSupportCen.z + 0.5 * parImpellerDiameter + parBaseplateBottomtoPumpCL - parBaseplateHeight
    
    Set objCollection = New Collection
    Set oBoxTopArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                                                   StartX, StartY, StartZ, _
                                                                  CenterX, CenterY, CenterZ, EndX, EndY, EndZ)
    
    objCollection.Add oBoxTopArc
    Set oBoxTopArc = Nothing
    
    Set oLineArc = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                 EndX, EndY, EndZ, StartX, StartY, StartZ)
    
    objCollection.Add oLineArc
    Set oLineArc = Nothing
    oStPoint.Set StartX, StartY, StartZ
    Set oAxisVec = New DVector
    oAxisVec.Set 1, 0, 0
    Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
    Set objArc = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, 0.03 * parCouplingtoMotorEnd, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objArc
    Set objArc = Nothing
    Set oBoxTopArc = Nothing
    Set oAxisVec = Nothing
    Set oMotorSupportCen = Nothing
    For Count = 1 To objCollection.Count
        objCollection.Remove 1
        oComplexStr.RemoveCurve True
    Next Count
    Set oComplexStr = Nothing
    
    'Insert your code for Baseplate(Place Base)(Output 24)
    'Make Bottomplane as defaultsurface
    Dim oBaseCenterPoint As Double
 
    'transformation from meters to inches
    Const INCHTRANS = 0.0254
    oBaseCenterPoint = 15 * INCHTRANS + parFoundationPortStartPoint
    oBaseCenter.Set -oBaseCenterPoint, 0, -parBaseplateBottomtoPumpCL

    Dim oTopSurPts(3) As IJDPosition
    Dim oBotSurPts(3) As IJDPosition
    If CmpDblGreaterthan(parBaseplateLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBaseplateWidth, LINEAR_TOLERANCE) Then
              
        For iCount = 0 To 3
             Set oTopSurPts(iCount) = New DPosition
             Set oBotSurPts(iCount) = New DPosition
        Next iCount
        
        'Point 1
        oBotSurPts(0).x = oBaseCenter.x
        oBotSurPts(0).y = oBaseCenter.y - parBaseplateWidth / 2
        oBotSurPts(0).z = oBaseCenter.z
        
        'Point 2
        oBotSurPts(1).x = oBaseCenter.x
        oBotSurPts(1).y = oBaseCenter.y + parBaseplateWidth / 2
        oBotSurPts(1).z = oBaseCenter.z
        
        'Point3
        oBotSurPts(2).x = oBaseCenter.x + parBaseplateLength
        oBotSurPts(2).y = oBaseCenter.y + parBaseplateWidth / 2
        oBotSurPts(2).z = oBaseCenter.z
        
        'Point 4
        oBotSurPts(3).x = oBaseCenter.x + parBaseplateLength
        oBotSurPts(3).y = oBaseCenter.y - parBaseplateWidth / 2
        oBotSurPts(3).z = oBaseCenter.z
        
        oTopSurPts(0).Set oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z + parBaseplateHeight
        oTopSurPts(1).Set oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z + parBaseplateHeight
        oTopSurPts(2).Set oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z + parBaseplateHeight
        oTopSurPts(3).Set oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z + parBaseplateHeight
        
        Set objCollection = PlaceTrapezoidWithPlanes(m_outputColl, oTopSurPts, oBotSurPts)
        'Set the Output
         iOutput = iOutput + 1
         For iCount = 1 To objCollection.Count
             m_outputColl.AddOutput "BasePlate_", objCollection(iCount)
         Next iCount
            
        'Create Default Surface(Output 25)
        Dim ObjDefSurface As Object
        Set ObjDefSurface = objCollection(2)
        'Set the Output
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
        Set ObjDefSurface = Nothing
        Set objCollection = Nothing
        Set oBaseCenter = Nothing
        
        'Create the Edges and Points for Base Plate(Output 26)
        Dim ObjEdgeColl As New Collection
            
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                    oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                    oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oTopSurPts(0).x, oTopSurPts(0).y, oTopSurPts(0).z, _
                                    oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                    oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                    oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oBotSurPts(3).x, oBotSurPts(3).y, oBotSurPts(3).z, _
                                    oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                    oBotSurPts(0).x, oBotSurPts(0).y, oBotSurPts(0).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                    oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oBotSurPts(1).x, oBotSurPts(1).y, oBotSurPts(1).z, _
                                    oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                    oTopSurPts(3).x, oTopSurPts(3).y, oTopSurPts(3).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                    oTopSurPts(1).x, oTopSurPts(1).y, oTopSurPts(1).z)
        ObjEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    oTopSurPts(2).x, oTopSurPts(2).y, oTopSurPts(2).z, _
                                    oBotSurPts(2).x, oBotSurPts(2).y, oBotSurPts(2).z)
        'Set the Output
        iOutput = iOutput + 1
        For iCount = 1 To ObjEdgeColl.Count
            m_outputColl.AddOutput "Edges_", ObjEdgeColl(iCount)
        Next iCount
        Set ObjEdgeColl = Nothing
            
        'Creating the Points on each Face of the Base Plate (Output 27)
        Dim ObjPointColl As New Collection
        
        'Points on Right and Left surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (oTopSurPts(1).x + oBotSurPts(2).x) / 2, (oTopSurPts(1).y + oBotSurPts(2).y) / 2, (oTopSurPts(1).z + oBotSurPts(2).z) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (oTopSurPts(0).x + oBotSurPts(3).x) / 2, (oTopSurPts(0).y + oBotSurPts(3).y) / 2, (oTopSurPts(0).z + oBotSurPts(3).z) / 2)
        'Points on Front and Back surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (oTopSurPts(3).x + oBotSurPts(2).x) / 2, (oTopSurPts(3).y + oBotSurPts(2).y) / 2, (oTopSurPts(3).z + oBotSurPts(2).z) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (oTopSurPts(0).x + oBotSurPts(1).x) / 2, (oTopSurPts(0).y + oBotSurPts(1).y) / 2, (oTopSurPts(0).z + oBotSurPts(1).z) / 2)
        'Points on Top and Bottom surfaces
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (oTopSurPts(0).x + oTopSurPts(2).x) / 2, (oTopSurPts(0).y + oTopSurPts(2).y) / 2, (oTopSurPts(0).z + oTopSurPts(2).z) / 2)
        ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, _
                                                (oBotSurPts(0).x + oBotSurPts(2).x) / 2, (oBotSurPts(0).y + oBotSurPts(2).y) / 2, (oBotSurPts(0).z + oBotSurPts(2).z) / 2)
        'Set the output
        iOutput = iOutput + 1
        For iCount = 1 To ObjPointColl.Count
            m_outputColl.AddOutput "Points_", ObjPointColl(iCount)
        Next iCount
        Set ObjPointColl = Nothing
    End If
    
    For iCount = 0 To 3
       Set oTopSurPts(iCount) = Nothing
       Set oBotSurPts(iCount) = Nothing
    Next iCount

    'Insert for Impellerbody(Output 28)
    Dim dlinepts(0 To 26) As Double
    Dim oLine As Object
    Dim oArc As Object
    Dim objCurve As Object
    Set oStPoint = New AutoMath.DPosition

    dlinepts(0) = -0.55 * parPumpFacetoDischargeCL
    dlinepts(1) = -0.2 * parBaseplateWidth
    dlinepts(2) = -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
    dlinepts(3) = -0.55 * parPumpFacetoDischargeCL
    dlinepts(4) = -0.2 * parBaseplateWidth
    dlinepts(5) = -0.5 * dCylinderDiameter2
    
    dlinepts(6) = -0.35 * parPumpFacetoDischargeCL
    dlinepts(7) = -0.2 * parBaseplateWidth
    dlinepts(8) = -0.5 * dCylinderDiameter2
    
    dlinepts(9) = 0
    dlinepts(10) = -0.2 * parBaseplateWidth
    dlinepts(11) = -(parBaseplateBottomtoPumpCL - parBaseplateHeight)
    
    dlinepts(12) = 0.35 * parPumpFacetoDischargeCL
    dlinepts(13) = -0.2 * parBaseplateWidth
    dlinepts(14) = -0.5 * dCylinderDiameter2
    
    dlinepts(15) = 0.55 * parPumpFacetoDischargeCL
    dlinepts(16) = -0.2 * parBaseplateWidth
    dlinepts(17) = -0.5 * dCylinderDiameter2

    dlinepts(18) = 0.55 * parPumpFacetoDischargeCL
    dlinepts(19) = -0.2 * parBaseplateWidth
    dlinepts(20) = -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
    dlinepts(21) = -0.35 * parPumpFacetoDischargeCL
    dlinepts(22) = -0.2 * parBaseplateWidth
    dlinepts(23) = -0.5 * dCylinderDiameter1

    dlinepts(24) = 0.35 * parPumpFacetoDischargeCL
    dlinepts(25) = -0.2 * parBaseplateWidth
    dlinepts(26) = -0.5 * dCylinderDiameter1

    Set objCollection = New Collection
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                          dlinepts(3), dlinepts(4), dlinepts(5))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                          dlinepts(6), dlinepts(7), dlinepts(8))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                          dlinepts(21), dlinepts(22), dlinepts(23))
    
    objCollection.Add oLine
    Set oLine = Nothing

    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(21), dlinepts(22), dlinepts(23), _
                                                            dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(24), dlinepts(25), dlinepts(26))
                                                            
    objCollection.Add oArc
    Set oArc = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(24), dlinepts(25), dlinepts(26), _
                                                          dlinepts(12), dlinepts(13), dlinepts(14))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                          dlinepts(15), dlinepts(16), dlinepts(17))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                          dlinepts(18), dlinepts(19), dlinepts(20))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(18), dlinepts(19), dlinepts(20), _
                                                          dlinepts(0), dlinepts(1), dlinepts(2))
    
    objCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oAxisVec = New DVector
    oAxisVec.Set 0, 1, 0
    Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
    Set objCurve = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, 0.4 * parBaseplateWidth, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objCurve
    Set objCurve = Nothing
    Set oGeomFactory = Nothing
    Set oAxisVec = Nothing
    Set oStPoint = Nothing
    For Count = 1 To objCollection.Count
        objCollection.Remove 1
        oComplexStr.RemoveCurve True
    Next Count
    Set objCollection = Nothing
    Set oComplexStr = Nothing

    '==========================================
    'Construction of Equipment Foundation Port(Output 29)
    '==========================================
    'Coordinate system is at the bottom of support.
    '
    '                      |------- |
    '   Y                  |        |
    '   ^                  |        |
    '   |                  |        |
    '   |                  |        |
    '   |                  |        |
    '   -----> X           |        |
    '   Symbol CS          |        |
    '                      |        |
    '                      |   X    |
    '                      |   ^    |
    '                      |   |    |
    '                      |   |    |
    '                      |---|--- |----> Y
    '                   Port CS
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = parYboltHole - parFoundationPortStartPoint
    dOrigin(1) = 0
    dOrigin(2) = -parBaseplateBottomtoPumpCL
    
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
    
    holes(0, 1) = -parXboltHole
    holes(0, 2) = -parYboltHole
    holes(1, 1) = parXboltHole
    holes(1, 2) = -parYboltHole
    holes(2, 1) = parXboltHole
    holes(2, 2) = parYboltHole
    holes(3, 1) = -parXboltHole
    holes(3, 2) = parYboltHole

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                           dXaxis(0), dXaxis(1), dXaxis(2), _
                           dZaxis(0), dZaxis(1), dZaxis(2))
    
    Call ObjFoundationPort.SetHoles(holes)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Insert your code for ControlPoint(Output 30)
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
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

