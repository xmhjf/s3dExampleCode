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
'   Creation Date:  Thursday, April 12 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Twentytwo Outputs:
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
    Dim iOutput As Double
    Dim parPumpFacetoDischargeCL As Double
    Dim parPumpCouplingFacetoDischargeCL As Double
    Dim parSuctionCLtoDischargeCL As Double
    Dim parCouplingtoMotorEnd As Double
    Dim parPumpCLtoDischargeFace As Double
    Dim parPumpCLtoSuctionFace As Double
    Dim parPumpCLtoDischargeCL As Double
    Dim parPumpCLtoSuctionCL As Double
    Dim parImpellerDiameter As Double
    Dim parDistbetweenShaftEnds As Double
    Dim parBaseplateLength As Double
    Dim parBaseplateWidth As Double
    Dim parBaseplateHeight As Double
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
    parPumpFacetoDischargeCL = arrayOfInputs(2)                 'Y
    parPumpCouplingFacetoDischargeCL = arrayOfInputs(3)         'W
    parSuctionCLtoDischargeCL = arrayOfInputs(4)                'X
    parCouplingtoMotorEnd = arrayOfInputs(5)                    'CMAX
    parPumpCLtoDischargeFace = arrayOfInputs(6)                 'T
    parPumpCLtoSuctionFace = arrayOfInputs(7)                   'U
    parPumpCLtoDischargeCL = arrayOfInputs(8)                   'R
    parPumpCLtoSuctionCL = arrayOfInputs(9)                     'S
    parImpellerDiameter = arrayOfInputs(10)
    parBaseplateLength = arrayOfInputs(11)                      'HB
    parBaseplateWidth = arrayOfInputs(12)                       'HA
    parBaseplateHeight = arrayOfInputs(13)
    parBaseplateBottomtoPumpCL = arrayOfInputs(14)              'HD
    parFoundationPortStartPoint = arrayOfInputs(15)             'HR
    parDistbetweenShaftEnds = arrayOfInputs(16)                 'DBSE
    parXboltHole = arrayOfInputs(17)
    parYboltHole = arrayOfInputs(18)
    
    m_oGeomHelper.OutputCollection = m_outputColl
    Set Center = New DPosition
    Center.Set 0, 0, 0
    iOutput = 0
    
    'Insert your code for Motor Body(Output 1)
    oStPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + parCouplingtoMotorEnd, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parImpellerDiameter
     
    'Insert your code for Shaft Ends(Output 2)
    'The Small Cylinder diameter is assumed to be 0.15 times the MotorDiameter
    Dim dSmallCylinder As Double
    dSmallCylinder = 0.15 * parImpellerDiameter
    oStPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * (parPumpCouplingFacetoDischargeCL - 0.98 * parPumpFacetoDischargeCL), 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dSmallCylinder

    'Insert your code for Shaft Ends(Output 3)
    oStPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + (0.2 * parCouplingtoMotorEnd - 0.2 * 0.25 * parCouplingtoMotorEnd), 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dSmallCylinder
    
    'Insert your code for Shaft Ends(Output 4)
    oStPoint.Set 0.98 * parPumpFacetoDischargeCL + 0.8 * (parPumpCouplingFacetoDischargeCL - 0.98 * parPumpFacetoDischargeCL), 0, 0
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
   
    'Insert your code for Pump Cylinder1(Output 6)
    'The cylinder1 diameter is assumed to be 0.4 times the ImpellerDiameter
    Dim dCylinderDiameter1 As Double
    dCylinderDiameter1 = 0.5 * parImpellerDiameter
    oStPoint.Set -parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set -0.85 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter1
           
    'Insert your code for Pump Cylinder2(Output 7)
    Dim dCylinderDiameter2 As Double
    dCylinderDiameter2 = 0.3 * parImpellerDiameter
    oStPoint.Set -0.8 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set -0.65 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter2

    'Insert your code for Pump Cylinder1(Output 8)
    'The cylinder1 diameter is assumed to be 0.4 times the ImpellerDiameter
    oStPoint.Set 0.65 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.8 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter2
           
    'Insert your code for Pump Cylinder2(Output 9)
    oStPoint.Set 0.83 * parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set 0.98 * parPumpFacetoDischargeCL, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dCylinderDiameter1
    
    'Insert your code for CouplingShaft(Output 10)
    oStPoint.Set -parPumpFacetoDischargeCL, 0, 0
    oEnPoint.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd, 0, 0
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, 0.1 * parImpellerDiameter
    
    'Insert your code for Supports Under Pump (Output 11)
    Dim dlinepts(0 To 14) As Double
    Dim oLine As IngrGeom3D.Line3d
    Dim objCurve As Object
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim objCollection As Collection
    
    dlinepts(0) = -0.8 * parPumpFacetoDischargeCL
    dlinepts(1) = -0.4 * parBaseplateWidth
    dlinepts(2) = -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
    dlinepts(3) = -0.8 * parPumpFacetoDischargeCL
    dlinepts(4) = -0.4 * parBaseplateWidth
    dlinepts(5) = -0.5 * dCylinderDiameter2
    
    dlinepts(6) = -0.65 * parPumpFacetoDischargeCL
    dlinepts(7) = -0.4 * parBaseplateWidth
    dlinepts(8) = -0.5 * dCylinderDiameter2
    
    dlinepts(9) = -0.65 * parPumpFacetoDischargeCL
    dlinepts(10) = -0.4 * parBaseplateWidth
    dlinepts(11) = -0.5 * parImpellerDiameter
    
    dlinepts(12) = -0.45 * parPumpFacetoDischargeCL
    dlinepts(13) = -0.4 * parBaseplateWidth
    dlinepts(14) = -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
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
                                                          dlinepts(9), dlinepts(10), dlinepts(11))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                          dlinepts(12), dlinepts(13), dlinepts(14))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                          dlinepts(0), dlinepts(1), dlinepts(2))
    
    objCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New DVector
    oAxisVec.Set 0, 1, 0
    Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
    Set objCurve = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, 0.8 * parBaseplateWidth, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objCurve
    Set objCurve = Nothing
    Set oAxisVec = Nothing
    Dim Count As Integer
    For Count = 1 To objCollection.Count
        objCollection.Remove 1
        oComplexStr.RemoveCurve True
    Next Count
    Set oComplexStr = Nothing
    Set objCollection = Nothing
    
    'PumpSupport2 (Output 12)
    dlinepts(0) = 0.45 * parPumpFacetoDischargeCL
    dlinepts(1) = -0.4 * parBaseplateWidth
    dlinepts(2) = -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
    dlinepts(3) = 0.65 * parPumpFacetoDischargeCL
    dlinepts(4) = -0.4 * parBaseplateWidth
    dlinepts(5) = -0.5 * parImpellerDiameter
    
    dlinepts(6) = 0.65 * parPumpFacetoDischargeCL
    dlinepts(7) = -0.4 * parBaseplateWidth
    dlinepts(8) = -0.5 * dCylinderDiameter2
    
    dlinepts(9) = 0.8 * parPumpFacetoDischargeCL
    dlinepts(10) = -0.4 * parBaseplateWidth
    dlinepts(11) = -0.5 * dCylinderDiameter2

    dlinepts(12) = 0.8 * parPumpFacetoDischargeCL
    dlinepts(13) = -0.4 * parBaseplateWidth
    dlinepts(14) = -parBaseplateBottomtoPumpCL + parBaseplateHeight
    
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
                                                          dlinepts(9), dlinepts(10), dlinepts(11))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                          dlinepts(12), dlinepts(13), dlinepts(14))
    
    objCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(12), dlinepts(13), dlinepts(14), _
                                                          dlinepts(0), dlinepts(1), dlinepts(2))
    
    objCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oAxisVec = New DVector
    oAxisVec.Set 0, 1, 0
    Set oComplexStr = PlaceTrCString(oStPoint, objCollection)
    Set objCurve = PlaceProjection(m_outputColl, oComplexStr, oAxisVec, 0.8 * parBaseplateWidth, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objCurve
    Set objCurve = Nothing
    For Count = 1 To objCollection.Count
        objCollection.Remove 1
        oComplexStr.RemoveCurve True
    Next Count
    Set objCollection = Nothing
    Set oComplexStr = Nothing

    'Insert your code for support under the motor(Output 13)
    Dim ObjMotorSuppBox As Object
    Dim oMotorSupportCen As New AutoMath.DPosition
    oMotorSupportCen.Set parPumpCouplingFacetoDischargeCL + parDistbetweenShaftEnds + 0.2 * parCouplingtoMotorEnd, 0, -parBaseplateBottomtoPumpCL + parBaseplateHeight
    oStPoint.Set oMotorSupportCen.x, oMotorSupportCen.y - 0.4 * parImpellerDiameter, oMotorSupportCen.z
    oEnPoint.Set oMotorSupportCen.x + 0.8 * parCouplingtoMotorEnd, oMotorSupportCen.y + 0.4 * parImpellerDiameter, -0.5 * parImpellerDiameter
    Set ObjMotorSuppBox = PlaceBox(m_outputColl, oStPoint, oEnPoint)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjMotorSuppBox
    Set ObjMotorSuppBox = Nothing
    
    'Insert your code for Impeller Body(Output 14-15)
    Dim oEqpLeftArc As Object
    Dim oEllipticalbody As Object
    Set oEqpLeftArc = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, -0.6 * parPumpFacetoDischargeCL, 0, 0, _
                           1, 0, 0, 0.3 * parImpellerDiameter, 0, 0, 2)
    oAxisVec.Set 1, 0, 0
    Set oEllipticalbody = PlaceProjection(m_outputColl, oEqpLeftArc, oAxisVec, 0.6 * parPumpFacetoDischargeCL, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oEllipticalbody
    Set oEllipticalbody = Nothing
    Set oEqpLeftArc = Nothing
    
    Set oEqpLeftArc = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, 0, 0, 0, _
                           1, 0, 0, 0.3 * 0.75 * parImpellerDiameter, 0, 0, 2 * 0.75)
    oAxisVec.Set 1, 0, 0
    Set oEllipticalbody = PlaceProjection(m_outputColl, oEqpLeftArc, oAxisVec, 0.65 * parPumpFacetoDischargeCL, True)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oEllipticalbody
    Set oEllipticalbody = Nothing
    Set oEqpLeftArc = Nothing
    Set oAxisVec = Nothing
    
    'Insert your code for  Plate(Output 16)
    Dim ObjHorPlate As Object
    oStPoint.Set -0.65 * parPumpFacetoDischargeCL, -0.4 * parImpellerDiameter, -0.1 * 0.5 * parImpellerDiameter
    oEnPoint.Set 0.65 * parPumpFacetoDischargeCL, 0.4 * parImpellerDiameter, 0.1 * 0.5 * parImpellerDiameter
    Set ObjHorPlate = PlaceBox(m_outputColl, oStPoint, oEnPoint)
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjHorPlate
    Set ObjHorPlate = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Insert your code for Baseplate(Place Base)(Output 17)
    'Make Bottomplane as defaultsurface
    Dim dBaseCenterPoint As Double
    
    'Transformation from meters to inches
    Const UNITSTRANS = 0.0254
    dBaseCenterPoint = 15 * UNITSTRANS + parFoundationPortStartPoint
    oBaseCenter.Set -dBaseCenterPoint, 0, -parBaseplateBottomtoPumpCL

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
           
        'Create Default Surface(Output 18)
        Dim ObjDefSurface As Object
        Set ObjDefSurface = objCollection(2)
        'Set the Output
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
        Set ObjDefSurface = Nothing
        Set objCollection = Nothing
        Set oBaseCenter = Nothing
        
        'Create the Edges and Points for Base Plate(Output 19)
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
            
        'Creating the Points on each Face of the Base Plate (Output 20)
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
    Set oGeomFactory = Nothing

    '==========================================
    'Construction of Equipment Foundation Port(Output 21)
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
    
    'Insert your code for ControlPoint(Output 22)
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

