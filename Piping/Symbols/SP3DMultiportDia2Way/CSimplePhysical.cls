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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         RUK
'   Creation Date:  Thursday, Sep 27 2007
'   Description:
'       PDB: 460
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 2 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'       PDB: 459
'       This is a 2-way diverter valve symbol. This is prepared based on Gemu's catalog.
'       Source: Gemu  Multiport Valves M600-3-2C, Aseptic valve manifold machined from a single block.
'       M600 valve manifold designs, developed and produced according to customer requirements/specifications
'       The symbol has multiple operators. Each operator for each output port
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   27.Sep.07       RUK     CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols. (Implemented part data basis: Default, 460)
'   09-June-2008    MP      CR-141585  Multiport valve symbols need to be enhanced to address Gemu valve requirements. (Implemented part data basis: 459)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error m\essages
Private PI       As Double
Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
      Set m_oGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1       As Double
    Dim pipeDiam2       As Double
    Dim pipeDiam3       As Double
    Dim flangeThick     As Double
    Dim flangeDiam      As Double
    Dim sptOffset1      As Double
    Dim depth1          As Double
    Dim sptOffset2      As Double
    Dim depth2          As Double
    Dim sptOffset3      As Double
    Dim depth3          As Double
    
    Dim iOutput     As Double
    
'    Dim parInletPortGeometry As Double
'    Dim parOutletPort1Geometry As Double
'    Dim parOutletPort2Geometry As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
'    Dim parElbowEndFacetoCenter As Double
'    Dim parOffsetBetOutlets As Double
'    Dim parOffset As Double
'    Dim parBodyWidth As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parInletPortGeometry = arrayOfInputs(2)
'    parOutletPort1Geometry = arrayOfInputs(3)
'    parOutletPort2Geometry = arrayOfInputs(4)
'    parFace1toCenter = arrayOfInputs(5)
'    parFace2toCenter = arrayOfInputs(6)
'    parFace3toCenter = arrayOfInputs(7)
'    parElbowEndFacetoCenter = arrayOfInputs(8)
'    parOffsetBetOutlets = arrayOfInputs(9)
'    parOffset = arrayOfInputs(10)
'    parBodyWidth = arrayOfInputs(11)
'    parInsulationThickness = arrayOfInputs(12)

    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim oDirVec As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacementPoint As AutoMath.DPosition
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Dim oOperatorPart As IJValveOperator
    Dim oMultiValCol As IJMultiPortValveAccCol
    Dim oOperatorOcc   As IJPartOcc
    Dim OpOrigin As IJDPosition
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    
    Set oDirVec = New DVector
    Set oPlacementPoint = New DPosition
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector
    Set OpOrigin = New DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    
    If lPartDataBasis <= 1 Or lPartDataBasis = MULTI_PORT_OPTIONS_2WAY Then
    
        Dim parInletPortGeometry As Double
        Dim parOutletPort1Geometry As Double
        Dim parOutletPort2Geometry As Double
        Dim parElbowEndFacetoCenter As Double
        Dim parOffsetBetOutlets As Double
        Dim parOffset As Double
        Dim parBodyWidth As Double
        
        'Inputs
        parInletPortGeometry = arrayOfInputs(2)
        parOutletPort1Geometry = arrayOfInputs(3)
        parOutletPort2Geometry = arrayOfInputs(4)
        parFace1toCenter = arrayOfInputs(5)
        parFace2toCenter = arrayOfInputs(6)
        parFace3toCenter = arrayOfInputs(7)
        parElbowEndFacetoCenter = arrayOfInputs(8)
        parOffsetBetOutlets = arrayOfInputs(9)
        parOffset = arrayOfInputs(10)
        parBodyWidth = arrayOfInputs(11)
        parInsulationThickness = arrayOfInputs(12)
    
        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, sptOffset3, depth3
        
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Dim oVector As AutoMath.DVector
        Dim oTransMat As AutoMath.DT4x4
        
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oVector = New DVector
        Set oTransMat = New DT4x4
        
        'Create the Valve Body
        Dim ObjValveBody As Object
        Dim dLineStrPts(0 To 20) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        
        Dim dActualWidth As Double
        Dim dDistBetOutlets As Double
        Dim dHeight As Double
        
        dDistBetOutlets = parOffsetBetOutlets + pipeDiam2 / 2 + pipeDiam3 / 2 + 0.0254
        dActualWidth = 3 * dDistBetOutlets / 2
        dHeight = (dActualWidth / 2) * Tan(PI / 3)
        
        dLineStrPts(0) = parBodyWidth / 2
        dLineStrPts(1) = -((dActualWidth - parBodyWidth) / 2) * Cos(PI / 3)
        dLineStrPts(2) = -((2 * dHeight / 3) - ((dActualWidth - parBodyWidth) / 2) * Sin(PI / 3))
    
        dLineStrPts(3) = dLineStrPts(0)
        dLineStrPts(4) = -dLineStrPts(1)
        dLineStrPts(5) = dLineStrPts(2)
    
        dLineStrPts(6) = dLineStrPts(0)
        dLineStrPts(7) = dDistBetOutlets / 2 + (((dHeight / 3) / Sin(PI / 3)) _
                                        - ((dActualWidth - parBodyWidth) / 2)) * Cos(PI / 3)
        dLineStrPts(8) = (((dHeight / 3) / Sin(PI / 3)) _
                                        - ((dActualWidth - parBodyWidth) / 2)) * Sin(PI / 3)
    
        dLineStrPts(9) = dLineStrPts(0)
        dLineStrPts(10) = parBodyWidth / 2
        dLineStrPts(11) = dHeight / 3
    
        dLineStrPts(12) = dLineStrPts(0)
        dLineStrPts(13) = -dLineStrPts(10)
        dLineStrPts(14) = dLineStrPts(11)
    
        dLineStrPts(15) = dLineStrPts(0)
        dLineStrPts(16) = -dLineStrPts(7)
        dLineStrPts(17) = dLineStrPts(8)
    
        dLineStrPts(18) = dLineStrPts(0)
        dLineStrPts(19) = dLineStrPts(1)
        dLineStrPts(20) = dLineStrPts(2)
        
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPts)
        
        oVector.Set -1, 0, 0
        Set ObjValveBody = PlaceProjection(m_OutputColl, oLineString, oVector, parBodyWidth, True)
    
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
        Set oLineString = Nothing
        
        'Create the Inlet Port Geometry
        Dim objInlet As Object
        Dim oTransVec As AutoMath.DVector
        
        Set oTransVec = New DVector
        If parInletPortGeometry = STRAIGHT_INLET Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oEnPoint.Set parFace1toCenter, 0, 0
            Set objInlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam1, True)
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            Set objInlet = CreatePortGeometry(m_OutputColl, parInletPortGeometry, oStPoint, pipeDiam1, _
                            parElbowEndFacetoCenter, parFace1toCenter - parBodyWidth / 2)
        End If
        
        'Set the Output
        m_OutputColl.AddOutput "Inlet", objInlet
        Set objInlet = Nothing
        
        'Create the Outlet Port1 Geometry
        Dim objOutlet As Object
        If parOutletPort1Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, parOffsetBetOutlets / 2, 0
            oEnPoint.Set -parFace2toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, parOffsetBetOutlets / 2, 0
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, pipeDiam2, _
                                parElbowEndFacetoCenter, parFace2toCenter - parBodyWidth / 2, _
                                0, PI, 0, oTransVec)
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, parOffsetBetOutlets / 2, 0
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, _
                                pipeDiam2, parOffset, parFace2toCenter - parBodyWidth / 2)
        End If
        
        'Set the Output
        m_OutputColl.AddOutput "Outlet1", objOutlet
        Set objOutlet = Nothing
        
        'Create the Outlet Port 2 Geometry
        If parOutletPort2Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, -parOffsetBetOutlets / 2, 0
            oEnPoint.Set -parFace3toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, -parOffsetBetOutlets / 2, 0
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort2Geometry, oStPoint, pipeDiam3, _
                                parElbowEndFacetoCenter, parFace3toCenter - parBodyWidth / 2, _
                                PI, PI, 0, oTransVec)
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, -parOffsetBetOutlets / 2, 0
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort2Geometry, oStPoint, pipeDiam3, _
                                parOffset, parFace3toCenter - parBodyWidth / 2, PI, 0, 0, oTransVec)
        End If
        
        'Set the Output
        m_OutputColl.AddOutput "Outlet2", objOutlet
        Set objOutlet = Nothing
        Set oTransVec = Nothing
        
        'Remove the references
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oVector = Nothing
        Set oTransMat = Nothing
        Set oGeomFactory = Nothing
        
        'Create the Inlet Nozzle
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                        flangeDiam, sptOffset1, depth1
        If parInletPortGeometry = STRAIGHT_INLET Then
            oDirVec.Set 1, 0, 0
            oPlacementPoint.Set (parFace1toCenter + sptOffset1 - depth1), 0, 0
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, 1, 0
            oPlacementPoint.Set parFace1toCenter, (parElbowEndFacetoCenter + sptOffset1 - depth1), 0
        End If
        
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet Nozzle 1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, _
                                        flangeDiam, sptOffset2, depth2
        If parOutletPort1Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace2toCenter + sptOffset2 - depth2), _
                                parOffsetBetOutlets / 2, 0
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, 1, 0
            oPlacementPoint.Set -parFace2toCenter, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + sptOffset2 - depth2), _
                        0
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace2toCenter + sptOffset2 - depth2), _
                    (parOffsetBetOutlets / 2 + parOffset), 0
        End If
        
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Create the Outlet Nozzle 2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, _
                                        flangeDiam, sptOffset3, depth3
        If parOutletPort2Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace3toCenter + sptOffset3 - depth3), _
                                -parOffsetBetOutlets / 2, 0
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, -1, 0
            oPlacementPoint.Set -parFace3toCenter, _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + sptOffset3 - depth3), _
                        0
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace3toCenter + sptOffset3 - depth3), _
                        -(parOffsetBetOutlets / 2 + parOffset), 0
        End If
        
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        Set oDirVec = Nothing
        Set oPlacementPoint = Nothing
        
    
        'Insert your code for Valve Operator Body
        Set oPipeComponent = oPartFclt
        On Error GoTo ErrorLabel
    
        If Not oPipeComponent Is Nothing Then
          oPipeComponent.GetValveOperatorsForPorts oMultiValCol
        End If
    
        OpOrigin.Set 0, 0, 0
        
        'Operator for port 2 (output 12)
        oDirX.Set 0, -Cos(PI / 3), -Sin(PI / 3)
        oDirY.Set 0, Sin(PI / 3), -Cos(PI / 3)
        oDirZ.Set 1, 0, 0
        
        oDirX.Length = 1
        oDirY.Length = 1
        
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(2)
        Dim oIJDInputsArg As IJDInputsArg
        Dim oIJDEditJDArgument As IJDEditJDArgument
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator1", _
                                                oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
            'Update Occurrence Attributes of Operator.
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
        
            UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", parBodyWidth, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                        "PositionerOffset", dHeight / 3, oIJDEditJDArgument
        
            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
        End If
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
    
        'Operator for port 3 (output 13)
        oDirX.Set 0, Sin(PI / 6), -Cos(PI / 6)
        oDirY.Set 0, -Cos(PI / 6), -Sin(PI / 6)
        oDirZ.Set -1, 0, 0
        
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
                                            
            'Update Occurrence Attributes of Operator.
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
        
            UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", parBodyWidth, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                    "PositionerOffset", dHeight / 3, oIJDEditJDArgument
        
            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
    
        End If
    
        Set oDirX = Nothing
        Set oDirY = Nothing
        Set oDirZ = Nothing
        Set OpOrigin = Nothing
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
        Set oPipeComponent = Nothing
        Set oMultiValCol = Nothing
        
    ElseIf lPartDataBasis = MULTI_PORT_OPTIONS_2WAY_GEMU Then
        
        Dim pipeDiam        As Double
        Dim cptOffset       As Double
        Dim depth           As Double

        Dim parValveBodyWidth As Double
        Dim parValveBodyDepth As Double
        Dim parValveBodyHeight As Double
        
        'Inputs
        parValveBodyWidth = arrayOfInputs(13)
        parValveBodyDepth = arrayOfInputs(14)
        parValveBodyHeight = arrayOfInputs(15)
        parFace1toCenter = arrayOfInputs(5)
        parFace2toCenter = arrayOfInputs(6)
        parFace3toCenter = arrayOfInputs(7)
        parInsulationThickness = arrayOfInputs(12)
        
        iOutput = 0
        
        'Origin is taken at the geometry center of the valve.
        
        m_oGeomHelper.OutputCollection = m_OutputColl
        
        'Insert your code for output 1(Valve Body)
        Dim dStepHeight As Double
        dStepHeight = 0.1 * parValveBodyHeight ''This is assuemd 10% of the total ValveBody Height
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
        Dim dValveBodyTopWidth As Double
        'Width of the Body Top is assumed to be 1.5 times the Pipe diameter of Nozzle 1.
        dValveBodyTopWidth = 1.5 * pipeDiam
        
        dLineStrPts(0) = parValveBodyWidth / 2
        dLineStrPts(1) = -parValveBodyHeight / 2
        dLineStrPts(2) = parValveBodyDepth / 2
        
        dLineStrPts(3) = parValveBodyWidth / 2
        dLineStrPts(4) = dStepHeight - parValveBodyHeight / 2
        dLineStrPts(5) = parValveBodyDepth / 2
                
        dLineStrPts(6) = dValveBodyTopWidth / 2
        dLineStrPts(7) = parValveBodyHeight / 2
        dLineStrPts(8) = parValveBodyDepth / 2
        
        dLineStrPts(9) = -dValveBodyTopWidth / 2
        dLineStrPts(10) = parValveBodyHeight / 2
        dLineStrPts(11) = parValveBodyDepth / 2
        
        dLineStrPts(12) = -parValveBodyWidth / 2
        dLineStrPts(13) = dStepHeight - parValveBodyHeight / 2
        dLineStrPts(14) = parValveBodyDepth / 2
        
        dLineStrPts(15) = -parValveBodyWidth / 2
        dLineStrPts(16) = -parValveBodyHeight / 2
        dLineStrPts(17) = parValveBodyDepth / 2
        
        dLineStrPts(18) = parValveBodyWidth / 2
        dLineStrPts(19) = -parValveBodyHeight / 2
        dLineStrPts(20) = parValveBodyDepth / 2
              
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPts)
                
        oDirVec.Set 0, 0, -1
                
        Set ObjValveBody = PlaceProjection(m_OutputColl, oLineString, oDirVec, parValveBodyDepth, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
        Set oLineString = Nothing
    
        'Insert your code for output 2(Nozzle 1 with length)
        oPlacementPoint.Set 0, parFace1toCenter + cptOffset - depth, 0
        oDirVec.Set 0, 1, 0
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, parFace1toCenter - parValveBodyHeight / 2)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Insert your code for output 3(Nozzle 2 with length)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
                            
        oPlacementPoint.Set parValveBodyWidth / 2 - 0.6 * pipeDiam, _
                            -parFace2toCenter - cptOffset + depth, 0
        oDirVec.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, parFace2toCenter - parValveBodyHeight / 2)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Insert your code for output 4(Nozzle 3 with length)
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
                            
        oPlacementPoint.Set -parValveBodyWidth / 2 + 0.6 * pipeDiam, _
                            -parFace2toCenter - cptOffset + depth, 0
        oDirVec.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, parFace2toCenter - parValveBodyHeight / 2)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Insert your code for output 15(Valve Operator Body)
        Dim dOperatorAngle As Double
    
        dOperatorAngle = Atn((parValveBodyHeight - dStepHeight) / ((parValveBodyWidth - dValveBodyTopWidth) / 2))
    
        On Error GoTo ErrorLabel
    
        If Not oPipeComponent Is Nothing Then
          oPipeComponent.GetValveOperatorsForPorts oMultiValCol
        End If
    
        OpOrigin.Set 0, 0, 0
    
        'Operator for port 2
        oDirX.Set 0, 0, 1 ''Operator local X-axis matches with Valve Z-axis
        oDirY.Set Sin(dOperatorAngle), Cos(dOperatorAngle), 0
        oDirZ.Set -Cos(dOperatorAngle), Sin(dOperatorAngle), 0
    
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(2)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator1", _
                                    oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        
        End If
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
    
        'Operator for port 3
        oDirX.Set 0, 0, 1 ''Operator local X-axis matches with Valve Z-axis
        oDirY.Set -Sin(dOperatorAngle), Cos(dOperatorAngle), 0
        oDirZ.Set -Cos(dOperatorAngle), -Sin(dOperatorAngle), 0
        
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        
        End If
    
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
        Set oPartFclt = Nothing
        Set oPlacementPoint = Nothing
        Set oDirVec = Nothing
        Set oDirX = Nothing
        Set oDirY = Nothing
        Set oDirZ = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
       Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                            vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)
    
    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel
    
    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
'    Dim vNewValue As Variant
    Dim lInputIndex As Long
    
    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long
    
    ''To find the index number of the Attibute to be updated in the operator code.
    Set oSymbol = oOccurrence
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    lCount = oInputs.Count
    For lCount1 = 1 To lCount
        Set oInput = oInputs.Item(lCount1)
        If StrComp(oInput.name, strAttribNameOnInterface, vbTextCompare) = 0 Then
            lInputIndex = oInput.index
            Exit For
        End If
    Next lCount1
    
    ''If lInputIndex  is ZERO then operator doesnt have the attribute to be updated as it's InputParameter,
    ''so that no need for futher excecution.
    If lInputIndex = 0 Then Exit Sub
        
    ' Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
'    vNewValue = dInputParameter
    oIJDAttr.Value = vInputParameterValue
    
    ' Provide this value as input argument to the operator symbol
    UpdateSymbolArgument CDbl(vInputParameterValue), lInputIndex, oIJDEditJDArgument
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
   
End Sub

Private Sub UpdateSymbolArgument(dblValue As Double, lOperatorAttributeArgumentIndex As Long, _
                    oIJDEditJDArgument As IJDEditJDArgument)
                    
    Const METHOD = "UpdateSymbolArgument"
    On Error GoTo ErrorLabel
    
    Dim oPC As IJDParameterContent
    Dim oArgument As IJDArgument
                
    Set oPC = New DParameterContent
    oPC.Type = igValue
    oPC.UomType = 0
    oPC.uomValue = dblValue

    Set oArgument = New DArgument
    oArgument.index = lOperatorAttributeArgumentIndex
    oArgument.Entity = oPC
    
    oIJDEditJDArgument.SetArg oArgument
    
    Set oArgument = Nothing
    Set oPC = Nothing
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
        
End Sub

Private Function GetIJDEditJDArgument(oOperatorOcc As IJPartOcc) As IJDEditJDArgument

    Const METHOD = "GetIJDEditJDArgument"
    On Error GoTo ErrorLabel

    Dim oIJDEditJDArgument As IJDEditJDArgument
    Dim oIEnumJDArgument As IEnumJDArgument
    Dim oSymbol As IJDSymbol
    Dim oIJDValuesArg As IJDValuesArg
    
    Set oSymbol = oOperatorOcc
    Set oIJDValuesArg = oSymbol.IJDValuesArg
    Set oIEnumJDArgument = oIJDValuesArg.GetValues(igINPUT_ARGUMENTS_SET)
    Set oIJDEditJDArgument = oIEnumJDArgument
    Set GetIJDEditJDArgument = oIJDEditJDArgument
    
    Exit Function
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Function

Private Function CreatePortGeometry(OutputColl As Object, ByVal PortGeom As Integer, ByVal oStartPoint As IJDPosition, _
            ByVal dDiamter As Double, ByVal dStartToEnd As Double, ByVal dHeight As Double, _
            Optional dRotAbtX As Double, Optional dRotAbtY As Double, Optional dRotAbtZ As Double, _
            Optional transVec As IJDVector) As Object
    Const METHOD = "CreatePortGeometry"
    On Error GoTo ErrorHandler

    Dim objPort As Object
    Dim oGeomFact As IngrGeom3D.GeometryFactory
    Dim oCenter As AutoMath.DPosition
    Dim oNormal As AutoMath.DVector
    Dim oTransMat As AutoMath.DT4x4
    Dim oCircle As IngrGeom3D.Circle3d
    
    Set oGeomFact = New GeometryFactory
    Set oCenter = New DPosition
    Set oNormal = New DVector
    Set oTransMat = New DT4x4
    
    Dim Surfset As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim iCount As Integer
    
    If PortGeom = INLET_WITH_90DEG_ELBOW Or PortGeom = OUTLET_WITH_90DEG_ELBOW Then
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Dim oTraceStr As IngrGeom3D.ComplexString3d
        Dim oCollection As Collection
        Dim oLine As IngrGeom3D.Line3d
        Dim oArc As IngrGeom3D.Arc3d
        
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oTraceStr = New ComplexString3d
        Set oCollection = New Collection
        Set oLine = New Line3d
        Set oArc = New Arc3d
        
        oCenter.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        oNormal.Set 1, 0, 0
        Set oCircle = oGeomFact.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oCenter.x, oCenter.y, oCenter.z, _
                                                oNormal.x, oNormal.y, oNormal.z, dDiamter / 2)
        
        
        oStPoint.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        oEnPoint.Set oStPoint.x + dHeight - 0.2 * dStartToEnd, oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        oCenter.Set oEnPoint.x, oStartPoint.y + 0.2 * dStartToEnd, oStartPoint.z
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStartPoint.x + dHeight, oCenter.y, oCenter.z
        
        Set oArc = oGeomFact.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                oCenter.x, oCenter.y, oCenter.z, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oArc
        
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x, oStartPoint.y + dStartToEnd, oEnPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        oStPoint.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
        
        Set Surfset = oGeomFact.GeometryServices.CreateBySingleSweep( _
                                OutputColl.ResourceManager, oTraceStr, oCircle, _
                                CircularCorner, 0, stnorm, ednorm, False)
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        Set oCollection = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oArc = Nothing
        Set oLine = Nothing
        Set oTraceStr = Nothing
    ElseIf PortGeom = OUTLET_WITH_OFFSET Then
        oCenter.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        oNormal.Set -1, 0, 0
        Set oCircle = oGeomFact.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oCenter.x, oCenter.y, oCenter.z, _
                                                oNormal.x, oNormal.y, oNormal.z, dDiamter / 2)
        Dim oLineStr As IngrGeom3D.LineString3d
        Dim dPoints(0 To 11) As Double
        
        dPoints(0) = oStartPoint.x
        dPoints(1) = oStartPoint.y
        dPoints(2) = oStartPoint.z
        
        dPoints(3) = oStartPoint.x - dHeight / 3
        dPoints(4) = dPoints(1)
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dPoints(3) - dHeight / 3
        dPoints(7) = oStartPoint.y + dStartToEnd
        dPoints(8) = dPoints(2)
        
        dPoints(9) = dPoints(6) - dHeight / 3
        dPoints(10) = oStartPoint.y + dStartToEnd
        dPoints(11) = dPoints(2)
        
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 4, dPoints)
        
        Set Surfset = oGeomFact.GeometryServices.CreateBySingleSweep( _
                                OutputColl.ResourceManager, oLineStr, oCircle, _
                                CircularCorner, 0, stnorm, ednorm, False)
        Set oLineStr = Nothing
    End If
    For Each objPort In Surfset
        If Not objPort Is Nothing Then
            Exit For
        End If
    Next objPort
    
    oTransMat.LoadIdentity
    If Not CmpDblEqual(dRotAbtX, LINEAR_TOLERANCE) Then
        oNormal.Set 1, 0, 0
        oTransMat.Rotate dRotAbtX, oNormal
    End If
    If Not CmpDblEqual(dRotAbtY, LINEAR_TOLERANCE) Then
        oNormal.Set 0, 1, 0
        oTransMat.Rotate dRotAbtY, oNormal
    End If
    If Not CmpDblEqual(dRotAbtZ, LINEAR_TOLERANCE) Then
        oNormal.Set 0, 0, 1
        oTransMat.Rotate dRotAbtZ, oNormal
    End If
    objPort.Transform oTransMat
    
    If Not transVec Is Nothing Then
        oTransMat.LoadIdentity
        oTransMat.Translate transVec
        objPort.Transform oTransMat
    End If
    
    Set CreatePortGeometry = objPort

    'Remove the References
    For iCount = 1 To Surfset.Count
        Surfset.Remove 1
    Next iCount
    Set Surfset = Nothing
    Set oCenter = Nothing
    Set oNormal = Nothing
    Set oCircle = Nothing
    Set oTransMat = Nothing
    Set objPort = Nothing
    Set oGeomFact = Nothing
    
    Exit Function
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function


