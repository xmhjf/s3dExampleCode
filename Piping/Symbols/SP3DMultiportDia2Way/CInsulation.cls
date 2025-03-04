VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
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
'   07.Nov.2007     RUK     TR-128456: Provided a check on end points of non flanged insulation port cylinders in which case small cylinders of negligible thickness will be created
'   09-June-2008    MP      CR-141585  Multiport valve symbols need to be enhanced to address Gemu valve requirements. (Implemented part data basis: 459)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Terminate()
      Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(5)
    parFace2toCenter = arrayOfInputs(6)
    parFace3toCenter = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(12)
   
    Dim iOutput     As Double
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Dim ObjValveBody As Object
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    
    Set oPipeComponent = oPartFclt
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    'Checking for the PartDataBasis
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
        parElbowEndFacetoCenter = arrayOfInputs(8)
        parOffsetBetOutlets = arrayOfInputs(9)
        parOffset = arrayOfInputs(10)
        parBodyWidth = arrayOfInputs(11)
        
        Dim pipeDiam1       As Double
        Dim pipeDiam2       As Double
        Dim pipeDiam3       As Double
        Dim flangeThick     As Double
        Dim flangeDiam1     As Double
        Dim flangeDiam2     As Double
        Dim flangeDiam3     As Double
        Dim sptOffset1      As Double
        Dim depth1          As Double
        Dim sptOffset2      As Double
        Dim depth2          As Double
        Dim sptOffset3      As Double
        Dim depth3          As Double
        
        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset3, depth3
        
        'Determine the Insulation Diameter
        Dim dInsDia1 As Double
        Dim dInsDia2 As Double
        Dim dInsDia3 As Double
        
        dInsDia1 = pipeDiam1 + 2 * parInsulationThickness
        dInsDia2 = pipeDiam2 + 2 * parInsulationThickness
        dInsDia3 = pipeDiam3 + 2 * parInsulationThickness
        
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Dim oVector As AutoMath.DVector
        Dim oTransMat As AutoMath.DT4x4
        
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oVector = New DVector
        Set oTransMat = New DT4x4
        
        'Create the Valve Body
        Dim dLineStrPts(0 To 20) As Double
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
        
        oTransMat.LoadIdentity
        oTransMat.[Scale] ((parBodyWidth + 2 * parInsulationThickness) / parBodyWidth)
        
        oVector.Set -1, 0, 0
        Set ObjValveBody = PlaceProjection(m_OutputColl, oLineString, oVector, parBodyWidth, True)
        ObjValveBody.Transform oTransMat
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
        Set oLineString = Nothing
        
        'Create the Inlet Port Geometry
        Dim oTransVec As AutoMath.DVector
        Dim objColl As Collection
        Dim iCount As Integer
        Set oTransVec = New DVector
        
        If parInletPortGeometry = STRAIGHT_INLET Then
            Dim objInlet As Object
            oStPoint.Set parBodyWidth / 2, 0, 0
            oEnPoint.Set parFace1toCenter, 0, 0
            Set objInlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia1, True)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "InsInletPort_", objInlet
            Set objInlet = Nothing
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            Set objColl = New Collection
            oStPoint.Set parBodyWidth / 2, 0, 0
            Set objColl = CreateInsPortGeometry(m_OutputColl, parInletPortGeometry, oStPoint, dInsDia1, _
                            parElbowEndFacetoCenter, parFace1toCenter - parBodyWidth / 2)
            'Set the Output
            iOutput = iOutput + 1
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsInletPort_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        End If
        
        'Create the Outlet Port1 Geometry
        If parOutletPort1Geometry = STRAIGHT_OUTLET Then
            Dim objOutlet As Object
            oStPoint.Set -parBodyWidth / 2, parOffsetBetOutlets / 2, 0
            oEnPoint.Set -parFace2toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia2, True)
        
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "InsOutletPort1_", objOutlet
            Set objOutlet = Nothing
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            Set objColl = New Collection
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, parOffsetBetOutlets / 2, 0
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, dInsDia2, _
                                parElbowEndFacetoCenter, parFace2toCenter - parBodyWidth / 2, _
                                0, PI, 0, oTransVec)
            'Set the Output
            iOutput = iOutput + 1
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort1_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            Set objColl = New Collection
            oStPoint.Set -parBodyWidth / 2, parOffsetBetOutlets / 2, 0
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, _
                                dInsDia2, parOffset, parFace2toCenter - parBodyWidth / 2)
            'Set the Output
            iOutput = iOutput + 1
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort1_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        End If
        
        'Create the Outlet Port 2 Geometry
        If parOutletPort2Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, -parOffsetBetOutlets / 2, 0
            oEnPoint.Set -parFace3toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia3, True)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "InsOutletPort2_", objOutlet
            Set objOutlet = Nothing
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
            Set objColl = New Collection
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, -parOffsetBetOutlets / 2, 0
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort2Geometry, oStPoint, dInsDia3, _
                                parElbowEndFacetoCenter, parFace3toCenter - parBodyWidth / 2, _
                                PI, PI, 0, oTransVec)
            'Set the Output
            iOutput = iOutput + 1
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort2_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            Set objColl = New Collection
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, -parOffsetBetOutlets / 2, 0
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort2Geometry, _
                                oStPoint, dInsDia3, parOffset, _
                                parFace3toCenter - parBodyWidth / 2, PI, 0, 0, oTransVec)
            'Set the Output
            iOutput = iOutput + 1
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort2_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        End If
        
        Set oTransVec = Nothing
        
        'Create the Inlet Nozzle
        Dim objNozzle As Object
        Dim dInsThick As Double
        'Determine the Insulation Diameter1 for the Inlet Nozzle
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, _
                                                            sptOffset1, depth1
        If CmpDblGreaterthanOrEqualTo(pipeDiam1, flangeDiam1) Then
            dInsDia1 = pipeDiam1 + 2 * parInsulationThickness
        Else
            dInsDia1 = flangeDiam1 + 2 * parInsulationThickness
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
            dInsDia1 = pipeDiam1 + 2 * parInsulationThickness
            dInsThick = 0
        Else
            dInsThick = parInsulationThickness
        End If
        
        If parInletPortGeometry = STRAIGHT_INLET Then
            oStPoint.Set parFace1toCenter, 0, 0
            oEnPoint.Set oStPoint.x - flangeThick - dInsThick, 0, 0
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parFace1toCenter, parElbowEndFacetoCenter, 0
            oEnPoint.Set parFace1toCenter, parElbowEndFacetoCenter - flangeThick - dInsThick, 0
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia1, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet nozzle 1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset2, depth2
        If CmpDblGreaterthanOrEqualTo(pipeDiam2, flangeDiam2) Then
            dInsDia2 = pipeDiam2 + 2 * parInsulationThickness
        Else
            dInsDia2 = flangeDiam2 + 2 * parInsulationThickness
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
            dInsDia2 = pipeDiam2 + 2 * parInsulationThickness
            dInsThick = 0
        Else
            dInsThick = parInsulationThickness
        End If
        
        If parOutletPort1Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parFace2toCenter, parOffsetBetOutlets / 2, 0
            oEnPoint.Set -(parFace2toCenter - flangeThick - dInsThick), _
                                                        oStPoint.y, oStPoint.z
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace2toCenter, (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter), 0
            oEnPoint.Set oStPoint.x, (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                                - flangeThick - dInsThick), 0
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace2toCenter, (parOffsetBetOutlets / 2 + parOffset), 0
            oEnPoint.Set -(parFace2toCenter - flangeThick - dInsThick), _
                                                (parOffsetBetOutlets / 2 + parOffset), 0
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia2, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet nozzle 2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset3, depth3
        If CmpDblGreaterthanOrEqualTo(pipeDiam3, flangeDiam3) Then
            dInsDia3 = pipeDiam3 + 2 * parInsulationThickness
        Else
            dInsDia3 = flangeDiam3 + 2 * parInsulationThickness
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
            dInsDia3 = pipeDiam3 + 2 * parInsulationThickness
            dInsThick = 0
        Else
            dInsThick = parInsulationThickness
        End If
        
        If parOutletPort2Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parFace3toCenter, -parOffsetBetOutlets / 2, 0
            oEnPoint.Set -(parFace3toCenter - flangeThick - dInsThick), _
                                                        -parOffsetBetOutlets / 2, 0
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace3toCenter, _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter), 0
            oEnPoint.Set -parFace3toCenter, -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter - _
                                flangeThick - dInsThick), 0
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace3toCenter, -(parOffsetBetOutlets / 2 + parOffset), 0
            oEnPoint.Set -(parFace3toCenter - flangeThick - dInsThick), _
                                -(parOffsetBetOutlets / 2 + parOffset), 0
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia3, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Remove the references
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oVector = Nothing
        Set oTransMat = Nothing
        Set oGeomFactory = Nothing
       
   
    ElseIf lPartDataBasis = MULTI_PORT_OPTIONS_2WAY_GEMU Then
    
        Dim pipeDiam        As Double
        Dim flangeDiam      As Double
        Dim cptOffset       As Double
        Dim depth           As Double
        
        Dim parValveBodyWidth As Double
        Dim parValveBodyDepth As Double
        Dim parValveBodyHeight As Double
        
        'Inputs
        parValveBodyWidth = arrayOfInputs(13)
        parValveBodyDepth = arrayOfInputs(14)
        parValveBodyHeight = arrayOfInputs(15)
            
        iOutput = 0
        
        'Origin is taken at the geometry center of the valve.
     
        'Insert your code for output 7(Insulated Valve Body)
        Dim dStepHeight As Double
        Dim oDirVec As New AutoMath.DVector
            
        dStepHeight = 0.1 * parValveBodyHeight ''This is assuemd 10% of the total ValveBody Height
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
        Dim dValveBodyTopWidth As Double
        ''Width of the Body Top is assumed to be 1.5 times the Pipe diameter of Nozzle 1.
        dValveBodyTopWidth = 1.5 * pipeDiam
        
        dLineStrPts(0) = parInsulationThickness + parValveBodyWidth / 2
        dLineStrPts(1) = -parInsulationThickness - parValveBodyHeight / 2
        dLineStrPts(2) = parInsulationThickness + parValveBodyDepth / 2
        
        dLineStrPts(3) = parInsulationThickness + parValveBodyWidth / 2
        dLineStrPts(4) = dStepHeight - parValveBodyHeight / 2
        dLineStrPts(5) = dLineStrPts(2)
                
        dLineStrPts(6) = parInsulationThickness + dValveBodyTopWidth / 2
        dLineStrPts(7) = parInsulationThickness + parValveBodyHeight / 2
        dLineStrPts(8) = dLineStrPts(2)
        
        dLineStrPts(9) = -parInsulationThickness - dValveBodyTopWidth / 2
        dLineStrPts(10) = parInsulationThickness + parValveBodyHeight / 2
        dLineStrPts(11) = dLineStrPts(2)
        
        dLineStrPts(12) = -parInsulationThickness - parValveBodyWidth / 2
        dLineStrPts(13) = dStepHeight - parValveBodyHeight / 2
        dLineStrPts(14) = dLineStrPts(2)
        
        dLineStrPts(15) = -parInsulationThickness - parValveBodyWidth / 2
        dLineStrPts(16) = -parInsulationThickness - parValveBodyHeight / 2
        dLineStrPts(17) = dLineStrPts(2)
    
        dLineStrPts(18) = parInsulationThickness + parValveBodyWidth / 2
        dLineStrPts(19) = -parInsulationThickness - parValveBodyHeight / 2
        dLineStrPts(20) = parInsulationThickness + parValveBodyDepth / 2
        
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStrPts)
                   
        oDirVec.Set 0, 0, -1
                
        Set ObjValveBody = PlaceProjection(m_OutputColl, oLineString, oDirVec, parValveBodyDepth, True)
        
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
    
    '   Insert your code for output 8(Nozzle 1 Insulation)
        Dim dInsulationDia As Double
        Dim oStartPoint As New AutoMath.DPosition
        Dim oEndPoint As New AutoMath.DPosition
       
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
       
        If pipeDiam < flangeDiam Then
            dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam + 2 * parInsulationThickness
        End If
        
        oStartPoint.Set 0, parFace1toCenter - parValveBodyHeight / 2, 0
        oEndPoint.Set 0, parFace1toCenter, 0
       ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
    
    '   Insert your code for output 9(Nozzle 2 Insulation)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
       
        If pipeDiam < flangeDiam Then
            dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam + 2 * parInsulationThickness
        End If
        
        oStartPoint.Set parValveBodyWidth / 2 - 0.6 * pipeDiam, -parFace2toCenter + parValveBodyHeight / 2, 0
        oEndPoint.Set parValveBodyWidth / 2 - 0.6 * pipeDiam, -parFace2toCenter, 0
       ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
    
    
    '   Insert your code for output 4(Nozzle 3 with length)
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, cptOffset, depth
       
        If pipeDiam < flangeDiam Then
            dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam + 2 * parInsulationThickness
        End If
        oStartPoint.Set -parValveBodyWidth / 2 + 0.6 * pipeDiam, -parFace3toCenter + parValveBodyHeight / 2, 0
        oEndPoint.Set -parValveBodyWidth / 2 + 0.6 * pipeDiam, -parFace3toCenter, 0
       ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
        
        Set oLineString = Nothing
        Set oStartPoint = Nothing
        Set oEndPoint = Nothing
        Set oPartFclt = Nothing
    End If
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Function CreateInsPortGeometry(OutputColl As Object, ByVal PortGeom As Integer, ByVal oStartPoint As IJDPosition, _
            ByVal dDiamter As Double, ByVal dStartToEnd As Double, ByVal dHeight As Double, _
            Optional dRotAbtX As Double, Optional dRotAbtY As Double, Optional dRotAbtZ As Double, _
            Optional transVec As IJDVector) As Collection
    Const METHOD = "CreateInsPortGeometry"
    On Error GoTo ErrorHandler

    Dim objPort As Object
    Dim objCollection As New Collection
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
                                CircularCorner, 0, stnorm, ednorm, True)
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
                                CircularCorner, 0, stnorm, ednorm, True)
        Set oLineStr = Nothing
    End If
    For Each objPort In Surfset
        If Not objPort Is Nothing Then
            objCollection.Add objPort
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
    For iCount = 1 To objCollection.Count
        objCollection.Item(iCount).Transform oTransMat
    Next iCount
    
    If Not transVec Is Nothing Then
        oTransMat.LoadIdentity
        oTransMat.Translate transVec
        For iCount = 1 To objCollection.Count
            objCollection.Item(iCount).Transform oTransMat
        Next iCount
    End If
    
    Set CreateInsPortGeometry = objCollection

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
