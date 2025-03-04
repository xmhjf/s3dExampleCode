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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         RUK
'   Creation Date:  Thursday Oct 18 2007
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 5 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   27.Sep.07       RUK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    Dim pipeDiam5        As Double
    Dim pipeDiam6        As Double
    Dim flangeThick     As Double
    Dim flangeDiam      As Double
    Dim cptOffset1       As Double
    Dim depth1           As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    Dim cptOffset3       As Double
    Dim depth3           As Double
    Dim cptOffset4       As Double
    Dim depth4           As Double
    Dim cptOffset5       As Double
    Dim depth5           As Double
    Dim cptOffset6       As Double
    Dim depth6           As Double
    
    Dim iOutput     As Double
    
    Dim parValveBodyWidth As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parFace5toCenter As Double
    Dim parFace6toCenter As Double
    Dim parInletPortGeometry As Double
    Dim parOutletPort1Geometry As Double
    Dim parOutletPort2Geometry As Double
    Dim parOutletPort3Geometry As Double
    Dim parOutletPort4Geometry As Double
    Dim parOutletPort5Geometry As Double
    Dim parElbowEndFacetoCenter As Double
    Dim parOffsetBetOutlets As Double
    Dim parBodyWidth As Double
    Dim parOffset As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveBodyWidth = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    parFace3toCenter = arrayOfInputs(5)
    parFace4toCenter = arrayOfInputs(6)
    parFace5toCenter = arrayOfInputs(7)
    parFace6toCenter = arrayOfInputs(8)
    parInletPortGeometry = arrayOfInputs(9)
    parOutletPort1Geometry = arrayOfInputs(10)
    parOutletPort2Geometry = arrayOfInputs(11)
    parOutletPort3Geometry = arrayOfInputs(12)
    parOutletPort4Geometry = arrayOfInputs(13)
    parOutletPort5Geometry = arrayOfInputs(14)
    parElbowEndFacetoCenter = arrayOfInputs(15)
    parOffsetBetOutlets = arrayOfInputs(16)
    parBodyWidth = arrayOfInputs(17)
    parOffset = arrayOfInputs(18)
    parInsulationThickness = arrayOfInputs(19)
    
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    If lPartDataBasis <= 1 Or lPartDataBasis = MULTI_PORT_OPTIONS_5WAY Then '4-way diaphragm diverter valve with inlet and outlet port options
        
        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, cptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, cptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam, cptOffset4, depth4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick, flangeDiam, cptOffset5, depth5
        RetrieveParameters 6, oPartFclt, m_OutputColl, pipeDiam6, flangeThick, flangeDiam, cptOffset6, depth6
        
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Dim oVector As AutoMath.DVector
        Dim oTransMat As AutoMath.DT4x4
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oVector = New DVector
        Set oTransMat = New DT4x4
        
        'Create the Valve Body
        Dim objValveBody As Object
        Dim dLineStrPts(0 To 26) As Double
        Dim oColl As Collection
        Dim dSide As Double
        Dim dPolyCenToSide As Double 'Distance between Polygon center to Side
        
        dSide = parBodyWidth + 0.0254
        dPolyCenToSide = (dSide / 2) / (Tan(36 * PI / 180))
        
        If CmpDblLessThan(2 * dPolyCenToSide, parOffsetBetOutlets + _
            ReturnMax5(pipeDiam1, pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5) + 0.0254) Then
            dPolyCenToSide = (parOffsetBetOutlets + _
                    ReturnMax5(pipeDiam1, pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5) + 0.0254) / 2
            dSide = 2 * dPolyCenToSide * Tan(36 * PI / 180)
        End If
        
        Set oColl = New Collection
        Dim oLine As IngrGeom3D.Line3d
        Dim iCount As Integer
        oVector.Set 1, 0, 0
        For iCount = 1 To 5
            oStPoint.Set parBodyWidth / 2, -dSide / 2, -dPolyCenToSide
            oEnPoint.Set oStPoint.x, dSide / 2, -dPolyCenToSide
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oColl.Add RotateObject(oLine, oVector, (iCount - 1) * (72 * PI / 180))
            Set oLine = Nothing
        Next iCount
        Dim oComplexStr As Object
        oStPoint.Set parBodyWidth / 2, -dSide / 2, -dPolyCenToSide
        Set oComplexStr = PlaceTrCString(oStPoint, oColl)
        oVector.Set -1, 0, 0
        Set objValveBody = PlaceProjection(m_OutputColl, oComplexStr, oVector, parBodyWidth, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing
        Set oComplexStr = Nothing
        For iCount = 1 To oColl.Count
            oColl.Remove 1
        Next iCount
        Set oColl = Nothing
        
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
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInlet
        Set objInlet = Nothing
        
        'Create the Outlet Port1 Geometry
        Dim objOutlet As Object
        If parOutletPort1Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace2toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, pipeDiam2, _
                                parElbowEndFacetoCenter, parFace2toCenter - parBodyWidth / 2, _
                                -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, _
                                pipeDiam2, parOffset, parFace2toCenter - parBodyWidth / 2, _
                                -PI / 2, 0, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        End If
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
        Set objOutlet = Nothing
        
        'Create the Outlet Port 2 Geometry
        If parOutletPort2Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace3toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 2 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
        
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort2Geometry, oStPoint, pipeDiam3, _
                                parElbowEndFacetoCenter, parFace3toCenter - parBodyWidth / 2, _
                                -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 2 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort2Geometry, oStPoint, _
                                pipeDiam3, parOffset, parFace3toCenter - parBodyWidth / 2, _
                                -PI / 2, 0, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 2 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        End If

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
        Set objOutlet = Nothing

        'Create the Outlet Port 3 Geometry
        If parOutletPort3Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace4toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam4, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 3 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, pipeDiam4, _
                                parElbowEndFacetoCenter, parFace4toCenter - parBodyWidth / 2, _
                                -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 3 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, pipeDiam4, _
                                parOffset, parFace4toCenter - parBodyWidth / 2, _
                                -PI / 2, 0, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 3 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        End If

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
        Set objOutlet = Nothing

        'Create the Outlet Port 4 Geometry
        If parOutletPort4Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace5toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam5, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 4 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort4Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort4Geometry, oStPoint, pipeDiam5, _
                                parElbowEndFacetoCenter, parFace5toCenter - parBodyWidth / 2, _
                                 -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 4 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort4Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort4Geometry, oStPoint, pipeDiam5, _
                                parOffset, parFace5toCenter - parBodyWidth / 2, _
                                -PI / 2, 0, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 4 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        End If

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
        Set objOutlet = Nothing
        
        'Create the Outlet Port 5 Geometry
        If parOutletPort5Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace6toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam6, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 5 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort5Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort5Geometry, oStPoint, pipeDiam6, _
                                parElbowEndFacetoCenter, parFace6toCenter - parBodyWidth / 2, _
                                 -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 5 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort5Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort5Geometry, oStPoint, pipeDiam6, _
                                parOffset, parFace6toCenter - parBodyWidth / 2, _
                                -PI / 2, 0, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 5 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        End If

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
        Set objOutlet = Nothing
        Set oTransVec = Nothing
        
        'Remove the references
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oVector = Nothing
        Set oTransMat = Nothing
        Set oGeomFactory = Nothing
        
        'Create the Inlet Nozzle
        Dim objNozzle As GSCADNozzleEntities.IJDNozzle
        Dim oPlacementPoint As AutoMath.DPosition
        Dim oDirVec As AutoMath.DVector
        
        Set oPlacementPoint = New DPosition
        Set oDirVec = New DVector
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                        flangeDiam, cptOffset1, depth1
        If parInletPortGeometry = STRAIGHT_INLET Then
            oDirVec.Set 1, 0, 0
            oPlacementPoint.Set (parFace1toCenter + cptOffset1 - depth1), 0, 0
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, 1, 0
            oPlacementPoint.Set parFace1toCenter, (parElbowEndFacetoCenter + cptOffset1 - depth1), 0
        End If
        
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet Nozzle 1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, _
                                        flangeDiam, cptOffset2, depth2
        If parOutletPort1Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace2toCenter + cptOffset2 - depth2), _
                        (parOffsetBetOutlets / 2) * Sin(72 * PI / 180), _
                        -(parOffsetBetOutlets / 2) * Cos(72 * PI / 180)
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, Sin(72 * PI / 180), -Cos(72 * PI / 180)
            oPlacementPoint.Set -parFace2toCenter, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset2 - depth2) * Sin(72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset2 - depth2) * Cos(72 * PI / 180)
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace2toCenter + cptOffset2 - depth2), _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(72 * PI / 180), _
                            -(parOffsetBetOutlets / 2 + parOffset) * Cos(72 * PI / 180)
        End If
        
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Create the Outlet Nozzle 2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, _
                                        flangeDiam, cptOffset3, depth3
        If parOutletPort2Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace3toCenter + cptOffset3 - depth3), _
                                (parOffsetBetOutlets / 2) * Sin(2 * 72 * PI / 180), _
                                -(parOffsetBetOutlets / 2) * Cos(2 * 72 * PI / 180)
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, Sin(2 * 72 * PI / 180), -Cos(2 * 72 * PI / 180)
            oPlacementPoint.Set -parFace3toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset3 - depth3) * Sin(2 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset3 - depth3) * Cos(2 * 72 * PI / 180)
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace3toCenter + cptOffset3 - depth3), _
                        (parOffsetBetOutlets / 2 + parOffset) * Sin(2 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parOffset) * Cos(2 * 72 * PI / 180)
        End If
        
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet Nozzle 3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, _
                                        flangeDiam, cptOffset4, depth4
        If parOutletPort3Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace4toCenter + cptOffset4 - depth4), _
                                (parOffsetBetOutlets / 2) * Sin(3 * 72 * PI / 180), _
                                -(parOffsetBetOutlets / 2) * Cos(3 * 72 * PI / 180)
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, Sin(3 * 72 * PI / 180), -Cos(3 * 72 * PI / 180)
            oPlacementPoint.Set -parFace4toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset4 - depth4) * Sin(3 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset4 - depth4) * Cos(3 * 72 * PI / 180)
                    
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace4toCenter + cptOffset4 - depth4), _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(3 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parOffset) * Cos(3 * 72 * PI / 180)
        End If
        
        Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet Nozzle 4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick, _
                                        flangeDiam, cptOffset5, depth5
        If parOutletPort4Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace5toCenter + cptOffset5 - depth5), _
                                (parOffsetBetOutlets / 2) * Sin(4 * 72 * PI / 180), _
                                -(parOffsetBetOutlets / 2) * Cos(4 * 72 * PI / 180)
        ElseIf parOutletPort4Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, Sin(4 * 72 * PI / 180), -Cos(4 * 72 * PI / 180)
            oPlacementPoint.Set -parFace5toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset5 - depth5) * Sin(4 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset5 - depth5) * Cos(4 * 72 * PI / 180)
                    
        ElseIf parOutletPort4Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace5toCenter + cptOffset5 - depth5), _
                        (parOffsetBetOutlets / 2 + parOffset) * Sin(4 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parOffset) * Cos(4 * 72 * PI / 180)
        End If
        
        Set objNozzle = CreateNozzle(5, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet Nozzle 5
        RetrieveParameters 6, oPartFclt, m_OutputColl, pipeDiam6, flangeThick, _
                                        flangeDiam, cptOffset6, depth6
        If parOutletPort5Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace6toCenter + cptOffset6 - depth6), _
                            (parOffsetBetOutlets / 2) * Sin(5 * 72 * PI / 180), _
                            -(parOffsetBetOutlets / 2) * Cos(5 * 72 * PI / 180)
        ElseIf parOutletPort5Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, Sin(5 * 72 * PI / 180), -Cos(5 * 72 * PI / 180)
            oPlacementPoint.Set -parFace6toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset6 - depth6) * Sin(5 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + cptOffset6 - depth6) * Cos(5 * 72 * PI / 180)

        ElseIf parOutletPort5Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace6toCenter + cptOffset6 - depth6), _
                        (parOffsetBetOutlets / 2 + parOffset) * Sin(5 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parOffset) * Cos(5 * 72 * PI / 180)
        End If

        Set objNozzle = CreateNozzle(6, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)

        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        
        Set oDirVec = Nothing
        Set oPlacementPoint = Nothing
        
    
        'Insert code for Valve Operator Body
        Dim oOperatorPart As IJValveOperator
        Dim oMultiValCol As IJMultiPortValveAccCol
        Dim oOperatorOcc   As IJPartOcc
        Dim OpOrigin As IJDPosition

        Dim oDirX As IJDVector
        Dim oDirY As IJDVector
        Dim oDirZ As IJDVector

        Set oDirX = New DVector
        Set oDirY = New DVector
        Set oDirZ = New DVector

        Set oPipeComponent = oPartFclt
        On Error GoTo ErrorLabel

        If Not oPipeComponent Is Nothing Then
          oPipeComponent.GetValveOperatorsForPorts oMultiValCol
        End If

        Set OpOrigin = New DPosition
        OpOrigin.Set 0, 0, 0

        'Operator for port 2 (output 12)
        oDirX.Set 0, -Cos(72 * PI / 180), -Sin(72 * PI / 180)
        oDirY.Set 0, Sin(72 * PI / 180), -Cos(72 * PI / 180)
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
                                        "PositionerOffset", dPolyCenToSide, oIJDEditJDArgument

            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
        End If
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing

        'Operator for 3rd nozzle (output 13)
        oDirX.Set 0, -Cos(2 * 72 * PI / 180), -Sin(2 * 72 * PI / 180)
        oDirY.Set 0, Sin(2 * 72 * PI / 180), -Cos(2 * 72 * PI / 180)
        oDirZ.Set 1, 0, 0

        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            'Update Occurrence Attributes of Operator.
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

            UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", parBodyWidth, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                    "PositionerOffset", dPolyCenToSide, oIJDEditJDArgument

            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update

        End If

        'Operator for 4th nozzle (output 14)
        oDirX.Set 0, -Cos(3 * 72 * PI / 180), -Sin(3 * 72 * PI / 180)
        oDirY.Set 0, Sin(3 * 72 * PI / 180), -Cos(3 * 72 * PI / 180)
        oDirZ.Set 1, 0, 0


        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(4)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator3", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            'Update Occurrence Attributes of Operator.
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

            UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", parBodyWidth, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                    "PositionerOffset", dPolyCenToSide, oIJDEditJDArgument

            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update

        End If

        'Operator for 5th nozzle (output 15)
        oDirX.Set 0, -Cos(4 * 72 * PI / 180), -Sin(4 * 72 * PI / 180)
        oDirY.Set 0, Sin(4 * 72 * PI / 180), -Cos(4 * 72 * PI / 180)
        oDirZ.Set 1, 0, 0

        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(5)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator4", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            'Update Occurrence Attributes of Operator.
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

            UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", parBodyWidth, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                    "PositionerOffset", dPolyCenToSide, oIJDEditJDArgument

            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update

        End If

        'Operator for 6th nozzle (output 15)
        oDirX.Set 0, -Cos(5 * 72 * PI / 180), -Sin(5 * 72 * PI / 180)
        oDirY.Set 0, Sin(5 * 72 * PI / 180), -Cos(5 * 72 * PI / 180)
        oDirZ.Set 1, 0, 0

        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(6)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator5", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            'Update Occurrence Attributes of Operator.
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

            UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", parBodyWidth, oIJDEditJDArgument
            UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", _
                                    "PositionerOffset", dPolyCenToSide, oIJDEditJDArgument

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

'''<{(Complex string curve begin)}>
Public Function PlaceString(OutputColl As Object, ByVal startPosition As AutoMath.DPosition, _
                                        ByVal curves As Collection) As IngrGeom3D.ComplexString3d

''' This function creates transient (non-persistent) complex string curve
''' based on collection of curves (IJElements with StartPoint/EndPoints methods)
''' exapmle of use
'''    Dim lines           As Collection
'''    Dim oLine           As IngrGeom3D.Line3d
'''    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
'''    Dim objCStr         As IngrGeom3D.ComplexString3d
'''
'''    Set lines = New Collection
'''    Set oGeomFactory = New IngrGeom3D.GeometryFactory
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0.5, 0, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, 0.5, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, -0.5, 0, 0)
'''    lines.Add oLine
'''    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0, 0, 0, 0, -0.5, 0)
'''    lines.Add oLine
'''    stPoint.Set 0, 0, -1
'''    Set objCStr = PlaceTrCString(stPoint, lines)
'''    Set oLine = Nothing
'''    Dim iCount As Integer
'''    For iCount = 1 To lines.Count
'''        lines.Remove 1
'''    Next iCount
'''    Set lines = Nothing
'''    ......     use our complex string (e.g. for projection or revolution)
'''    Set objCStr = Nothing
    
    Const METHOD = "PlaceString:"
    On Error GoTo ErrorHandler
        
    Dim objCString  As IngrGeom3D.ComplexString3d
    Dim curPoint    As New AutoMath.DPosition
    Dim curve       As IngrGeom3D.IJCurve
    Dim trCurve     As IngrGeom3D.IJTransform
    Dim objCurve    As Object
    'Dim Elems       As New DynElements
    Dim vMove       As New AutoMath.DVector
    Dim vCompare    As New AutoMath.DVector
    Dim Elems As IJElements
    Set Elems = New JObjectCollection

    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim z2 As Double
    
    Set curPoint = startPosition
    For Each objCurve In curves
        Set curve = objCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        vCompare.Set x2 - x1, y2 - y1, z2 - z1
        If vCompare.Length < LINEAR_TOLERANCE Then
'            If vCompare.Length = 0 Then
'                MsgBox "Start and end points of a curve are the same"
'            Else
'                MsgBox "Start and end points of a curve are too close"
'            End If
            Exit For
        End If
        vMove.Set curPoint.x - x1, curPoint.y - y1, curPoint.z - z1
        Dim tForm   As New AutoMath.DT4x4
        tForm.Translate vMove
        Set trCurve = objCurve
        trCurve.Transform tForm
        Set tForm = Nothing
        Elems.Add trCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        curPoint.Set x2, y2, z2
    Next objCurve
    Set objCString = geomFactory.ComplexStrings3d.CreateByCurves( _
                                        OutputColl.ResourceManager, Elems)
                                                    
    Set PlaceString = objCString
    Set objCString = Nothing
    Set geomFactory = Nothing
    Elems.Clear
    Set Elems = Nothing

    Exit Function
    
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD

End Function
'''<{(Complex string curve end)}>
    
    
