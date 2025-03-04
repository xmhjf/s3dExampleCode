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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Thursday, Sep 27 2007
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 5 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   27.Sep.07       RUK     Created
'   07.Nov.2007     RUK     TR-128456: Provided a check on end points of non flanged insulation port cylinders in which case small cylinders of negligible thickness will be created
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

    
    Dim pipeDiam1       As Double
    Dim pipeDiam2       As Double
    Dim pipeDiam3       As Double
    Dim pipeDiam4       As Double
    Dim pipeDiam5       As Double
    Dim pipeDiam6       As Double
    Dim flangeThick     As Double
    Dim flangeDiam1     As Double
    Dim flangeDiam2     As Double
    Dim flangeDiam3     As Double
    Dim flangeDiam4     As Double
    Dim flangeDiam5     As Double
    Dim flangeDiam6     As Double
    Dim sptOffset1      As Double
    Dim depth1          As Double
    Dim sptOffset2      As Double
    Dim depth2          As Double
    Dim sptOffset3      As Double
    Dim depth3          As Double
    Dim sptOffset4      As Double
    Dim depth4          As Double
    Dim sptOffset5      As Double
    Dim depth5          As Double
    Dim sptOffset6      As Double
    Dim depth6          As Double
    
    Dim iOutput     As Double
    
    
    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If lPartDataBasis <= 1 Or lPartDataBasis = MULTI_PORT_OPTIONS_5WAY Then
    
        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam4, sptOffset4, depth4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick, flangeDiam5, sptOffset5, depth5
        RetrieveParameters 6, oPartFclt, m_OutputColl, pipeDiam6, flangeThick, flangeDiam6, sptOffset6, depth6
        
        'Determine the Insulation Diameter
        Dim dInsDia1 As Double
        Dim dInsDia2 As Double
        Dim dInsDia3 As Double
        Dim dInsDia4 As Double
        Dim dInsDia5 As Double
        Dim dInsDia6 As Double
        
        dInsDia1 = pipeDiam1 + 2 * parInsulationThickness
        dInsDia2 = pipeDiam2 + 2 * parInsulationThickness
        dInsDia3 = pipeDiam3 + 2 * parInsulationThickness
        dInsDia4 = pipeDiam4 + 2 * parInsulationThickness
        dInsDia5 = pipeDiam5 + 2 * parInsulationThickness
        dInsDia6 = pipeDiam6 + 2 * parInsulationThickness
        
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
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
        
        oTransMat.LoadIdentity
        oTransMat.[Scale] ((parBodyWidth + 2 * parInsulationThickness) / parBodyWidth)
        
        oVector.Set -1, 0, 0
        Set objValveBody = PlaceProjection(m_OutputColl, oComplexStr, oVector, parBodyWidth, True)
        objValveBody.Transform oTransMat
        
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
            Set objInlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia1, True)
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            Set objInlet = CreatePortGeometry(m_OutputColl, parInletPortGeometry, oStPoint, dInsDia1, _
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
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia2, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort1Geometry, oStPoint, dInsDia2, _
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
                                dInsDia2, parOffset, parFace2toCenter - parBodyWidth / 2, _
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
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia3, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 2 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
        
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort2Geometry, oStPoint, dInsDia3, _
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
                                dInsDia3, parOffset, parFace3toCenter - parBodyWidth / 2, _
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
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia4, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 3 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, dInsDia4, _
                                parElbowEndFacetoCenter, parFace4toCenter - parBodyWidth / 2, _
                                -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 3 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, dInsDia4, _
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
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia5, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 4 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort4Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort4Geometry, oStPoint, dInsDia5, _
                                parElbowEndFacetoCenter, parFace5toCenter - parBodyWidth / 2, _
                                 -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 4 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort4Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort4Geometry, oStPoint, dInsDia5, _
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
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia6, True)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 5 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
            
        ElseIf parOutletPort5Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort5Geometry, oStPoint, dInsDia6, _
                                parElbowEndFacetoCenter, parFace6toCenter - parBodyWidth / 2, _
                                 -PI / 2, PI, 0, oTransVec)
            oTransMat.LoadIdentity
            oVector.Set 1, 0, 0
            oTransMat.Rotate 5 * 72 * PI / 180, oVector
            objOutlet.Transform oTransMat
        
        ElseIf parOutletPort5Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort5Geometry, oStPoint, dInsDia6, _
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
            oStPoint.Set -parFace2toCenter, (parOffsetBetOutlets / 2) * Sin(72 * PI / 180), _
                                    -(parOffsetBetOutlets / 2) * Cos(72 * PI / 180)
            oEnPoint.Set -(parFace2toCenter - flangeThick - dInsThick), _
                                                        oStPoint.y, oStPoint.z
        ElseIf parOutletPort1Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace2toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Sin(72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Cos(72 * PI / 180)
            oEnPoint.Set oStPoint.x, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Sin(72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Cos(72 * PI / 180)
        ElseIf parOutletPort1Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace2toCenter, _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(72 * PI / 180), _
                            -(parOffsetBetOutlets / 2 + parOffset) * Cos(72 * PI / 180)
            oEnPoint.Set -(parFace2toCenter - flangeThick - dInsThick), _
                                (parOffsetBetOutlets / 2 + parOffset) * Sin(72 * PI / 180), _
                                -(parOffsetBetOutlets / 2 + parOffset) * Cos(72 * PI / 180)
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
            oStPoint.Set -parFace3toCenter, (parOffsetBetOutlets / 2) * Sin(2 * 72 * PI / 180), _
                                    -(parOffsetBetOutlets / 2) * Cos(2 * 72 * PI / 180)
            oEnPoint.Set -(parFace3toCenter - flangeThick - dInsThick), _
                                                        oStPoint.y, oStPoint.z
        ElseIf parOutletPort2Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace3toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Sin(2 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Cos(2 * 72 * PI / 180)
            oEnPoint.Set oStPoint.x, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Sin(2 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Cos(2 * 72 * PI / 180)
        ElseIf parOutletPort2Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace3toCenter, _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(2 * 72 * PI / 180), _
                            -(parOffsetBetOutlets / 2 + parOffset) * Cos(2 * 72 * PI / 180)
            oEnPoint.Set -(parFace3toCenter - flangeThick - dInsThick), _
                                            oStPoint.y, oStPoint.z
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia3, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet nozzle 3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam4, sptOffset4, depth4
        If CmpDblGreaterthanOrEqualTo(pipeDiam4, flangeDiam4) Then
            dInsDia4 = pipeDiam4 + 2 * parInsulationThickness
        Else
            dInsDia4 = flangeDiam4 + 2 * parInsulationThickness
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
            dInsDia4 = pipeDiam4 + 2 * parInsulationThickness
            dInsThick = 0
        Else
            dInsThick = parInsulationThickness
        End If
        
        If parOutletPort3Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parFace4toCenter, (parOffsetBetOutlets / 2) * Sin(3 * 72 * PI / 180), _
                                    -(parOffsetBetOutlets / 2) * Cos(3 * 72 * PI / 180)
            oEnPoint.Set -(parFace4toCenter - flangeThick - dInsThick), _
                                                        oStPoint.y, oStPoint.z
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace4toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Sin(3 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Cos(3 * 72 * PI / 180)
            oEnPoint.Set oStPoint.x, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Sin(3 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Cos(3 * 72 * PI / 180)
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace4toCenter, _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(3 * 72 * PI / 180), _
                            -(parOffsetBetOutlets / 2 + parOffset) * Cos(3 * 72 * PI / 180)
            oEnPoint.Set -(parFace4toCenter - flangeThick - dInsThick), _
                                            oStPoint.y, oStPoint.z
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia4, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        
        'Create the Outlet nozzle 4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick, flangeDiam5, sptOffset5, depth5
        If CmpDblGreaterthanOrEqualTo(pipeDiam5, flangeDiam5) Then
            dInsDia5 = pipeDiam5 + 2 * parInsulationThickness
        Else
            dInsDia5 = flangeDiam5 + 2 * parInsulationThickness
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
            dInsDia5 = pipeDiam5 + 2 * parInsulationThickness
            dInsThick = 0
        Else
            dInsThick = parInsulationThickness
        End If
        
        If parOutletPort4Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parFace5toCenter, (parOffsetBetOutlets / 2) * Sin(4 * 72 * PI / 180), _
                                    -(parOffsetBetOutlets / 2) * Cos(4 * 72 * PI / 180)
            oEnPoint.Set -(parFace5toCenter - flangeThick - dInsThick), _
                                                        oStPoint.y, oStPoint.z
        ElseIf parOutletPort4Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace5toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Sin(4 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Cos(4 * 72 * PI / 180)
            oEnPoint.Set oStPoint.x, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Sin(4 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Cos(4 * 72 * PI / 180)
        ElseIf parOutletPort4Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace5toCenter, _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(4 * 72 * PI / 180), _
                            -(parOffsetBetOutlets / 2 + parOffset) * Cos(4 * 72 * PI / 180)
            oEnPoint.Set -(parFace5toCenter - flangeThick - dInsThick), _
                                            oStPoint.y, oStPoint.z
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia5, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the Outlet nozzle 5
        RetrieveParameters 6, oPartFclt, m_OutputColl, pipeDiam6, flangeThick, flangeDiam6, sptOffset6, depth6
        If CmpDblGreaterthanOrEqualTo(pipeDiam6, flangeDiam6) Then
            dInsDia6 = pipeDiam6 + 2 * parInsulationThickness
        Else
            dInsDia6 = flangeDiam6 + 2 * parInsulationThickness
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
            dInsDia6 = pipeDiam6 + 2 * parInsulationThickness
            dInsThick = 0
        Else
            dInsThick = parInsulationThickness
        End If
        
        If parOutletPort5Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parFace6toCenter, (parOffsetBetOutlets / 2) * Sin(5 * 72 * PI / 180), _
                                    -(parOffsetBetOutlets / 2) * Cos(5 * 72 * PI / 180)
            oEnPoint.Set -(parFace6toCenter - flangeThick - dInsThick), _
                                                        oStPoint.y, oStPoint.z
        ElseIf parOutletPort5Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace6toCenter, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Sin(5 * 72 * PI / 180), _
                    -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter) * Cos(5 * 72 * PI / 180)
            oEnPoint.Set oStPoint.x, _
                        (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Sin(5 * 72 * PI / 180), _
                        -(parOffsetBetOutlets / 2 + parElbowEndFacetoCenter _
                                    - flangeThick - dInsThick) * Cos(5 * 72 * PI / 180)
        ElseIf parOutletPort5Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace6toCenter, _
                            (parOffsetBetOutlets / 2 + parOffset) * Sin(5 * 72 * PI / 180), _
                            -(parOffsetBetOutlets / 2 + parOffset) * Cos(5 * 72 * PI / 180)
            oEnPoint.Set -(parFace6toCenter - flangeThick - dInsThick), _
                                            oStPoint.y, oStPoint.z
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia6, True)
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
    End If
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

