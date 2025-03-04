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
'   Author:         svsmylav
'   Creation Date:  Thursday, July 22, 2004
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is 72pdf. PDS symbol MC=VS3WD.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006      KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'  17.Oct.2007      RUK     CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'                           Added code for the new PDB value 461.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Terminate()
      Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub
Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    ReturnMax3 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    Dim flangeThick     As Double
    Dim flangeDiam      As Double
    Dim sptOffset1       As Double
    Dim depth1           As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    Dim sptOffset3       As Double
    Dim depth3           As Double
    Dim sptOffset4       As Double
    Dim depth4           As Double
    
    Dim iOutput     As Double
    
    Dim parValveBodyHeight As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parNozzleCentertoCenter As Double
    Dim parOffset As Double
    Dim parInletPortGeometry As Double
    Dim parOutletPort1Geometry As Double
    Dim parOutletPort2Geometry As Double
    Dim parOutletPort3Geometry As Double
    Dim parElbowEndFacetoCenter As Double
    Dim parOffsetBetOutlets As Double
    Dim parBodyWidth As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveBodyHeight = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    parFace3toCenter = arrayOfInputs(5)
    parFace4toCenter = arrayOfInputs(6)
    parNozzleCentertoCenter = arrayOfInputs(7)
    parOffset = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oLineString As IngrGeom3D.LineString3d
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacementPoint As AutoMath.DPosition
    Dim oDirVec As AutoMath.DVector
    Dim oOperatorPart As IJValveOperator
    Dim oMultiValCol As IJMultiPortValveAccCol
    Dim oOperatorOcc   As IJPartOcc
    Dim OpOrigin As IJDPosition
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    
    Dim objValveBody As Object
    
    If lPartDataBasis <= 1 Then 'Default
    
        Dim dNozzle2FacetoBodyFace  As Double
        Dim dNozzle3FacetoBodyFace  As Double
        Dim dNozzle4FacetoBodyFace  As Double
        
        dNozzle2FacetoBodyFace = parFace2toCenter - parValveBodyHeight / 2
        dNozzle3FacetoBodyFace = parFace3toCenter - parValveBodyHeight / 2
        dNozzle4FacetoBodyFace = parFace4toCenter - parValveBodyHeight / 2
        
    '   Insert your code for output 1(Valve Body)
        Dim ValveBodyWidth As Double
    
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam, sptOffset4, depth4
        
        Dim cirRadius As Double
        Dim cirOffesetRadius As Double     'Radius of circle after considering nozzle offest
    ' CircumRadius of inner circle is parPortOffset/Sqr(3)
        cirRadius = parNozzleCentertoCenter / Sqr(3)
        cirOffesetRadius = cirRadius + parOffset
    
    ' Determine the largest pipe diameter of port 2,3 and 4
        Dim maxpipeDia As Double
        maxpipeDia = ReturnMax3(pipeDiam2, pipeDiam3, pipeDiam4)
    
    '   The half of side of triangle to corner is stored in Dim_5
        Dim Dim_5 As Double
    '   As per PDS Eden Dim_5 is calculated using cirRadius and maxpipeDia as follows.
    '   Compute Center of triangle to face
        Dim CenterOfTriangleToFace As Double
        CenterOfTriangleToFace = cirRadius + maxpipeDia / 2 + 0.25 * 0.0254 'Eden code uses 0.25 inch gap; 0.0254 meter/inch
    
        Dim_5 = CenterOfTriangleToFace / Tan(PI / 6)
        ValveBodyWidth = 2 * Dim_5
        
    'Create a line string and project it
        Dim LineStrPts(0 To 11)  As Double
    
    'Point 1
        LineStrPts(0) = -Dim_5
        LineStrPts(1) = -parValveBodyHeight / 2
        LineStrPts(2) = -CenterOfTriangleToFace
    
    'Point 2
        LineStrPts(3) = Dim_5
        LineStrPts(4) = LineStrPts(1)
        LineStrPts(5) = -CenterOfTriangleToFace
    
    'Point 3
        LineStrPts(6) = 0
        LineStrPts(7) = LineStrPts(1)
        LineStrPts(8) = -CenterOfTriangleToFace + ValveBodyWidth * Sin(PI / 3)
    
    'Point 4
        LineStrPts(9) = LineStrPts(0)
        LineStrPts(10) = LineStrPts(1)
        LineStrPts(11) = LineStrPts(2)
        
        Dim axisVect    As New AutoMath.DVector
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = New LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, LineStrPts)
         
        axisVect.Set 0, 1, 0
        Set objValveBody = PlaceProjection(m_OutputColl, oLineString, axisVect, parValveBodyHeight, True)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing
        Set axisVect = Nothing
    
        Dim sp3dElem1 As IJDObject
        Set sp3dElem1 = oLineString
        sp3dElem1.Remove
        Set oLineString = Nothing
    
    '   Insert your code for output 2(Nozzle 2  Cylinder Body)
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
    
        oStPoint.Set 0, -parValveBodyHeight / 2, -cirRadius
        oEnPoint.Set 0, -parValveBodyHeight / 2 - (dNozzle2FacetoBodyFace / 3), -cirRadius
            
    ' Set the output
        m_oGeomHelper.CreateCylinder "Nozzle2CylinderBody", oStPoint, oEnPoint, pipeDiam2
    
    '   Insert your code for output 3(Nozzle 2 Curved Body)
        Dim oRuledSurface As IngrGeom3D.RuledSurface3d
        Dim oTopCircle As IngrGeom3D.Circle3d
        Dim oBottomCircle As IngrGeom3D.Circle3d
        Dim oCenterPoint As New AutoMath.DPosition
        
        ''Top circle
        oCenterPoint.Set 0, -parValveBodyHeight / 2 - (dNozzle2FacetoBodyFace / 3), -cirRadius
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, pipeDiam2 / 2)
        
        ''Bottom circle
        oCenterPoint.Set 0, -parValveBodyHeight / 2 - (2 * dNozzle2FacetoBodyFace / 3), -cirOffesetRadius
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, pipeDiam2 / 2)
        
        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)
    
    ' Set the output
        m_OutputColl.AddOutput "Nozzle2CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing
    
    '   Insert your code for output 4(Nozzle 3  Cylinder Body)
        oStPoint.Set parNozzleCentertoCenter / 2, -parValveBodyHeight / 2, (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        oEnPoint.Set parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle3FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
            
    ' Set the output
        m_oGeomHelper.CreateCylinder "Nozzle3CylinderBody", oStPoint, oEnPoint, pipeDiam3
    
    '   Insert your code for output 5(Nozzle 3 Curved Body)
        ''Top circle
        oCenterPoint.Set parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle3FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, pipeDiam3 / 2)
        
        ''Bottom circle
        oCenterPoint.Set cirOffesetRadius * Cos(PI / 6), -(parValveBodyHeight / 2) - (2 * dNozzle3FacetoBodyFace / 3), _
                cirOffesetRadius * Sin(PI / 6)
                            
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, pipeDiam3 / 2)
        
        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)
    
    ' Set the output
        m_OutputColl.AddOutput "Nozzle3CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing
    
    '   Insert your code for output 6(Nozzle 4  Cylinder Body)
        oStPoint.Set -parNozzleCentertoCenter / 2, -parValveBodyHeight / 2, (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        oEnPoint.Set -parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle4FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
    ' Set the output
        m_oGeomHelper.CreateCylinder "Nozzle4CylinderBody", oStPoint, oEnPoint, pipeDiam4
        
    '   Insert your code for output 7(Nozzle 4  Curved Body)
        ''Top circle
        oCenterPoint.Set -parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle4FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
                        
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, pipeDiam4 / 2)
        
        ''Bottom circle
        oCenterPoint.Set -cirOffesetRadius * Cos(PI / 6), -(parValveBodyHeight / 2) - (2 * dNozzle4FacetoBodyFace / 3), _
            cirOffesetRadius * Sin(PI / 6)
                            
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, pipeDiam4 / 2)
        
        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                            oTopCircle, oBottomCircle, True)
    
    ' Set the output
        m_OutputColl.AddOutput "Nozzle4CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing
        Set oGeomFactory = Nothing
    
    '   Insert your code for output 8(Nozzle 1 with length)
        Set oPlacementPoint = New DPosition
        Set oDirVec = New DVector
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                            flangeThick, flangeDiam, sptOffset1, depth1
        oPlacementPoint.Set 0, parFace1toCenter + sptOffset1 - depth1, 0
        oDirVec.Set 0, 1, 0
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, (parFace1toCenter - parValveBodyHeight / 2))
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    '   Insert your code for output 9(Nozzle 2 with length)
        oPlacementPoint.Set 0, -parFace2toCenter - sptOffset2 + depth2, -cirOffesetRadius
        oDirVec.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, dNozzle2FacetoBodyFace / 3)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    '   Insert your code for output 10(Nozzle 3 with length)
        oPlacementPoint.Set cirOffesetRadius * Cos(PI / 6), -parFace3toCenter - sptOffset3 + depth3, _
                cirOffesetRadius * Sin(PI / 6)
        oDirVec.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, dNozzle3FacetoBodyFace / 3)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
    '   Insert your code for output 11(Nozzle 4 with length)
        oPlacementPoint.Set -cirOffesetRadius * Cos(PI / 6), -parFace4toCenter - sptOffset4 + depth4, _
            cirOffesetRadius * Sin(PI / 6)
        oDirVec.Set 0, -1, 0
        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, _
                                oDirVec, oPlacementPoint, dNozzle4FacetoBodyFace / 3)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    ' Insert your code for output 11(Valve Operator Body)
        Set oDirX = New DVector
        Set oDirY = New DVector
        Set oDirZ = New DVector
    
        Set oPipeComponent = oPartFclt
        On Error GoTo ErrorLabel
    '    Dim oOperatorPart As IJDPart
    '    Dim oOperatorOcc   As IJPartOcc
    
        If Not oPipeComponent Is Nothing Then
          oPipeComponent.GetValveOperatorsForPorts oMultiValCol
        End If
    
        Set OpOrigin = New DPosition
        OpOrigin.Set 0, 0, 0
    
    ''Operator for port 2 (output 12)
          oDirX.Set 0, -1, 0 ''Operator local X-axis matches with Valve negative Y-axis
          oDirY.Set 0, 0, -1 ''Operator local Y-axis matches with Valve negative Z-axis
          oDirZ.Set 1, 0, 0 ''Operator local Z-axis matches with Valve X-axis
    
          Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(2)
                If Not oOperatorPart Is Nothing Then
                    Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator1", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
    
                End If
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
    
    ''Operator for 3rd nozzle (output 13)
        oDirX.Set 0, -1, 0  ''Operator local X-axis matches with Valve negative Y-axis
        oDirY.Set Cos(PI / 6), 0, Sin(PI / 6)
        oDirZ.Set -Sin(PI / 6), 0, Cos(PI / 6)
    
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
                If Not oOperatorPart Is Nothing Then
                    Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
    
                End If
    
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
    
     ''Operator for 4th nozzle
        oDirX.Set 0, -1, 0 ''Operator local X-axis matches with Valve negative Y-axis
        oDirY.Set -Cos(PI / 6), 0, Sin(PI / 6)
        oDirZ.Set -Sin(PI / 6), 0, -Cos(PI / 6)
    
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(4)
                If Not oOperatorPart Is Nothing Then
                    Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator3", _
                                            oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
    
                End If
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
        Set oOperatorOcc = Nothing
        Set oOperatorPart = Nothing
        Set oPipeComponent = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oCenterPoint = Nothing
        Set oPlacementPoint = Nothing
        Set oDirVec = Nothing
        Set oDirX = Nothing
        Set oDirY = Nothing
        Set oDirZ = Nothing
        Set OpOrigin = Nothing
        Set sp3dElem1 = Nothing
     
     ElseIf lPartDataBasis = MULTI_PORT_OPTIONS_3WAY Then '3-way diaphragm diverter valve with inlet and outlet port options
        
        parInletPortGeometry = arrayOfInputs(10)
        parOutletPort1Geometry = arrayOfInputs(11)
        parOutletPort2Geometry = arrayOfInputs(12)
        parOutletPort3Geometry = arrayOfInputs(13)
        parElbowEndFacetoCenter = arrayOfInputs(14)
        parOffsetBetOutlets = arrayOfInputs(15)
        parBodyWidth = arrayOfInputs(16)
        
        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam, sptOffset4, depth4
        
        Dim oVector As AutoMath.DVector
        Dim oTransMat As AutoMath.DT4x4
        
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oVector = New DVector
        Set oTransMat = New DT4x4
        
        'Create the Valve Body
        Dim dLineStrPts(0 To 20) As Double
        Set oLineString = New LineString3d
        
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
        Set objValveBody = PlaceProjection(m_OutputColl, oLineString, oVector, parBodyWidth, True)
    
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing
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
        
        'Create the Outlet Port 3 Geometry
        If parOutletPort3Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace4toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam4, True)
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, pipeDiam4, _
                                parElbowEndFacetoCenter, parFace4toCenter - parBodyWidth / 2, _
                                PI / 2, PI, 0, oTransVec)
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objOutlet = CreatePortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, pipeDiam4, _
                                parOffset, parFace4toCenter - parBodyWidth / 2, 3 * PI / 2, 0, 0, oTransVec)
        End If
        
        'Set the Output
        m_OutputColl.AddOutput "Outlet3", objOutlet
        Set objOutlet = Nothing
        Set oTransVec = Nothing
        
        
        'Remove the references
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oVector = Nothing
        Set oTransMat = Nothing
        Set oGeomFactory = Nothing
        
        'Create the Inlet Nozzle
        Set oPlacementPoint = New DPosition
        Set oDirVec = New DVector
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
        
        'Create the Outlet Nozzle 3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, _
                                        flangeDiam, sptOffset4, depth4
        If parOutletPort3Geometry = STRAIGHT_OUTLET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace4toCenter + sptOffset4 - depth4), 0, _
                                -parOffsetBetOutlets / 2
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oDirVec.Set 0, 0, 1
            oPlacementPoint.Set -parFace4toCenter, 0, _
                    (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter + sptOffset4 - depth4)
                    
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oDirVec.Set -1, 0, 0
            oPlacementPoint.Set -(parFace4toCenter + sptOffset4 - depth4), 0, _
                        -(parOffsetBetOutlets / 2 + parOffset)
        End If
        
        Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        Set oDirVec = Nothing
        Set oPlacementPoint = Nothing
        
    
        'Insert your code for Valve Operator Body
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
    
        'Operator for 3rd nozzle (output 13)
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
        
        'Operator for 4th nozzle (output 14)
        oDirX.Set 0, 1, 0
        oDirY.Set 0, 0, 1
        oDirZ.Set 1, 0, 0
        
        Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(4)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator3", _
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

