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
''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'       Default:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is 72pdf. PDS symbol MC=VS3WD.

'       PDB: 462
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 2 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'       PDB: 880
'       This is a 4-way diverver valve symbol. This is prepared based on Gemu's catalog.
'       Source: Gemu Multiport Valves M600-5-4B, Aseptic valve manifold machined from a single block.
'       M600 valve manifold designs, developed and produced according to customer requirements/specifications
'       The symbol has multiple operators. Each operator for each output port
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  17.Oct.2007      RUK     CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'                           Added code for the new PDB value 461.
'   07.Nov.2007     RUK     TR-128456: Provided a check on end points of non flanged insulation port cylinders in which case small cylinders of negligible thickness will be created
'   09-June-2008    MP      CR-141585  Multiport valve symbols need to be enhanced to address Gemu valve requirements. (Implemented part data basis: 459)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "Insulation:"    'Used for error messages
Private PI As Double
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

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim flangeDiam As Double
    Dim cptOffset As Double
    Dim depth As Double
    Dim pipeDiam1 As Double
    Dim pipeDiam2 As Double
    Dim pipeDiam3 As Double
    Dim pipeDiam4 As Double
    Dim pipeDiam5 As Double
    Dim flangeDiam1 As Double
    Dim flangeDiam2 As Double
    Dim flangeDiam3 As Double
    Dim flangeDiam4 As Double
    Dim flangeDiam5 As Double
    Dim cptOffset1 As Double
    Dim depth1 As Double
    Dim cptOffset2 As Double
    Dim depth2 As Double
    Dim cptOffset3 As Double
    Dim depth3 As Double
    Dim cptOffset4 As Double
    Dim depth4 As Double
    Dim cptOffset5 As Double
    Dim depth5 As Double
    Dim iOutput As Double

    Dim parValveBodyWidth As Double
    Dim parValveBodyDepth As Double
    Dim parValveBodyHeight As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parFace5toCenter As Double
    Dim parOffset As Double
    Dim parInletPortGeometry As Double
    Dim parOutletPort1Geometry As Double
    Dim parOutletPort2Geometry As Double
    Dim parOutletPort3Geometry As Double
    Dim parOutletPort4Geometry As Double
    Dim parElbowEndFacetoCenter As Double
    Dim parOffsetBetOutlets As Double
    Dim parBodyWidth As Double
    Dim parInsulationThickness As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveBodyWidth = arrayOfInputs(2)
    parValveBodyDepth = arrayOfInputs(3)
    parValveBodyHeight = arrayOfInputs(4)
    parFace1toCenter = arrayOfInputs(5)
    parFace2toCenter = arrayOfInputs(6)
    parFace3toCenter = arrayOfInputs(7)
    parFace4toCenter = arrayOfInputs(8)
    parFace5toCenter = arrayOfInputs(9)
    'parOffset = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)

    iOutput = 0

    m_oGeomHelper.OutputCollection = m_OutputColl

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim objNozzle As Object
    Dim ObjValveBody As Object

    If lPartDataBasis <= 1 Then    'Default
        Dim dInsulationDia1 As Double
        Dim dInsulationDia2 As Double
        Dim dInsulationDia3 As Double
        Dim dInsulationDia4 As Double
        Dim dInsulationDia5 As Double

        ' Inputs
        parOffset = arrayOfInputs(10)

        '   Insert your code for output 19(Insulation Valve Body)
        Dim oStartPoint As New AutoMath.DPosition
        Dim oEndPoint As New AutoMath.DPosition

        oStartPoint.Set -parValveBodyWidth / 2 - parInsulationThickness, _
                        -parValveBodyHeight / 2 - parInsulationThickness, _
                        parValveBodyDepth / 2 + parInsulationThickness
        oEndPoint.Set parValveBodyWidth / 2 + parInsulationThickness, _
                      parValveBodyHeight / 2 + parInsulationThickness, _
                      -parValveBodyDepth / 2 - parInsulationThickness

        Set ObjValveBody = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing

        '   Insert your code for output 2(Nozzle 2  Cylinder Body)

        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptOffset, depth
        If pipeDiam < flangeDiam Then
            dInsulationDia2 = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia2 = pipeDiam + 2 * parInsulationThickness
        End If

        oStartPoint.Set 0, -parValveBodyHeight / 2, (parValveBodyDepth - pipeDiam) / 2
        oEndPoint.Set 0, -(parValveBodyHeight / 2) - (parFace2toCenter / 3), (parValveBodyDepth - pipeDiam) / 2

        ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle2CylinderBody", oStartPoint, oEndPoint, dInsulationDia2

        '   Insert your code for output 3(Nozzle 2  Curved Body)
        Set oGeomFactory = New GeometryFactory
        Dim oRuledSurface As IngrGeom3D.RuledSurface3d
        Dim oTopCircle As IngrGeom3D.Circle3d
        Dim oBottomCircle As IngrGeom3D.Circle3d
        Dim oCenterPoint As New AutoMath.DPosition

        ''Top circle
        oCenterPoint.Set 0, -(parValveBodyHeight / 2) - (parFace2toCenter / 3), _
                         (parValveBodyDepth - pipeDiam) / 2
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                           oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia2 / 2)

        ''Bottom circle
        oCenterPoint.Set 0, -(parValveBodyHeight / 2) - (2 * parFace2toCenter / 3), parOffset
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                              oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia2 / 2)

        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)

        ' Set the output
        m_OutputColl.AddOutput "InsNozzle2CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing

        '   Insert your code for output 4(Nozzle 3  Cylinder Body)

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptOffset, depth

        If pipeDiam < flangeDiam Then
            dInsulationDia3 = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia3 = pipeDiam + 2 * parInsulationThickness
        End If

        oStartPoint.Set (parValveBodyWidth - pipeDiam) / 2, _
                        -parValveBodyHeight / 2, 0
        oEndPoint.Set (parValveBodyWidth - pipeDiam) / 2, _
                      -(parValveBodyHeight / 2) - (parFace3toCenter / 3), 0

        ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle3CylinderBody", oStartPoint, oEndPoint, dInsulationDia3

        '   Insert your code for output 5(Nozzle 3  Curved Body)

        ''Top circle
        oCenterPoint.Set (parValveBodyWidth - pipeDiam) / 2, _
                         -(parValveBodyHeight / 2) - (parFace3toCenter / 3), 0
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                           oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia3 / 2)

        ''Bottom circle
        oCenterPoint.Set parOffset, -(parValveBodyHeight / 2) - (2 * parFace3toCenter / 3), 0

        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                              oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia3 / 3)

        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)

        ' Set the output
        m_OutputColl.AddOutput "InsNozzle3CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing

        '   Insert your code for output 6(Nozzle 4  Cylinder Body)

        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptOffset, depth

        If pipeDiam < flangeDiam Then
            dInsulationDia4 = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia4 = pipeDiam + 2 * parInsulationThickness
        End If

        oStartPoint.Set 0, -parValveBodyHeight / 2, -(parValveBodyDepth - pipeDiam) / 2
        oEndPoint.Set 0, -(parValveBodyHeight / 2) - (parFace4toCenter / 3), -(parValveBodyDepth - pipeDiam) / 2

        ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle4CylinderBody", oStartPoint, oEndPoint, dInsulationDia4

        '   Insert your code for output 7(Nozzle 4  Curved Body)
        ''Top circle
        oCenterPoint.Set 0, -(parValveBodyHeight / 2) - (parFace4toCenter / 3), _
                         -(parValveBodyDepth - pipeDiam) / 2
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                           oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia4 / 2)

        ''Bottom circle
        oCenterPoint.Set 0, -(parValveBodyHeight / 2) - (2 * parFace4toCenter / 3), -parOffset

        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                              oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia4 / 2)

        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)

        ' Set the output
        m_OutputColl.AddOutput "InsNozzle4CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing

        '   Insert your code for output 8(Nozzle 5  Cylinder Body)

        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptOffset, depth

        If pipeDiam < flangeDiam Then
            dInsulationDia5 = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia5 = pipeDiam + 2 * parInsulationThickness
        End If

        oStartPoint.Set -(parValveBodyWidth - pipeDiam) / 2, _
                        -parValveBodyHeight / 2, 0
        oEndPoint.Set -(parValveBodyWidth - pipeDiam) / 2, _
                      -(parValveBodyHeight / 2) - (parFace5toCenter / 3), 0

        ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle5CylinderBody", oStartPoint, oEndPoint, dInsulationDia5

        '  Insert your code for output 9(Nozzle 5  Curved Body)

        ''Top circle
        oCenterPoint.Set -(parValveBodyWidth - pipeDiam) / 2, _
                         -(parValveBodyHeight / 2) - (parFace5toCenter / 3), 0
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                           oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia5 / 2)

        ''Bottom circle
        oCenterPoint.Set -parOffset, -(parValveBodyHeight / 2) - (2 * parFace5toCenter / 3), 0

        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                                                                              oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia5 / 2)

        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)

        ' Set the output
        m_OutputColl.AddOutput "InsNozzle5CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing

        '   Insert your code for output 10(Nozzle 1 Insulation)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptOffset, depth

        If pipeDiam < flangeDiam Then
            dInsulationDia1 = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDia1 = pipeDiam + 2 * parInsulationThickness
        End If

        oStartPoint.Set 0, parValveBodyHeight / 2, 0
        oEndPoint.Set 0, parValveBodyHeight / 2 + parFace1toCenter, 0

        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia1

        '   Insert your code for output 11(Nozzle 2 Insulation)
        oStartPoint.Set 0, -(parValveBodyHeight / 2) - (2 * parFace2toCenter / 3), parOffset
        oEndPoint.Set 0, -parFace2toCenter - (parValveBodyHeight / 2), parOffset

        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia2

        '   Insert your code for output 12(Nozzle 3 Insulation)
        oStartPoint.Set parOffset, -(parValveBodyHeight / 2) - (2 * parFace3toCenter / 3), 0
        oEndPoint.Set parOffset, -parFace3toCenter - (parValveBodyHeight / 2), 0

        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia3

        '   Insert your code for output 13(Nozzle 4 Insulation)
        oStartPoint.Set 0, -(parValveBodyHeight / 2) - (2 * parFace4toCenter / 3), -parOffset
        oEndPoint.Set 0, -parFace4toCenter - (parValveBodyHeight / 2), -parOffset

        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia4

        '   Insert your code for output 14(Nozzle 5 Insulation)
        oStartPoint.Set -parOffset, -(parValveBodyHeight / 2) - (2 * parFace5toCenter / 3), 0
        oEndPoint.Set -parOffset, -parFace5toCenter - (parValveBodyHeight / 2), 0

        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia5

    ElseIf lPartDataBasis = MULTI_PORT_OPTIONS_4WAY Then

        parOffset = arrayOfInputs(10)
        parInletPortGeometry = arrayOfInputs(12)
        parOutletPort1Geometry = arrayOfInputs(13)
        parOutletPort2Geometry = arrayOfInputs(14)
        parOutletPort3Geometry = arrayOfInputs(15)
        parOutletPort4Geometry = arrayOfInputs(16)
        parElbowEndFacetoCenter = arrayOfInputs(17)
        parOffsetBetOutlets = arrayOfInputs(18)
        parBodyWidth = arrayOfInputs(19)

        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, cptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, cptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, cptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam4, cptOffset4, depth4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick, flangeDiam, cptOffset5, depth5

        'Determine the Insulation Diameter
        Dim dInsDia1 As Double
        Dim dInsDia2 As Double
        Dim dInsDia3 As Double
        Dim dInsDia4 As Double
        Dim dInsDia5 As Double

        dInsDia1 = pipeDiam1 + 2 * parInsulationThickness
        dInsDia2 = pipeDiam2 + 2 * parInsulationThickness
        dInsDia3 = pipeDiam3 + 2 * parInsulationThickness
        dInsDia4 = pipeDiam4 + 2 * parInsulationThickness
        dInsDia5 = pipeDiam5 + 2 * parInsulationThickness

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
        Dim oLineString As IngrGeom3D.LineString3d
        Dim dLineStrPts(0 To 26) As Double
        Set oLineString = New LineString3d

        Dim dSide As Double

        If CmpDblGreaterthan(parBodyWidth, parOffsetBetOutlets _
                                         + ReturnMax4(pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5)) Then
            dSide = parBodyWidth + 0.0254
        Else
            dSide = parOffsetBetOutlets _
                  + ReturnMax4(pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5) + 0.0254
        End If

        dLineStrPts(0) = parBodyWidth / 2
        dLineStrPts(1) = parBodyWidth / 2
        dLineStrPts(2) = dSide / 2

        dLineStrPts(3) = dLineStrPts(0)
        dLineStrPts(4) = -dLineStrPts(1)
        dLineStrPts(5) = dLineStrPts(2)

        dLineStrPts(6) = dLineStrPts(0)
        dLineStrPts(7) = -dLineStrPts(2)
        dLineStrPts(8) = dLineStrPts(1)

        dLineStrPts(9) = dLineStrPts(0)
        dLineStrPts(10) = -dLineStrPts(2)
        dLineStrPts(11) = -dLineStrPts(1)

        dLineStrPts(12) = dLineStrPts(0)
        dLineStrPts(13) = -dLineStrPts(1)
        dLineStrPts(14) = -dLineStrPts(2)

        dLineStrPts(15) = dLineStrPts(0)
        dLineStrPts(16) = dLineStrPts(1)
        dLineStrPts(17) = -dLineStrPts(2)

        dLineStrPts(18) = dLineStrPts(0)
        dLineStrPts(19) = dLineStrPts(2)
        dLineStrPts(20) = -dLineStrPts(1)

        dLineStrPts(21) = dLineStrPts(0)
        dLineStrPts(22) = dLineStrPts(2)
        dLineStrPts(23) = dLineStrPts(1)

        dLineStrPts(24) = dLineStrPts(0)
        dLineStrPts(25) = dLineStrPts(1)
        dLineStrPts(26) = dLineStrPts(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dLineStrPts)

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
            m_OutputColl.AddOutput "InsInletPort_", objInlet
            Set objInlet = Nothing
        ElseIf parInletPortGeometry = INLET_WITH_90DEG_ELBOW Then
            Set objColl = New Collection
            oStPoint.Set parBodyWidth / 2, 0, 0
            Set objColl = CreateInsPortGeometry(m_OutputColl, parInletPortGeometry, oStPoint, dInsDia1, _
                                                parElbowEndFacetoCenter, parFace1toCenter - parBodyWidth / 2)
            'Set the Output
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
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort2_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        End If

        'Create the Outlet Port 3 Geometry
        If parOutletPort3Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -parFace4toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia4, True)
            'Set the Output
            m_OutputColl.AddOutput "InsOutletPort3_", objOutlet
            Set objOutlet = Nothing
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            Set objColl = New Collection
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, dInsDia4, _
                                                parElbowEndFacetoCenter, parFace4toCenter - parBodyWidth / 2, _
                                              3 * PI / 2, PI, 0, oTransVec)
            'Set the Output
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort3_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            Set objColl = New Collection
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, -parOffsetBetOutlets / 2
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort3Geometry, _
                                                oStPoint, dInsDia4, parOffset, _
                                                parFace4toCenter - parBodyWidth / 2, 3 * PI / 2, 0, 0, oTransVec)
            'Set the Output
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort3_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        End If

        'Create the Outlet Port 4 Geometry
        If parOutletPort4Geometry = STRAIGHT_OUTLET Then
            oStPoint.Set -parBodyWidth / 2, 0, parOffsetBetOutlets / 2
            oEnPoint.Set -parFace5toCenter, oStPoint.y, oStPoint.z
            Set objOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia5, True)
            'Set the Output
            m_OutputColl.AddOutput "InsOutletPort4_", objOutlet
            Set objOutlet = Nothing
        ElseIf parOutletPort4Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, parOffsetBetOutlets / 2
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort4Geometry, oStPoint, dInsDia5, _
                                                parElbowEndFacetoCenter, parFace5toCenter - parBodyWidth / 2, _
                                                PI / 2, PI, 0, oTransVec)
            'Set the Output
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort4_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        ElseIf parOutletPort4Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parBodyWidth / 2, 0, 0
            oTransVec.Set 0, 0, parOffsetBetOutlets / 2
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort4Geometry, oStPoint, dInsDia5, _
                                                parOffset, parFace5toCenter - parBodyWidth / 2, PI / 2, 0, 0, oTransVec)
            'Set the Output
            For iCount = 1 To objColl.Count
                m_OutputColl.AddOutput "InsOutletPort4_", objColl.Item(iCount)
            Next iCount
            For iCount = 1 To objColl.Count
                objColl.Remove 1
            Next iCount
            Set objColl = Nothing
        End If
        Set oTransVec = Nothing

        'Create the Inlet Nozzle
        Dim dInsThick As Double
        'Determine the Insulation Diameter1 for the Inlet Nozzle
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, _
                           cptOffset1, depth1
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
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, cptOffset2, depth2
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
            oStPoint.Set -parFace2toCenter, _
                         (parOffsetBetOutlets / 2 + parOffset), 0
            oEnPoint.Set -(parFace2toCenter - flangeThick - dInsThick), _
                         (parOffsetBetOutlets / 2 + parOffset), 0
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia2, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing


        'Create the Outlet nozzle 2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, cptOffset3, depth3
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

        'Create the Outlet nozzle 3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, _
                           flangeDiam, cptOffset4, depth4
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
            oStPoint.Set -parFace4toCenter, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -(parFace4toCenter - flangeThick - dInsThick), 0, _
                         -parOffsetBetOutlets / 2
        ElseIf parOutletPort3Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace4toCenter, 0, (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter)
            oEnPoint.Set -parFace4toCenter, 0, (parOffsetBetOutlets / 2 _
                                              + parElbowEndFacetoCenter - flangeThick - dInsThick)
        ElseIf parOutletPort3Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace4toCenter, 0, -(parOffsetBetOutlets / 2 + parOffset)
            oEnPoint.Set -(parFace4toCenter - flangeThick - dInsThick), 0, _
                         -(parOffsetBetOutlets / 2 + parOffset)
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia4, True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        'Create the Outlet nozzle 4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick, _
                           flangeDiam, cptOffset5, depth5
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
            oStPoint.Set -parFace5toCenter, 0, -parOffsetBetOutlets / 2
            oEnPoint.Set -(parFace5toCenter - flangeThick - dInsThick), 0, _
                         -parOffsetBetOutlets / 2
        ElseIf parOutletPort4Geometry = OUTLET_WITH_90DEG_ELBOW Then
            oStPoint.Set -parFace5toCenter, 0, (parOffsetBetOutlets / 2 + parElbowEndFacetoCenter)
            oEnPoint.Set -parFace5toCenter, 0, (parOffsetBetOutlets / 2 _
                                              + parElbowEndFacetoCenter - flangeThick - dInsThick)
        ElseIf parOutletPort4Geometry = OUTLET_WITH_OFFSET Then
            oStPoint.Set -parFace5toCenter, 0, -(parOffsetBetOutlets / 2 + parOffset)
            oEnPoint.Set -(parFace5toCenter - flangeThick - dInsThick), 0, _
                         -(parOffsetBetOutlets / 2 + parOffset)
        End If
        Set objNozzle = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia5, True)
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
    ElseIf lPartDataBasis = MULTI_PORT_OPTIONS_4WAY_GEMU Then

        iOutput = 0

        'Origin is taken at the geometry center of the valve.

        m_oGeomHelper.OutputCollection = m_OutputColl

        '   Insert your code for output 12(Insulated Valve Body)

        Dim oTopLineString As IngrGeom3D.LineString3d
        Dim oBottomLineString As IngrGeom3D.LineString3d
        Dim dLinePoints(0 To 14) As Double
        Dim dStepHeight As Double
        Dim oDirVec As New AutoMath.DVector

        Set oGeomFactory = New GeometryFactory
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                           flangeThick, flangeDiam1, cptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                           flangeThick, flangeDiam2, cptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                           flangeThick, flangeDiam3, cptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, _
                           flangeThick, flangeDiam4, cptOffset4, depth4
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, _
                           flangeThick, flangeDiam5, cptOffset5, depth5

        ''Step height is assumed to be 1.5 times the biggest Pipe diameter of Port2,Port3,Port4,Port5
        dStepHeight = 1.5 * Max5(pipeDiam1, pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5)

        Dim dValveBodyTopWidth As Double
        ''Width of the Body Top is assumed to be 1.5 times the Pipe diameter of Nozzle 1.
        ''Top surface is assume to be a Squre.
        dValveBodyTopWidth = 1.5 * pipeDiam1
        ''Valve Body top surface Line points.
        dLinePoints(0) = parInsulationThickness + dValveBodyTopWidth / 2
        dLinePoints(1) = parInsulationThickness + parValveBodyHeight / 2
        dLinePoints(2) = parInsulationThickness + dValveBodyTopWidth / 2

        dLinePoints(3) = dLinePoints(0)
        dLinePoints(4) = dLinePoints(1)
        dLinePoints(5) = -dLinePoints(2)

        dLinePoints(6) = -dLinePoints(3)
        dLinePoints(7) = dLinePoints(4)
        dLinePoints(8) = dLinePoints(5)

        dLinePoints(9) = dLinePoints(6)
        dLinePoints(10) = dLinePoints(7)
        dLinePoints(11) = dLinePoints(2)

        dLinePoints(12) = parInsulationThickness + dValveBodyTopWidth / 2
        dLinePoints(13) = parInsulationThickness + parValveBodyHeight / 2
        dLinePoints(14) = parInsulationThickness + dValveBodyTopWidth / 2

        Set oTopLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePoints)

        ''Valve Body Bottom surface Line points.
        dLinePoints(0) = parInsulationThickness + parValveBodyWidth / 2
        dLinePoints(1) = dStepHeight - parValveBodyHeight / 2
        dLinePoints(2) = parInsulationThickness + parValveBodyDepth / 2

        dLinePoints(3) = dLinePoints(0)
        dLinePoints(4) = dLinePoints(1)
        dLinePoints(5) = -dLinePoints(2)

        dLinePoints(6) = -dLinePoints(3)
        dLinePoints(7) = dLinePoints(4)
        dLinePoints(8) = dLinePoints(5)

        dLinePoints(9) = dLinePoints(6)
        dLinePoints(10) = dLinePoints(7)
        dLinePoints(11) = dLinePoints(2)
        dLinePoints(12) = parInsulationThickness + parValveBodyWidth / 2
        dLinePoints(13) = dStepHeight - parValveBodyHeight / 2
        dLinePoints(14) = parInsulationThickness + parValveBodyDepth / 2

        Set oBottomLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePoints)

        Set ObjValveBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopLineString, _
                                                                       oBottomLineString, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing

        '   Insert your code for output 13(Stepped part of valve body Insualtion)
        Dim ObjValveBodyStep As Object

        oStartPoint.Set dLinePoints(0), dLinePoints(1), dLinePoints(2)
        oEndPoint.Set dLinePoints(6), -parInsulationThickness - parValveBodyHeight / 2, dLinePoints(8)

        Set ObjValveBodyStep = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyStep
        Set ObjValveBodyStep = Nothing

        '   Insert your code for output 14(Nozzle 1 insulation)
        Dim dInsulationDia As Double

        If pipeDiam1 < flangeDiam1 Then
            dInsulationDia = flangeDiam1 + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam1 + 2 * parInsulationThickness
        End If
        oStartPoint.Set 0, parFace1toCenter - parValveBodyHeight / 2, 0
        oEndPoint.Set 0, parFace1toCenter, 0
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

        '   Insert your code for output 15(Nozzle 2 Insualation)

        If pipeDiam2 < flangeDiam2 Then
            dInsulationDia = flangeDiam2 + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
        End If

        oStartPoint.Set parFace2toCenter - parValveBodyWidth / 2, -parValveBodyHeight / 2 + dStepHeight / 2, 0
        oEndPoint.Set parFace2toCenter, -parValveBodyHeight / 2 + dStepHeight / 2, 0
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

        '   Insert your code for output 16(Nozzle 3 Insualation)

        If pipeDiam3 < flangeDiam3 Then
            dInsulationDia = flangeDiam3 + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam3 + 2 * parInsulationThickness
        End If
        oStartPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, -parFace3toCenter + parValveBodyDepth / 2
        oEndPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, -parFace3toCenter

        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

        '   Insert your code for output 17(Nozzle 4 Insualation)

        If pipeDiam4 < flangeDiam4 Then
            dInsulationDia = flangeDiam4 + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam4 + 2 * parInsulationThickness
        End If

        oStartPoint.Set -parFace4toCenter + parValveBodyWidth / 2, -parValveBodyHeight / 2 + dStepHeight / 2, 0
        oEndPoint.Set -parFace4toCenter, -parValveBodyHeight / 2 + dStepHeight / 2, 0
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia


        '   Insert your code for output 18(Nozzle 5 Insualation)
        If pipeDiam5 < flangeDiam5 Then
            dInsulationDia = flangeDiam5 + 2 * parInsulationThickness
        Else
            dInsulationDia = pipeDiam5 + 2 * parInsulationThickness
        End If

        oStartPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, parFace5toCenter - parValveBodyDepth / 2
        oEndPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, parFace5toCenter
        ' Set the output
        m_oGeomHelper.CreateCylinder "ABCD", oStartPoint, oEndPoint, dInsulationDia

        Set oStartPoint = Nothing
        Set oEndPoint = Nothing

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
