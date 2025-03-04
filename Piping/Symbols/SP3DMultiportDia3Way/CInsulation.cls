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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, July 22, 2004
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is 72pdf.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006      KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'  17.Oct.2007      RUK     CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'                           Added code for the new PDB value 461.
'   07.Nov.2007     RUK     TR-128456: Provided a check on end points of non flanged insulation port cylinders in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI       As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

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
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim flangeDiam3      As Double
    Dim flangeDiam4      As Double
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

' Inputs
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
    Dim objValveBody As Object
    
    If lPartDataBasis <= 1 Then 'Default
        
        Dim dInsulationDia1 As Double
        Dim dInsulationDia2 As Double
        Dim dInsulationDia3 As Double
        Dim dInsulationDia4 As Double
        
        Dim dNozzle2FacetoBodyFace  As Double
        Dim dNozzle3FacetoBodyFace  As Double
        Dim dNozzle4FacetoBodyFace  As Double
        
        dNozzle2FacetoBodyFace = parFace2toCenter - parValveBodyHeight / 2
        dNozzle3FacetoBodyFace = parFace3toCenter - parValveBodyHeight / 2
        dNozzle4FacetoBodyFace = parFace4toCenter - parValveBodyHeight / 2
        
    '   Insert your code for output 19(Insulation Valve Body)
        Dim ValveBodyWidth As Double
        
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
    
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam4, sptOffset4, depth4
        
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
        CenterOfTriangleToFace = cirRadius + maxpipeDia / 2 + 0.25 * 0.0254 '0.0254 meter/inch
    
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
         
        Dim transUp         As New AutoMath.DT4x4
    
    ' Use existing ComplexString object, just move it up (on Y axis)
        transUp.[Scale] ((parValveBodyHeight + 2 * parInsulationThickness) / parValveBodyHeight)
        axisVect.Set 0, 1, 0
        Set objValveBody = PlaceProjection(m_OutputColl, oLineString, axisVect, parValveBodyHeight, True)
        objValveBody.Transform transUp  'This will ensure uniform insulation thickness
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing
        Set axisVect = Nothing
        Set transUp = Nothing
        Dim sp3dElem1 As IJDObject
        Set sp3dElem1 = oLineString
        sp3dElem1.Remove
        Set oLineString = Nothing
    
    '   Insert your code for output 2(Nozzle 2  Cylinder Body)
        If pipeDiam2 < flangeDiam2 Then
            dInsulationDia2 = flangeDiam2 + 2 * parInsulationThickness
        Else
            dInsulationDia2 = pipeDiam2 + 2 * parInsulationThickness
        End If
        
        oStPoint.Set 0, -parValveBodyHeight / 2, -cirRadius
        oEnPoint.Set 0, -parValveBodyHeight / 2 - (dNozzle2FacetoBodyFace / 3), -cirRadius
            
    ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle2CylinderBody", oStPoint, oEnPoint, dInsulationDia2
    
    '   Insert your code for output 3(Nozzle 2  Curved Body)
        Dim oRuledSurface As IngrGeom3D.RuledSurface3d
        Dim oTopCircle As IngrGeom3D.Circle3d
        Dim oBottomCircle As IngrGeom3D.Circle3d
        Dim oCenterPoint As New AutoMath.DPosition
        
        ''Top circle
        oCenterPoint.Set 0, -parValveBodyHeight / 2 - (dNozzle2FacetoBodyFace / 3), -cirRadius
    
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia2 / 2)
        
        ''Bottom circle
        oCenterPoint.Set 0, -parValveBodyHeight / 2 - (2 * dNozzle2FacetoBodyFace / 3), -cirOffesetRadius
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia2 / 2)
        
        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)
    
    ' Set the output
        m_OutputColl.AddOutput "InsNozzle2CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing
    
    '   Insert your code for output 4(Nozzle 3  Cylinder Body)
        If pipeDiam3 < flangeDiam3 Then
            dInsulationDia3 = flangeDiam3 + 2 * parInsulationThickness
        Else
            dInsulationDia3 = pipeDiam3 + 2 * parInsulationThickness
        End If
        
        oStPoint.Set parNozzleCentertoCenter / 2, -parValveBodyHeight / 2, (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        oEnPoint.Set parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle3FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
            
    ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle3CylinderBody", oStPoint, oEnPoint, dInsulationDia3
    
    '   Insert your code for output 5(Nozzle 3 Curved Body)
        ''Top circle
        oCenterPoint.Set parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle3FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia3 / 2)
        
        ''Bottom circle
        oCenterPoint.Set cirOffesetRadius * Cos(PI / 6), -(parValveBodyHeight / 2) - (2 * dNozzle3FacetoBodyFace / 3), _
                    cirOffesetRadius * Sin(PI / 6)
                            
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia3 / 2)
        
        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)
    
    ' Set the output
        m_OutputColl.AddOutput "InsNozzle3CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing
        
    '   Insert your code for output 6(Nozzle 4  Cylinder Body)
        If pipeDiam4 < flangeDiam4 Then
            dInsulationDia4 = flangeDiam4 + 2 * parInsulationThickness
        Else
            dInsulationDia4 = pipeDiam4 + 2 * parInsulationThickness
        End If
            
        oStPoint.Set -parNozzleCentertoCenter / 2, -parValveBodyHeight / 2, (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        oEnPoint.Set -parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle4FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
    
    ' Set the output
        m_oGeomHelper.CreateCylinder "InsNozzle4CylinderBody", oStPoint, oEnPoint, dInsulationDia4
        
    '   Insert your code for output 7(Nozzle 4  Curved Body)
        ''Top circle
        oCenterPoint.Set -parNozzleCentertoCenter / 2, _
                        -(parValveBodyHeight / 2) - (dNozzle4FacetoBodyFace / 3), (parNozzleCentertoCenter / 2) * Tan(PI / 6)
        Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia4 / 2)
        
        ''Bottom circle
        oCenterPoint.Set -cirOffesetRadius * Cos(PI / 6), -(parValveBodyHeight / 2) - (2 * dNozzle4FacetoBodyFace / 3), _
                        cirOffesetRadius * Sin(PI / 6)
                            
        Set oBottomCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, oCenterPoint.x, _
                            oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dInsulationDia4 / 2)
        
        Set oRuledSurface = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopCircle, oBottomCircle, True)
    
    ' Set the output
        m_OutputColl.AddOutput "InsNozzle4CurvedBody", oRuledSurface
        Set oRuledSurface = Nothing
        Set oTopCircle = Nothing
        Set oBottomCircle = Nothing
    
    '   Insert your code for output 10(Nozzle 1 Insulation)
        If pipeDiam1 < flangeDiam1 Then
            dInsulationDia1 = flangeDiam1 + 2 * parInsulationThickness
        Else
            dInsulationDia1 = pipeDiam1 + 2 * parInsulationThickness
        End If
                            
        oStPoint.Set 0, parValveBodyHeight / 2, 0
        oEnPoint.Set 0, parValveBodyHeight / 2 + parFace1toCenter, 0
                            
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia1
        
    '   Insert your code for output 11(Nozzle 2 Insulation)
        oStPoint.Set 0, -(parValveBodyHeight / 2) - (2 * dNozzle2FacetoBodyFace / 3), _
                            -cirOffesetRadius
        oEnPoint.Set 0, -parFace2toCenter, -cirOffesetRadius
    
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia2
    
    '   Insert your code for output 12(Nozzle 3 Insulation)
        oStPoint.Set cirOffesetRadius * Cos(PI / 6), -(parValveBodyHeight / 2) - (2 * dNozzle3FacetoBodyFace / 3), _
                        cirOffesetRadius * Sin(PI / 6)
        oEnPoint.Set cirOffesetRadius * Cos(PI / 6), -parFace3toCenter, _
                         cirOffesetRadius * Sin(PI / 6)
                            
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia3
        
    '   Insert your code for output 13(Nozzle 4 Insulation)
        oStPoint.Set -cirOffesetRadius * Cos(PI / 6), -(parValveBodyHeight / 2) - (2 * dNozzle4FacetoBodyFace / 3), _
                    cirOffesetRadius * Sin(PI / 6)
                        
        oEnPoint.Set -cirOffesetRadius * Cos(PI / 6), -parFace4toCenter, _
                        cirOffesetRadius * Sin(PI / 6)
                            
        ' Set the output
        iOutput = iOutput + 1
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDia4
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oGeomFactory = Nothing
        Set oCenterPoint = Nothing
        Set sp3dElem1 = Nothing
      
    ElseIf lPartDataBasis = MULTI_PORT_OPTIONS_3WAY Then
        parInletPortGeometry = arrayOfInputs(10)
        parOutletPort1Geometry = arrayOfInputs(11)
        parOutletPort2Geometry = arrayOfInputs(12)
        parOutletPort3Geometry = arrayOfInputs(13)
        parElbowEndFacetoCenter = arrayOfInputs(14)
        parOffsetBetOutlets = arrayOfInputs(15)
        parBodyWidth = arrayOfInputs(16)
        
        'Retrieve the nozzle parameters
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, sptOffset1, depth1
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset2, depth2
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam3, sptOffset3, depth3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam4, sptOffset4, depth4
        
        'Determine the Insulation Diameter
        Dim dInsDia1 As Double
        Dim dInsDia2 As Double
        Dim dInsDia3 As Double
        Dim dInsDia4 As Double
        
        dInsDia1 = pipeDiam1 + 2 * parInsulationThickness
        dInsDia2 = pipeDiam2 + 2 * parInsulationThickness
        dInsDia3 = pipeDiam3 + 2 * parInsulationThickness
        dInsDia4 = pipeDiam4 + 2 * parInsulationThickness
        
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
        
        oTransMat.LoadIdentity
        oTransMat.[Scale] ((parBodyWidth + 2 * parInsulationThickness) / parBodyWidth)
        
        oVector.Set -1, 0, 0
        Set objValveBody = PlaceProjection(m_OutputColl, oLineString, oVector, parBodyWidth, True)
        objValveBody.Transform oTransMat
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
        Set objValveBody = Nothing
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
            oTransVec.Set 0, 0, parOffsetBetOutlets / 2
            Set objColl = CreateInsPortGeometry(m_OutputColl, parOutletPort3Geometry, oStPoint, dInsDia4, _
                                parElbowEndFacetoCenter, parFace4toCenter - parBodyWidth / 2, _
                                PI / 2, PI, 0, oTransVec)
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
        
        'Create the Outlet nozzle 3
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, _
                                                            flangeDiam, sptOffset4, depth4
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


