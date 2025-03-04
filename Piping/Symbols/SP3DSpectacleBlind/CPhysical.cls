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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Sunday, May 11 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Source for PDB 1046 - ASME 16.48 Spectacle blind
'   Source for PDB 1047 - www.resistoflex.com/pdf/specblind.pdf
'   Source for PDB 1048 - Marine Company Standard for JIS 5K
'   Source for PDB 1049 - http://www.swecofab.com/spectacle/600_female_oct_ring.htm
'   Source for PDB 1050 - www.micromold.com
'   Change History:
'   dd.mmm.yyyy          who                 change description
'   -----------         -----             ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006         KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007         KKC                 TR-126158  Corrected the Blind Diameter to be equal to the Flange Outer Diameter
'   12.Aug.2008         KKC                 CR-146914  Spectacle blind symbol to be specified by figure 8 inter-axis distance
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInlineSpect As Object
    Dim ObjInterSpectCyl As Object
    Dim ObjDisc As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parFig8InterAxis As Double
    Dim parSpectaclePosition As Double
    Dim parDiscDiameter As Double
    Dim parWebWidth As Double
    Dim parWebThickness As Double
    Dim parRadius As Double
    Dim parCentertoCenter As Double
    Dim parWidth As Double
    Dim parBoltHoleDiameter As Double
    Dim parWebSpacing As Double
    Dim parCenterOffset As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnpoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnpoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    If lPartDataBasis <= 1 Then
        parFacetoFace = arrayOfInputs(2)
        parFig8InterAxis = arrayOfInputs(3)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        If UBound(arrayOfInputs) >= 6 Then
            parDiscDiameter = arrayOfInputs(6)
        End If
    ElseIf lPartDataBasis = 1046 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parWebWidth = arrayOfInputs(8)
        parWebThickness = arrayOfInputs(9)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1047 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parRadius = arrayOfInputs(10)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1048 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parRadius = arrayOfInputs(10)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1049 Then
        parFacetoFace = arrayOfInputs(2)
        parCentertoCenter = arrayOfInputs(7)
        parWebWidth = arrayOfInputs(8)
        parWebThickness = arrayOfInputs(9)
        parWebSpacing = arrayOfInputs(11)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        
    ElseIf lPartDataBasis = 1050 Then
        parFacetoFace = arrayOfInputs(2)
        parCenterOffset = arrayOfInputs(12)
        parWidth = arrayOfInputs(13)
        parBoltHoleDiameter = arrayOfInputs(14)
        parSpectaclePosition = arrayOfInputs(4)
        parInsulationThickness = arrayOfInputs(5)
        parCentertoCenter = 2 * parCenterOffset
        parRadius = parWidth / 2
    End If
    iOutput = 0
    If lPartDataBasis <= 1 Then
    
     ' Insert your code for output 1(In line Spectacle)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        If CmpDblEqual(parDiscDiameter, 0) And Not CmpDblEqual(flangeDiam, 0) Then
            parDiscDiameter = flangeDiam
        Else
            parDiscDiameter = pipeDiam * 1.1
        End If
        Dim Port1Position   As New AutoMath.DPosition
        Port1Position.Set -parFacetoFace / 2, 0, 0
        Dim Port2Position   As New AutoMath.DPosition
        Port2Position.Set parFacetoFace / 2, 0, 0
    
        Set ObjInlineSpect = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, parDiscDiameter, False)
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlineSpect
        Set ObjInlineSpect = Nothing
        Set Port1Position = Nothing
        Set Port2Position = Nothing
        
     ' Insert your code for output 2(Cylinder bet Discs)
        Dim parInterDiscsDiameter As Double
        parInterDiscsDiameter = parFacetoFace / 2
        Dim CylStPosition   As New AutoMath.DPosition
        CylStPosition.Set 0, parFig8InterAxis * 0.3, 0
        Dim CylEndPosition   As New AutoMath.DPosition
        CylEndPosition.Set 0, parFig8InterAxis * 0.7, 0
        Set ObjInterSpectCyl = PlaceCylinder(m_OutputColl, CylStPosition, CylEndPosition, parInterDiscsDiameter, False)
    
    ' Set the output
        m_OutputColl.AddOutput "OBjWeb_", ObjInterSpectCyl
        Set ObjInterSpectCyl = Nothing
        Set CylStPosition = Nothing
        Set CylEndPosition = Nothing
        
     ' Insert your code for output 3(Disc)
         RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        If parSpectaclePosition = 1 Then 'when parSpectaclePosition=1 open
            Dim BlankStPosition   As New AutoMath.DPosition
            BlankStPosition.Set -parFacetoFace / 2, parFig8InterAxis, 0
            Dim BlankEndPosition   As New AutoMath.DPosition
            BlankEndPosition.Set parFacetoFace / 2, parFig8InterAxis, 0
        
            Set ObjDisc = PlaceCylinder(m_OutputColl, BlankStPosition, BlankEndPosition, parDiscDiameter, True)
            Set BlankStPosition = Nothing
            Set BlankEndPosition = Nothing
    
        Else ' when parSpectaclePosition=2 closed
            Dim axis        As New AutoMath.DVector
            axis.Set 1, 0, 0
            Dim OpenDiscAxisCentPosition   As New AutoMath.DPosition
            OpenDiscAxisCentPosition.Set 0, parFig8InterAxis, 0
            
            Dim parPipetoDisc As Double
            parPipetoDisc = (parDiscDiameter - pipeDiam) / 2
            Dim parHubThick As Double
            parHubThick = parPipetoDisc * 0.2
            
            Dim lines           As Collection
            Dim oLine           As IngrGeom3D.Line3d
            Dim oGeomFactory    As IngrGeom3D.GeometryFactory
            Dim objCStr         As IngrGeom3D.ComplexString3d
        
            Set lines = New Collection
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, -parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, -parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0, parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0)
            lines.Add oLine
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.5, parFig8InterAxis + parDiscDiameter / 2, 0, parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.5, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2 + parHubThick, 0, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0, -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0)
            lines.Add oLine
            
            Dim stPoint   As New AutoMath.DPosition
            stPoint.Set -parFacetoFace * 0.6, parFig8InterAxis + pipeDiam / 2, 0
            Set objCStr = PlaceTrCString(stPoint, lines)
            Set oLine = Nothing
            Dim iCount As Integer
            For iCount = 1 To lines.Count
                lines.Remove 1
            Next iCount
            Set lines = Nothing
            'Use our complex string for revolution
            Set ObjDisc = PlaceRevolution(m_OutputColl, objCStr, axis, OpenDiscAxisCentPosition, PI * 2, False)
            Set objCStr = Nothing
            Set axis = Nothing
            Set OpenDiscAxisCentPosition = Nothing
            Set oGeomFactory = Nothing
            Set stPoint = Nothing
        End If
        
    ' Set the output
        m_OutputColl.AddOutput "OBjWeb_", ObjDisc
        Set ObjDisc = Nothing
        
    ElseIf lPartDataBasis = 1046 Or lPartDataBasis = 1049 Then
        
        ' Insert your code for output 1(In line Spectacle)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        Dim oPort1Position   As New AutoMath.DPosition
        
        oPort1Position.Set -parFacetoFace / 2, 0, 0
        Dim oPort2Position   As New AutoMath.DPosition
        oPort2Position.Set parFacetoFace / 2, 0, 0
        Set ObjInlineSpect = PlaceCylinder(m_OutputColl, oPort1Position, oPort2Position, flangeDiam, False)
        
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlineSpect
        Set ObjInlineSpect = Nothing
        Set oPort1Position = Nothing
        Set oPort2Position = Nothing
        
    ' Insert your code for output 2(web between Discs)
        If lPartDataBasis = 1046 Then
            oStPoint.Set -parWebThickness / 2, 0.4 * flangeDiam, parWebWidth / 2
            oEnpoint.Set parWebThickness / 2, (parCentertoCenter - 0.6 * pipeDiam), -parWebWidth / 2
            Dim ObjWeb As Object
            Set ObjWeb = PlaceBox(m_OutputColl, oStPoint, oEnpoint)
        
        ' Set the output
            m_OutputColl.AddOutput "OBjWeb_", ObjWeb
            Set ObjWeb = Nothing
            
        ElseIf lPartDataBasis = 1049 Then
            oStPoint.Set -parWebThickness / 2, 0.4 * flangeDiam, parWebSpacing / 2 + parWebThickness
            oEnpoint.Set parWebThickness / 2, (parCentertoCenter - 0.6 * pipeDiam), parWebSpacing / 2
            
            Set ObjWeb = PlaceBox(m_OutputColl, oStPoint, oEnpoint)
        ' Set the output
            m_OutputColl.AddOutput "OBjWeb_", ObjWeb
            Set ObjWeb = Nothing
        ' Insert code the Web 2
            oStPoint.Set -parWebThickness / 2, 0.4 * flangeDiam, -parWebSpacing / 2 - parWebThickness
            oEnpoint.Set parWebThickness / 2, (parCentertoCenter - 0.6 * pipeDiam), -parWebSpacing / 2
            Dim ObjWeb1 As Object
            Set ObjWeb1 = PlaceBox(m_OutputColl, oStPoint, oEnpoint)
        ' Set the output
            m_OutputColl.AddOutput "OBjWeb_", ObjWeb1
            Set ObjWeb = Nothing
        
        End If
        
        ' Insert your code for output 3(Disc)
         RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        If parSpectaclePosition = 1 Then 'when parSpectaclePosition=1 open
            Dim oBlankStPosition   As New AutoMath.DPosition
            oBlankStPosition.Set -parFacetoFace / 2, parCentertoCenter, 0
            Dim oBlankEndPosition   As New AutoMath.DPosition
            oBlankEndPosition.Set parFacetoFace / 2, parCentertoCenter, 0
        
            Set ObjDisc = PlaceCylinder(m_OutputColl, oBlankStPosition, oBlankEndPosition, flangeDiam, True)
            Set BlankStPosition = Nothing
            Set BlankEndPosition = Nothing
    
        Else ' when parSpectaclePosition=2 closed
            Dim oAxis  As AutoMath.DVector
            Set oAxis = New AutoMath.DVector
            oAxis.Set 1, 0, 0
            Dim oOpenDiscAxisCentPosition  As AutoMath.DPosition
            Set oOpenDiscAxisCentPosition = New AutoMath.DPosition
            oOpenDiscAxisCentPosition.Set 0, parCentertoCenter, 0
            Dim dHubThk As Double
            dHubThk = (flangeDiam - pipeDiam) / 2
             
            Set lines = New Collection
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0)
            lines.Add oLine
            
            oStPoint.Set -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0
            Set objCStr = PlaceTrCString(oStPoint, lines)
            Set oLine = Nothing
            For iCount = 1 To lines.Count
                lines.Remove 1
            Next iCount
            Set lines = Nothing
            'Use our complex string for revolution
            Set ObjDisc = PlaceRevolution(m_OutputColl, objCStr, oAxis, oOpenDiscAxisCentPosition, PI * 2, False)
            Set objCStr = Nothing
            Set axis = Nothing
            Set OpenDiscAxisCentPosition = Nothing
            Set oGeomFactory = Nothing
            Set oStPoint = Nothing
        End If
        
    ' Set the output
        m_OutputColl.AddOutput "OBjWeb_", ObjDisc
        Set ObjDisc = Nothing
        
    ElseIf lPartDataBasis = 1047 Or lPartDataBasis = 1050 Then
        
        ' Insert your code for output 1(In line Spectacle)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        oPort1Position.Set -parFacetoFace / 2, 0, 0
        oPort2Position.Set parFacetoFace / 2, 0, 0
        Set ObjInlineSpect = PlaceCylinder(m_OutputColl, oPort1Position, oPort2Position, flangeDiam, False)
        
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlineSpect
        Set ObjInlineSpect = Nothing
        Set oPort1Position = Nothing
        Set oPort2Position = Nothing
        
    ' Insert your code for the outer geometry(Output 2)
        Dim OMdPoint As AutoMath.DPosition
        Set OMdPoint = New AutoMath.DPosition
        
        oStPoint.Set -0.45 * parFacetoFace, parCentertoCenter - parRadius * Cos(PI / 4), parRadius * Sin(PI / 4)
        OMdPoint.Set -0.45 * parFacetoFace, parCentertoCenter + parRadius, 0
        oEnpoint.Set -0.45 * parFacetoFace, parCentertoCenter - parRadius * Cos(PI / 4), -parRadius * Sin(PI / 4)
        Dim ObjArc1 As IngrGeom3D.Arc3d
        Set ObjArc1 = New IngrGeom3D.Arc3d
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set ObjArc1 = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, OMdPoint.x, _
                                                    OMdPoint.y, OMdPoint.z, oEnpoint.x, oEnpoint.y, oEnpoint.z)
        Dim oCollection As Collection
        Set oCollection = New Collection
        oCollection.Add ObjArc1
        Set ObjArc1 = Nothing
        Dim ObjEllipticalArc As IngrGeom3D.EllipticalArc3d
        Dim dMMRatio As Double
        dMMRatio = 0.2
        Dim OCenterPoint As AutoMath.DPosition
        Set OCenterPoint = New AutoMath.DPosition
        OCenterPoint.Set -0.45 * parFacetoFace, parCentertoCenter / 2, -parRadius * Sin(PI / 4)
        
        Set ObjEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, OCenterPoint.x, OCenterPoint.y, OCenterPoint.z, _
                                                                     1, 0, 0, 0, (parCentertoCenter - 2 * parRadius * Cos(PI / 4)) / 2, 0, dMMRatio, 0, PI)
                                                            
       
        oCollection.Add ObjEllipticalArc
        Dim ObjArc2 As IngrGeom3D.Arc3d
        oStPoint.Set -0.45 * parFacetoFace, parRadius * Cos(PI / 4), -parRadius * Sin(PI / 4)
        OMdPoint.Set -0.45 * parFacetoFace, -parRadius, 0
        oEnpoint.Set -0.45 * parFacetoFace, parRadius * Cos(PI / 4), parRadius * Sin(PI / 4)
        
        Set ObjArc2 = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, OMdPoint.x, _
                                                    OMdPoint.y, OMdPoint.z, oEnpoint.x, oEnpoint.y, oEnpoint.z)

        OCenterPoint.Set -0.45 * parFacetoFace, parCentertoCenter / 2, parRadius * Sin(PI / 4)
        oCollection.Add ObjArc2
        Dim ObjEllipticalArc1 As IngrGeom3D.EllipticalArc3d

        Set ObjEllipticalArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, OCenterPoint.x, OCenterPoint.y, OCenterPoint.z, _
                                                                     1, 0, 0, 0, -(parCentertoCenter / 2 - (parRadius * Cos(PI / 4))), 0, dMMRatio, 0, PI)
        oCollection.Add ObjEllipticalArc1
        Dim ObjCmplxStr As IngrGeom3D.ComplexString3d
        oStPoint.Set -0.45 * parFacetoFace, parCentertoCenter - parRadius * Cos(PI / 4), parRadius * Sin(PI / 4)
        Set ObjCmplxStr = PlaceTrCString(oStPoint, oCollection)
        Dim oTopSurface As IngrGeom3D.Plane3d
        Set oTopSurface = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, -0.45 * parFacetoFace, 0, 0, 1, 0, 0)
        oTopSurface.AddBoundary ObjCmplxStr
        If lPartDataBasis = 1050 Then
                Dim ObjHoleCur1 As IngrGeom3D.Circle3d
                Set ObjHoleCur1 = New IngrGeom3D.Circle3d
                Set ObjHoleCur1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -0.45 * parFacetoFace, parCentertoCenter / 2, 0, 1, 0, 0, parBoltHoleDiameter / 2)
                Dim ObjCmplxStr1 As IngrGeom3D.ComplexString3d
                Set ObjCmplxStr1 = New IngrGeom3D.ComplexString3d
                ObjCmplxStr1.AddCurve ObjHoleCur1, True
                oTopSurface.AddHole ObjCmplxStr1
        End If
        ' Insert your code for output 3(Disc)
         
        If parSpectaclePosition = 1 Then 'when parSpectaclePosition=1 open
            Dim oSurfset  As IngrGeom3D.IJElements
            Set oSurfset = oGeomFactory.GeometryServices.CreateByProjection(m_OutputColl.ResourceManager, oTopSurface, 1, 0, 0, 0.9 * parFacetoFace, 1)
            Dim oProjection As Object
            For Each oProjection In oSurfset
                m_OutputColl.AddOutput "Spectacle_", oProjection
            Next oProjection
            Dim ObjFlange As Object
            oStPoint.Set -parFacetoFace / 2, parCentertoCenter, 0
            oEnpoint.Set parFacetoFace / 2, parCentertoCenter, 0
            If CmpDblGreaterthan(flangeDiam, (2 * parRadius)) Then flangeDiam = 2 * parRadius
            Set ObjFlange = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, flangeDiam, True)
            m_OutputColl.AddOutput "SpectaclePro_", ObjFlange
            Set ObjFlange = Nothing
           
        ElseIf parSpectaclePosition = 2 Then ' when parSpectaclePosition=2 closed
            Set oAxis = New AutoMath.DVector
            oAxis.Set 1, 0, 0
            Set oOpenDiscAxisCentPosition = New AutoMath.DPosition
            oOpenDiscAxisCentPosition.Set 0, parCentertoCenter, 0
            dHubThk = (flangeDiam - pipeDiam) / 2
             
            Set lines = New Collection
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0)
            lines.Add oLine
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2 + dHubThk, 0, -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0)
            lines.Add oLine
            
            oStPoint.Set -parFacetoFace / 2, parCentertoCenter - flangeDiam / 2, 0
            Set objCStr = PlaceTrCString(oStPoint, lines)
            Set oLine = Nothing
            For iCount = 1 To lines.Count
                lines.Remove 1
            Next iCount
            Set lines = Nothing
            'Use our complex string for revolution
            Set ObjDisc = PlaceRevolution(m_OutputColl, objCStr, oAxis, oOpenDiscAxisCentPosition, PI * 2, False)
            m_OutputColl.AddOutput "ObjDisc_", ObjDisc
            Set ObjDisc = Nothing
            
            Dim ObjHoleCur As IngrGeom3D.Circle3d
            Set ObjHoleCur = New IngrGeom3D.Circle3d
            Set ObjHoleCur = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -0.45 * parFacetoFace, parCentertoCenter, 0, 1, 0, 0, pipeDiam / 2)
            Set ObjCmplxStr = New IngrGeom3D.ComplexString3d
            ObjCmplxStr.AddCurve ObjHoleCur, True
            oTopSurface.AddHole ObjCmplxStr
            Dim oSurfset1  As IngrGeom3D.IJElements
            Set oSurfset1 = oGeomFactory.GeometryServices.CreateByProjection(m_OutputColl.ResourceManager, oTopSurface, 1, 0, 0, 0.9 * parFacetoFace, 1)
            Dim oProjection1 As Object
            For Each oProjection1 In oSurfset1
                m_OutputColl.AddOutput "Spectacle_", oProjection1
            Next oProjection1
        End If
    ElseIf lPartDataBasis = 1048 Then
    
        ' Insert your code for output 1(In line Spectacle)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        oPort1Position.Set -parFacetoFace / 2, 0, 0
        oPort2Position.Set parFacetoFace / 2, 0, 0
        Set ObjInlineSpect = PlaceCylinder(m_OutputColl, oPort1Position, oPort2Position, flangeDiam, False)
        
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInlineSpect
        Set ObjInlineSpect = Nothing
        Set oPort1Position = Nothing
        Set oPort2Position = Nothing
        
    ' Insert your code for the outer geometry(Output 2)
        Set OMdPoint = New AutoMath.DPosition
        
        oStPoint.Set -0.45 * parFacetoFace, parCentertoCenter, -parRadius
        OMdPoint.Set -0.45 * parFacetoFace, parCentertoCenter + parRadius, 0
        oEnpoint.Set -0.45 * parFacetoFace, parCentertoCenter, parRadius
        Set ObjArc1 = New IngrGeom3D.Arc3d
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        Set ObjArc1 = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, OMdPoint.x, _
                                                    OMdPoint.y, OMdPoint.z, oEnpoint.x, oEnpoint.y, oEnpoint.z)
        Set oCollection = New Collection
        oCollection.Add ObjArc1
        Set ObjArc1 = Nothing
        oStPoint.Set -0.45 * parFacetoFace, parCentertoCenter, parRadius
        oEnpoint.Set -0.45 * parFacetoFace, 0, parRadius
        
        Dim Oline1 As IngrGeom3D.Line3d
        Set Oline1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oEnpoint.x, _
                                                                oEnpoint.y, oEnpoint.z)
                                                            
        oCollection.Add Oline1
        oStPoint.Set -0.45 * parFacetoFace, 0, parRadius
        OMdPoint.Set -0.45 * parFacetoFace, -parRadius, 0
        oEnpoint.Set -0.45 * parFacetoFace, 0, -parRadius
        
        Set ObjArc2 = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, OMdPoint.x, _
                                                    OMdPoint.y, OMdPoint.z, oEnpoint.x, oEnpoint.y, oEnpoint.z)
        oCollection.Add ObjArc2
    
        oStPoint.Set -0.45 * parFacetoFace, 0, -parRadius
        oEnpoint.Set -0.45 * parFacetoFace, parCentertoCenter, -parRadius
        
        Dim Oline2 As IngrGeom3D.Line3d
        Set Oline2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oEnpoint.x, _
                                                                oEnpoint.y, oEnpoint.z)
                                                            
        oCollection.Add Oline2
        oStPoint.Set -0.45 * parFacetoFace, parCentertoCenter, -parRadius
        Set ObjCmplxStr = New IngrGeom3D.ComplexString3d
        Set ObjCmplxStr = PlaceTrCString(oStPoint, oCollection)
        
        If parSpectaclePosition = 1 Then 'when parSpectaclePosition=1 open
            Set oAxis = New AutoMath.DVector
            oAxis.Set 1, 0, 0
            Dim ObjCurProjection As Object
            Set ObjCurProjection = PlaceProjection(m_OutputColl, ObjCmplxStr, oAxis, 0.9 * parFacetoFace, True)
            
            'Set the Output
            m_OutputColl.AddOutput "SpectaclePro_", ObjCurProjection
            Set oCollection = Nothing
            Set ObjCmplxStr = Nothing
           
        ElseIf parSpectaclePosition = 2 Then ' when parSpectaclePosition=2 closed
            Set oTopSurface = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, -0.45 * parFacetoFace, 0, 0, 1, 0, 0)
            oTopSurface.AddBoundary ObjCmplxStr
            Set ObjHoleCur = New IngrGeom3D.Circle3d
            Set ObjHoleCur = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -0.45 * parFacetoFace, parCentertoCenter, 0, 1, 0, 0, pipeDiam / 2)
            Set ObjCmplxStr = New IngrGeom3D.ComplexString3d
            ObjCmplxStr.AddCurve ObjHoleCur, True
            oTopSurface.AddHole ObjCmplxStr
            Set oSurfset = oGeomFactory.GeometryServices.CreateByProjection(m_OutputColl.ResourceManager, oTopSurface, 1, 0, 0, 0.9 * parFacetoFace, 1)
            For Each oProjection In oSurfset
                m_OutputColl.AddOutput "Spectacle_", oProjection
            Next oProjection
        End If
        Set ObjCurProjection = Nothing
     End If
' Place Nozzle 1
       
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)
    
    oPlacePoint.Set -faceToFace / 2 - cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set faceToFace / 2 + cptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
