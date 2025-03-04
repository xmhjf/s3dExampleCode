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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Friday 9, Jul 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details were taken from Page no 11 in Techlok.pdf (Techlok catalog).
'    This symbol consists of three physical outputs: two ports and a cone and one insulation output.

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  13.MAR.2008     MA   CR-136876 Implemeted Part Data Basis for the values 865, 866, 868, 869 and 870.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeDiam2      As Double
    Dim flangeThick2     As Double

    Dim iOutput     As Double

    Dim parFacetoface As Double
    Dim parInsulationThickness As Double
    Dim parBodyOutsideDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoface = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parBodyOutsideDiameter = arrayOfInputs(4)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                            sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                            sptOffset, depth
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    ' This symbol is based on the following three Part data basis values that govern its geometry,
    '   Hub, specified by face to face (Default case)
    '   Hub, butt weld - 865
    '   Hub, socket weld - 866
    '   Hub, reducing butt weld - 868
    '   Hub, enlarging butt weld - 869
    '   Hub, reducing with two hub ends - 870
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim dLineStringPoints(0 To 20) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim objHubBodyIns As Object
    Dim objInsCyl As Object
    Dim dLineStringPoints1(0 To 23) As Double
    Dim dLineStringPoints2(0 To 29) As Double
    
    Select Case lPartDataBasis
        
        Case Is <= 1
    
            'Insert your code for output 1(Hub insulation)
            Dim dStartDia As Double
            Dim dEndDia As Double
            
            oStPoint.Set flangeThick, 0, 0
            oEnPoint.Set parFacetoface, 0, 0
        
            'Assumption: Start diameter of the cone is pipe diameter plus 60% of the difference between flange diameter and pipe diameter
            dStartDia = pipeDiam + 0.6 * Abs(flangeDiam - pipeDiam) + 2 * parInsulationThickness
            
            If CmpDblGreaterthan(pipeDiam2, flangeDiam2) Then
                dEndDia = pipeDiam2 + 2 * parInsulationThickness
            Else
                dEndDia = flangeDiam2 + 2 * parInsulationThickness
            End If
            
            Set objHubBodyIns = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                                                    dStartDia, dEndDia)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBodyIns
            Set objHubBodyIns = Nothing
            
        Case 865
            
            parBodyOutsideDiameter = arrayOfInputs(4)
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints(0) = flangeThick
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = flangeThick
            dLineStringPoints(4) = parBodyOutsideDiameter / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = 0.5 * parFacetoface
            dLineStringPoints(7) = parBodyOutsideDiameter / 2 + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = 0.8 * parFacetoface
            dLineStringPoints(10) = pipeDiam2 / 2 + parInsulationThickness
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = parFacetoface
            dLineStringPoints(13) = pipeDiam2 / 2 + parInsulationThickness
            dLineStringPoints(14) = 0
            
            'Point 6
            dLineStringPoints(15) = parFacetoface
            dLineStringPoints(16) = 0
            dLineStringPoints(17) = 0
            
            'Point 7
            dLineStringPoints(18) = flangeThick
            dLineStringPoints(19) = 0
            dLineStringPoints(20) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, -0.00001, 0
            Set objHubBodyIns = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBodyIns
            Set objHubBodyIns = Nothing
            
        Case 866   'Hub, socket weld
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints(0) = flangeThick
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = flangeThick
            dLineStringPoints(4) = flangeDiam / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = flangeThick + 0.2 * parFacetoface
            dLineStringPoints(7) = 0.4 * flangeDiam + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = flangeThick + 0.4 * parFacetoface
            dLineStringPoints(10) = flangeDiam2 / 2 + parInsulationThickness
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = parFacetoface
            dLineStringPoints(13) = flangeDiam2 / 2 + parInsulationThickness
            dLineStringPoints(14) = 0
            
            'Point 6
            dLineStringPoints(15) = parFacetoface
            dLineStringPoints(16) = 0
            dLineStringPoints(17) = 0
            
            'Point 7
            dLineStringPoints(18) = flangeThick
            dLineStringPoints(19) = 0
            dLineStringPoints(20) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, -0.00001, 0
            Set objHubBodyIns = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBodyIns
            Set objHubBodyIns = Nothing
            
        Case 868    'Hub, reducing butt weld
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints1(0) = flangeThick
            dLineStringPoints1(1) = 0
            dLineStringPoints1(2) = 0
            
            'Point 2
            dLineStringPoints1(3) = flangeThick
            dLineStringPoints1(4) = flangeDiam / 2 + parInsulationThickness
            dLineStringPoints1(5) = 0
            
            'Point 3
            dLineStringPoints1(6) = flangeThick + 0.1 * parFacetoface
            dLineStringPoints1(7) = 0.3 * flangeDiam + parInsulationThickness
            dLineStringPoints1(8) = 0
            
            'Point 4
            dLineStringPoints1(9) = flangeThick + 0.2 * parFacetoface
            dLineStringPoints1(10) = 0.3 * flangeDiam + parInsulationThickness
            dLineStringPoints1(11) = 0
            
            'Point 5
            dLineStringPoints1(12) = flangeThick + 0.5 * parFacetoface
            dLineStringPoints1(13) = pipeDiam2 / 2 + parInsulationThickness
            dLineStringPoints1(14) = 0
            
            'Point 6
            dLineStringPoints1(15) = parFacetoface
            dLineStringPoints1(16) = pipeDiam2 / 2 + parInsulationThickness
            dLineStringPoints1(17) = 0
            
            'Point 7
            dLineStringPoints1(18) = parFacetoface
            dLineStringPoints1(19) = 0
            dLineStringPoints1(20) = 0
            
            'Point 8
            dLineStringPoints1(21) = flangeThick
            dLineStringPoints1(22) = 0
            dLineStringPoints1(23) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStringPoints1)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, -0.00001, 0
            Set objHubBodyIns = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBodyIns
            Set objHubBodyIns = Nothing
            
        Case 869    'Hub, enlarging butt weld
        
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints1(0) = flangeThick
            dLineStringPoints1(1) = 0
            dLineStringPoints1(2) = 0
            
            'Point 2
            dLineStringPoints1(3) = flangeThick
            dLineStringPoints1(4) = flangeDiam / 2 + parInsulationThickness
            dLineStringPoints1(5) = 0
            
            'Point 3
            dLineStringPoints1(6) = flangeThick + 0.1 * parFacetoface
            dLineStringPoints1(7) = 0.4 * flangeDiam + parInsulationThickness
            dLineStringPoints1(8) = 0
            
            'Point 4
            dLineStringPoints1(9) = flangeThick + 0.3 * parFacetoface
            dLineStringPoints1(10) = 0.4 * flangeDiam + parInsulationThickness
            dLineStringPoints1(11) = 0
            
            'Point 5
            dLineStringPoints1(12) = 0.9 * parFacetoface
            dLineStringPoints1(13) = pipeDiam2 / 2 + parInsulationThickness
            dLineStringPoints1(14) = 0
            
            'Point 6
            dLineStringPoints1(15) = parFacetoface
            dLineStringPoints1(16) = pipeDiam2 / 2 + parInsulationThickness
            dLineStringPoints1(17) = 0
            
            'Point 7
            dLineStringPoints1(18) = parFacetoface
            dLineStringPoints1(19) = 0
            dLineStringPoints1(20) = 0
            
            'Point 8
            dLineStringPoints1(21) = flangeThick
            dLineStringPoints1(22) = 0
            dLineStringPoints1(23) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 8, dLineStringPoints1)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, -0.00001, 0
            Set objHubBodyIns = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBodyIns
            Set objHubBodyIns = Nothing
        
        Case 870   'Hub, reducing with two hub ends
        
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints2(0) = flangeThick
            dLineStringPoints2(1) = 0
            dLineStringPoints2(2) = 0
            
            'Point 2
            dLineStringPoints2(3) = flangeThick
            dLineStringPoints2(4) = flangeDiam / 2 + parInsulationThickness
            dLineStringPoints2(5) = 0
            
            'Point 3
            dLineStringPoints2(6) = flangeThick + 0.1 * parFacetoface
            dLineStringPoints2(7) = 0.4 * flangeDiam + parInsulationThickness
            dLineStringPoints2(8) = 0
            
            'Point 4
            dLineStringPoints2(9) = flangeThick + 0.3 * parFacetoface
            dLineStringPoints2(10) = 0.4 * flangeDiam + parInsulationThickness
            dLineStringPoints2(11) = 0
            
            'Point 5
            dLineStringPoints2(12) = flangeThick + 0.5 * parFacetoface
            dLineStringPoints2(13) = 0.4 * flangeDiam2 + parInsulationThickness
            dLineStringPoints2(14) = 0
            
            'Point 6
            dLineStringPoints2(15) = flangeThick + 0.6 * parFacetoface
            dLineStringPoints2(16) = 0.4 * flangeDiam2 + parInsulationThickness
            dLineStringPoints2(17) = 0
            
            'Point 7
            dLineStringPoints2(18) = parFacetoface - flangeThick2
            dLineStringPoints2(19) = flangeDiam2 / 2 + parInsulationThickness
            dLineStringPoints2(20) = 0
            
            'Point 8
            dLineStringPoints2(21) = parFacetoface
            dLineStringPoints2(22) = flangeDiam2 / 2 + parInsulationThickness
            dLineStringPoints2(23) = 0
            
            'Point 9
            dLineStringPoints2(24) = parFacetoface
            dLineStringPoints2(25) = 0
            dLineStringPoints2(26) = 0
            
            'Point 10
            dLineStringPoints2(27) = flangeThick
            dLineStringPoints2(28) = 0
            dLineStringPoints2(29) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 10, dLineStringPoints2)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, -0.00001, 0
            Set objHubBodyIns = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBodyIns
            Set objHubBodyIns = Nothing
        
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.

    End Select
    
    If (lPartDataBasis = 865) Or (lPartDataBasis = 866) Or (lPartDataBasis = 868) Or _
                        (lPartDataBasis = 869) Or (lPartDataBasis = 870) Then
    
        'Place Insulation for Port 1
        Dim dFlangeBodyThickness As Double
        If CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
            If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
            oStPoint.Set 0, 0, 0
            dFlangeBodyThickness = flangeThick + parInsulationThickness
            If CmpDblGreaterthan(dFlangeBodyThickness, parFacetoface) Then dFlangeBodyThickness = parFacetoface
            oEnPoint.Set dFlangeBodyThickness, 0, 0
            
            Dim objInsulatedPort1 As Object
            Set objInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            flangeDiam + 2 * parInsulationThickness, True)
    
            'Set the output
            m_OutputColl.AddOutput "InsPort1", objInsulatedPort1
            Set objInsulatedPort1 = Nothing
        End If
    
        'Place Insulation for Port 2
        If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = NEGLIGIBLE_THICKNESS
            oStPoint.Set parFacetoface, 0, 0
            dFlangeBodyThickness = parFacetoface - flangeThick2 - parInsulationThickness
            If CmpDblLessThan(dFlangeBodyThickness, 0) Then dFlangeBodyThickness = 0
            oEnPoint.Set dFlangeBodyThickness, 0, 0
            
            Dim objInsulatedPort2 As Object
            Set objInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    flangeDiam2 + 2 * parInsulationThickness, True)
    
            'Set the output
            m_OutputColl.AddOutput "InsPort2", objInsulatedPort2
            Set objInsulatedPort2 = Nothing
        End If
    
    End If
 
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenterPoint = Nothing
    Set oAxisVect = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

