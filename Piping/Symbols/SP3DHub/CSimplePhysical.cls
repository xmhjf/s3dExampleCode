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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday 9, Jul 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details were taken from Page no 11 in Techlok.pdf (Techlok catalog).
'    This symbol consists of three physical outputs: two ports and a cone and one insulation output.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  13.MAR.2008     MA   CR-136876 Implemeted Part Data Basis for the values 865, 866, 868, 869 and 870.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI  As Double

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
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double

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
                                                            sptOffset2, depth2
    
    'Center is taken at port 1
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
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    Dim dLineStringPoints(0 To 20) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim objHubBody As Object
    Dim dLineStringPoints1(0 To 23) As Double
    Dim dLineStringPoints2(0 To 29) As Double
    
    Select Case lPartDataBasis
        
        Case Is <= 1

            'Insert your code for output 1(Hub Body)
            Dim dStartDia As Double
            oStPoint.Set flangeThick, 0, 0
            oEnPoint.Set parFacetoface, 0, 0
        
            'Assumption: Start diameter of the cone is pipe diameter plus 60% of the difference between flange diameter and pipe diameter
            dStartDia = pipeDiam + 0.6 * Abs(flangeDiam - pipeDiam)
        
            Set objHubBody = PlaceCone(m_OutputColl, oStPoint, oEnPoint, dStartDia, pipeDiam2)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBody
            Set objHubBody = Nothing
            
        Case 865    'Hub, butt weld
            
            parBodyOutsideDiameter = arrayOfInputs(4)
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints(0) = flangeThick
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = flangeThick
            dLineStringPoints(4) = parBodyOutsideDiameter / 2
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = 0.5 * parFacetoface
            dLineStringPoints(7) = parBodyOutsideDiameter / 2
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = 0.8 * parFacetoface
            dLineStringPoints(10) = pipeDiam2 / 2
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = parFacetoface
            dLineStringPoints(13) = pipeDiam2 / 2
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
            Set objHubBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBody
            Set objHubBody = Nothing
            
        Case 866   'Hub, socket weld
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints(0) = flangeThick
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = flangeThick
            dLineStringPoints(4) = flangeDiam / 2
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = flangeThick + 0.2 * parFacetoface
            dLineStringPoints(7) = 0.4 * flangeDiam
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = flangeThick + 0.4 * parFacetoface
            dLineStringPoints(10) = flangeDiam2 / 2
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = parFacetoface
            dLineStringPoints(13) = flangeDiam2 / 2
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
            Set objHubBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBody
            Set objHubBody = Nothing
            
        Case 868    'Hub, reducing butt weld
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints1(0) = flangeThick
            dLineStringPoints1(1) = 0
            dLineStringPoints1(2) = 0
            
            'Point 2
            dLineStringPoints1(3) = flangeThick
            dLineStringPoints1(4) = flangeDiam / 2
            dLineStringPoints1(5) = 0
            
            'Point 3
            dLineStringPoints1(6) = flangeThick + 0.1 * parFacetoface
            dLineStringPoints1(7) = 0.3 * flangeDiam
            dLineStringPoints1(8) = 0
            
            'Point 4
            dLineStringPoints1(9) = flangeThick + 0.2 * parFacetoface
            dLineStringPoints1(10) = 0.3 * flangeDiam
            dLineStringPoints1(11) = 0
            
            'Point 5
            dLineStringPoints1(12) = flangeThick + 0.5 * parFacetoface
            dLineStringPoints1(13) = pipeDiam2 / 2
            dLineStringPoints1(14) = 0
            
            'Point 6
            dLineStringPoints1(15) = parFacetoface
            dLineStringPoints1(16) = pipeDiam2 / 2
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
            Set objHubBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBody
            Set objHubBody = Nothing
            
        Case 869    'Hub, enlarging butt weld
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints1(0) = flangeThick
            dLineStringPoints1(1) = 0
            dLineStringPoints1(2) = 0
            
            'Point 2
            dLineStringPoints1(3) = flangeThick
            dLineStringPoints1(4) = flangeDiam / 2
            dLineStringPoints1(5) = 0
            
            'Point 3
            dLineStringPoints1(6) = flangeThick + 0.1 * parFacetoface
            dLineStringPoints1(7) = 0.4 * flangeDiam
            dLineStringPoints1(8) = 0
            
            'Point 4
            dLineStringPoints1(9) = flangeThick + 0.3 * parFacetoface
            dLineStringPoints1(10) = 0.4 * flangeDiam
            dLineStringPoints1(11) = 0
            
            'Point 5
            dLineStringPoints1(12) = 0.9 * parFacetoface
            dLineStringPoints1(13) = pipeDiam2 / 2
            dLineStringPoints1(14) = 0
            
            'Point 6
            dLineStringPoints1(15) = parFacetoface
            dLineStringPoints1(16) = pipeDiam2 / 2
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
            Set objHubBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBody
            Set objHubBody = Nothing
            
        Case 870   'Hub, reducing with two hub ends
            
            'Insert your code for output 1(Body of Hub)
            'Point 1
            dLineStringPoints2(0) = flangeThick
            dLineStringPoints2(1) = 0
            dLineStringPoints2(2) = 0
            
            'Point 2
            dLineStringPoints2(3) = flangeThick
            dLineStringPoints2(4) = flangeDiam / 2
            dLineStringPoints2(5) = 0
            
            'Point 3
            dLineStringPoints2(6) = flangeThick + 0.1 * parFacetoface
            dLineStringPoints2(7) = 0.4 * flangeDiam
            dLineStringPoints2(8) = 0
            
            'Point 4
            dLineStringPoints2(9) = flangeThick + 0.3 * parFacetoface
            dLineStringPoints2(10) = 0.4 * flangeDiam
            dLineStringPoints2(11) = 0
            
            'Point 5
            dLineStringPoints2(12) = flangeThick + 0.5 * parFacetoface
            dLineStringPoints2(13) = 0.4 * flangeDiam2
            dLineStringPoints2(14) = 0
            
            'Point 6
            dLineStringPoints2(15) = flangeThick + 0.6 * parFacetoface
            dLineStringPoints2(16) = 0.4 * flangeDiam2
            dLineStringPoints2(17) = 0
            
            'Point 7
            dLineStringPoints2(18) = parFacetoface - flangeThick2
            dLineStringPoints2(19) = flangeDiam2 / 2
            dLineStringPoints2(20) = 0
            
            'Point 8
            dLineStringPoints2(21) = parFacetoface
            dLineStringPoints2(22) = flangeDiam2 / 2
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
            Set objHubBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHubBody
            Set objHubBody = Nothing
            
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.

    End Select
            
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenterPoint = Nothing
    Set oAxisVect = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing
            
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    oPlacePoint.Set parFacetoface + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

