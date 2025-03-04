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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Friday, Apr 11 2008
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'  08.Apr.2008     RRK      CR-136576: Enhance in-line silencer symbol to provide more realistic dimensional basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim sptOffset2 As Double
    Dim flangeDiam2 As Double
    Dim depth2 As Double

    Dim iOutput As Double

    Dim parFacetoFace As Double
    Dim parDiameter As Double
    Dim parInsulationThickness As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parDiameter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    ' Insert your code for output 4(Body of VentSilencer)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2

    Dim ObjSilencerBody As IngrGeom3D.Revolution3d
    Dim dCone1BaseDia As Double
    Dim dCone1TopDia As Double
    Dim dCone2BaseDia As Double
    Dim dCone2TopDia As Double
    Dim dCone3BaseDia As Double
    Dim dCone3TopDia As Double
    Dim dCone1Length As Double
    Dim dCone2Length As Double
    Dim dCone3Length As Double
    Dim dCone4Length As Double

    Dim dCylLength As Double
    
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    If lPartDataBasis <= 1 Or lPartDataBasis = 952 Then
        'Cone 1 is the cone at flange of port 1. In case of plain ends it will be cylinder
        'Cone 2 is the cone joining cylinder and cone 1
        'Cone 3 is the cone at flange of port 2. In case of plain ends it will be cylinder
        'Cone 4 is the cone joining cylinder and cone 2

        'Setting the base and top diameters of cones based on whther the end is flanged or not
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            'Assumptions:
            '1)Cone 1 top diameter is 50% of diameter of the silencer
            '2)Cone 1 base diamter is taken as mean of flange diameter of port 1 and cone 1 top diameter
            '3)Cone 3 top diameter is 50% of diameter of the silencer, which is same as cone 1 top diameter
            '4)Cone 1 base diamter is taken as mean of flange diameter of port 2 and cone 2 top diameter
            dCone1TopDia = 0.5 * parDiameter
            dCone1BaseDia = (flangeDiam + dCone1TopDia) / 2
            dCone3TopDia = 0.5 * parDiameter
            dCone3BaseDia = (flangeDiam2 + dCone3TopDia) / 2
        Else
            dCone1TopDia = pipeDiam
            dCone1BaseDia = dCone1TopDia
            dCone3TopDia = pipeDiam2
            dCone3BaseDia = dCone3TopDia
        End If

        'Setting the cylinder length
        'Assumption: Length of the cylinder is taken as 60% length of the silencer excluding flange thickness values of port 1 and port 2
        dCylLength = 0.6 * (parFacetoFace - flangeThick - flangeThick2)

        'Setting the lengths of the cones
        'Assumptions:
        '1)Length of cone 1 is 25% of half of the length of the silencer excluding half the cylinder length and flange thickness of port 1.
        '2)Length of cone 3 is 25% of half of the length of the silencer excluding half the cylinder length and flange thickness of port 2.
        
        dCone1Length = 0.25 * (parFacetoFace / 2 - dCylLength / 2 - flangeThick)
        dCone2Length = parFacetoFace / 2 - dCylLength / 2 - flangeThick - dCone1Length
        dCone3Length = 0.25 * (parFacetoFace / 2 - dCylLength / 2 - flangeThick2)
        dCone4Length = parFacetoFace / 2 - dCylLength / 2 - flangeThick2 - dCone3Length
        
        'Setting the points for the line string
        Dim dPoints() As Double    'representing points in the X-Y plane
        ReDim dPoints(0 To 17)
        'Point 1
        dPoints(0) = -parFacetoFace / 2 + flangeThick
        dPoints(1) = dCone1BaseDia / 2
        dPoints(2) = 0

        'Point 2
        dPoints(3) = dPoints(0) + dCone1Length
        dPoints(4) = dCone1TopDia / 2
        dPoints(5) = 0

        'Point 3
        dPoints(6) = dPoints(3) + dCone2Length
        dPoints(7) = parDiameter / 2
        dPoints(8) = 0

        'Point 4
        dPoints(9) = -dPoints(6)
        dPoints(10) = dPoints(7)
        dPoints(11) = 0

        'Point 5
        dPoints(12) = parFacetoFace / 2 - flangeThick2 - dCone3Length
        dPoints(13) = dCone3TopDia / 2
        dPoints(14) = 0

        'Point 6
        dPoints(15) = dPoints(12) + dCone3Length
        dPoints(16) = dCone3BaseDia / 2
        dPoints(17) = 0

        Dim oLineString As IngrGeom3D.LineString3d
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dPoints)

        Set ObjSilencerBody = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                                       oLineString, 1, 0, 0, 0, 0, 0, 2 * PI, True)
                                                                       
        Set oLineString = Nothing

    ElseIf lPartDataBasis = 953 Then
        
        'Cone 1 is the cone at flange of port 1. In case of plain ends it will be cylinder
        'Cone 2 is the cone at flange of port 2. In case of plain ends it will be cylinder
        'Cone 3 is the cone joining cylinder and cone 2
        
        'Setting the cylinder length
        'Assumption: Length of the cylinder is taken as 60% length of the silencer excluding flange thickness values of port 1 and port 2
        dCylLength = 0.6 * (parFacetoFace - flangeThick - flangeThick2)
        
        'Setting the lengths of the cone
        'Assumptions:
        '1)Length of cone 1 is 25% of half of the length of the silencer excluding half the cylinder length and flange thickness of port 1.
        '2)Length of cone 2 is 25% of half of the length of the silencer excluding half the cylinder length and flange thickness of port 2.
        
        dCone1Length = 0.25 * (parFacetoFace / 2 - dCylLength / 2 - flangeThick)
        dCone2Length = 0.25 * (parFacetoFace / 2 - dCylLength / 2 - flangeThick2)
        dCone3Length = parFacetoFace / 2 - dCylLength / 2 - flangeThick2 - dCone2Length
        
        'Setting the sum of lengths of cylinder 2(of the diffuser) and ellipsoid portion
        Dim dCyl2PlusEllipsoidLength As Double
        dCyl2PlusEllipsoidLength = parFacetoFace / 2 - dCylLength / 2 - flangeThick - dCone1Length
        
        Dim dCyl2Length As Double
    
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            'Assumptions:
            '1)Cone 1 top diameter is 20% of diameter of the silencer
            '2)Cone 1 base diamter is taken as mean of flange diameter of port 1 and cone 1 top diameter
            '3)Cone 2 top diameter is 50% of diameter of the silencer
            '4)Cone 2 base diamter is taken as mean of flange diameter of port 2 and cone 2 top diameter
            dCone1TopDia = 0.2 * parDiameter
            dCone1BaseDia = (flangeDiam + dCone1TopDia) / 2
            dCone2TopDia = 0.5 * parDiameter
            dCone2BaseDia = (flangeDiam2 + dCone2TopDia) / 2
            dCyl2Length = 0.5 * dCyl2PlusEllipsoidLength
        Else
            dCone1TopDia = pipeDiam
            dCone1BaseDia = dCone1TopDia
            dCone2TopDia = pipeDiam2
            dCone2BaseDia = dCone2TopDia
            dCyl2Length = 0.001
        End If
        
        Dim dEllipsoidLength As Double
        dEllipsoidLength = dCyl2PlusEllipsoidLength - dCyl2Length
        
        'Setting the points for the line string
        ReDim dPoints(0 To 20)
        
        'Point 1
        dPoints(0) = -parFacetoFace / 2 + flangeThick
        dPoints(1) = dCone1BaseDia / 2
        dPoints(2) = 0

        'Point 2
        dPoints(3) = dPoints(0) + dCone1Length
        dPoints(4) = dCone1TopDia / 2
        dPoints(5) = 0

        'Point 3
        dPoints(6) = dPoints(3) + dCyl2Length
        dPoints(7) = dCone1TopDia / 2
        dPoints(8) = 0

        'Point 4
        dPoints(9) = dPoints(3) + dCyl2PlusEllipsoidLength
        dPoints(10) = parDiameter / 2
        dPoints(11) = 0

        'Point 5
        dPoints(12) = -dPoints(9)
        dPoints(13) = parDiameter / 2
        dPoints(14) = 0

        'Point 6
        dPoints(15) = parFacetoFace / 2 - flangeThick2 - dCone2Length
        dPoints(16) = dCone2TopDia / 2
        dPoints(17) = 0

        'Point 7
        dPoints(18) = dPoints(15) + dCone2Length
        dPoints(19) = dCone2BaseDia / 2
        dPoints(20) = 0
        
        Dim oStPoint As AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        
        Dim objCollection As Collection
        Set objCollection = New Collection
        
        Dim oLine As IngrGeom3D.Line3d
        
        'Creating line 1
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(0), dPoints(1), dPoints(2), _
                                                         dPoints(3), dPoints(4), dPoints(5))
        objCollection.Add oLine
    
        'Creating line 2
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(3), dPoints(4), dPoints(5), _
                                                         dPoints(6), dPoints(7), dPoints(8))

        objCollection.Add oLine
        
        'Creating ellptical arc
        Dim dMajorRadius As Double
        Dim dMinorRadius As Double

        dMajorRadius = parDiameter / 2 - dCone1TopDia / 2
        dMinorRadius = dEllipsoidLength
        
        Dim oEllipse As Object
        Set oEllipse = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                       Nothing, _
                       dPoints(9), dCone1TopDia / 2, 0, _
                       0, 0, -1, _
                       0, dMajorRadius, 0, _
                       dMinorRadius / dMajorRadius, _
                       -PI / 2, PI / 2)

        objCollection.Add oEllipse
        Set oEllipse = Nothing
        
        'Creating line 3
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(9), dPoints(10), dPoints(11), _
                                                         dPoints(12), dPoints(13), dPoints(14))

        objCollection.Add oLine
        
        'Creating line 4
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(12), dPoints(13), dPoints(14), _
                                                         dPoints(15), dPoints(16), dPoints(17))

        objCollection.Add oLine
        
        'Creating line 5
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(15), dPoints(16), dPoints(17), _
                                                         dPoints(18), dPoints(19), dPoints(20))

        objCollection.Add oLine

        oStPoint.Set dPoints(0), dPoints(1), dPoints(2)

        Dim objBodyOutline As IngrGeom3D.ComplexString3d
        Set objBodyOutline = PlaceTrCString(oStPoint, objCollection)

        Set ObjSilencerBody = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                                       objBodyOutline, 1, 0, 0, 0, 0, 0, 2 * PI, True)
        Set oStPoint = Nothing
        Set objCollection = Nothing
        Set oLine = Nothing
        Set objBodyOutline = Nothing
    End If

    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSilencerBody
    Set ObjSilencerBody = Nothing
    Set oGeomFactory = Nothing

    ' Place Nozzle 1

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    ' Place Nozzle 2
    oPlacePoint.Set parFacetoFace / 2 + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD


End Sub

