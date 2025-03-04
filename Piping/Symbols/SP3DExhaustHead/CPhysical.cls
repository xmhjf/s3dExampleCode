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
'   Author:          Babu Govindarajan
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the catalog available at URL http://www.nciweb.net/exhaust.htm
'   This symbol is based on the following four Part data basis values that govern its geometry,
'   1.Exhaust Head,Type 1-900:www.nciweb.net/Anderson%20PDF/Exhaust_Heads.pdf
'   2.Exhaust Head,Type 2-901:www.nciweb.net/Anderson%20PDF/Exhaust_Heads.pdf
'   3.Exhaust Head,Type 3-902:www.nciweb.net/Anderson%20PDF/Exhaust_Heads_EH.pdf
'   4.Exhaust Head,Type 4-903:www.pennseparator.com/ehsd2.htm
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   30.Jan.2006     Sundar(svsmylav)        RI-28367: Deleted reference to 'SN=S7AZ' of PDS in the
'                                           revision history and added hyper link to Anderson's site.
'                                           Also, modified symbol input parameters to match that of the
'                                           manufacturer's skecth.
'  24.Apr.2008      VRK                     CR-136579: Enhance exhaust head symbol
'                                            to provide more realistic dimensional basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double

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
    Dim pipeDiam3 As Double
    Dim flangeThick3 As Double
    Dim sptOffset3 As Double
    Dim flangeDiam3 As Double
    Dim depth3 As Double
    Dim iOutput As Integer

    Dim ObjExhaustCone As Object
    Dim parExhaustDiameter As Double
    Dim parExhaustHeight As Double
    Dim parPortVOffset As Double
    Dim parPortHOffset As Double
    Dim parDiameter As Double
    Dim parHeight As Double
    Dim parDrainPortXOffset As Double
    Dim parDrainPortYOffset As Double
    Dim parInletPortXOffset As Double
    Dim parInsulationThickness As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '    parExhaustDiameter = arrayOfInputs(2) 'A
    '    parExhaustHeight = arrayOfInputs(3)   'B
    '    parPortVOffset = arrayOfInputs(4)     'C
    '    parPortHOffset = arrayOfInputs(5)     'E
    '    parInsulationThickness = arrayOfInputs(6)
    '    parDiameter = arrayOfInputs(7)
    '    parHeight = arrayOfInputs(8)
    '    parDrainPortXOffset =arrayOfInputs(9)
    '    parDrainPortYOffset =arrayOfInputs(10)
    '    parInletPortXOffset=arrayOfInputs(11)

    iOutput = 0

    ' Assumption: Origin of the symbol is taken to be at center of port 1


    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim oStpoint As AutoMath.DPosition
    Dim oEnpoint As AutoMath.DPosition
    Set oStpoint = New AutoMath.DPosition
    Set oEnpoint = New AutoMath.DPosition
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory
    Dim oCenter As New AutoMath.DPosition
    Dim oArc As IngrGeom3D.EllipticalArc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim objCurve As Object
    Dim objCollection As Collection
    Dim dMMRatio As Double
    Dim dHorizontalOffset As Double
    Dim objCylinder As Object
    Dim oAxisVec As AutoMath.DVector
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle


    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                       flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                       flangeDiam2, sptOffset2, depth2

    If lPartDataBasis = 900 Or lPartDataBasis = 901 Or lPartDataBasis = 902 _
                        Or lPartDataBasis = 903 Then
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, _
                           flangeDiam3, sptOffset3, depth3
    End If

    'This symbol is based on the following four Part data basis values that govern its geometry,
    'Exhaust Head,Type 1-900
    'Exhaust Head,Type 2-901
    'Exhaust Head,Type 3-902
    'Exhaust Head,Type 4-903

    Select Case lPartDataBasis
    'Earlier, the default case exhaust head is implemented to have two ports in v2007.
    'We kept it as is for backward compatibility. Other part data basis values are
    'implemented to have three ports.
    Case Is <= 1
    
        ' Insert your code for output 1(Exhaust Cone)
        parExhaustDiameter = arrayOfInputs(2) 'A
        parExhaustHeight = arrayOfInputs(3)   'B
        parPortVOffset = arrayOfInputs(4)     'C
        parPortHOffset = arrayOfInputs(5)     'E

        Dim baseCenter As AutoMath.DPosition
        Dim topCenter As AutoMath.DPosition
        Set baseCenter = New AutoMath.DPosition
        Set topCenter = New AutoMath.DPosition
        baseCenter.Set 0, flangeThick, 0
        topCenter.Set 0, parExhaustHeight, 0
        Set ObjExhaustCone = PlaceCone(m_OutputColl, baseCenter, topCenter, pipeDiam / 2, parExhaustDiameter / 2, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExhaustCone
        Set ObjExhaustCone = Nothing
        Set baseCenter = Nothing
        Set topCenter = Nothing

        ' Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set 0, -sptOffset + depth, 0
        oDir.Set 0, -1, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        ' Place Nozzle 2
        oPlacePoint.Set parPortHOffset, parPortVOffset - sptOffset + depth, 0
        oDir.Set 0, -1, 0
        Dim dNozzleLength As Double
        dNozzleLength = parExhaustHeight - parPortVOffset    'B -C
        If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 900
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        dHorizontalOffset = 0.1 * parHeight

        Set objCollection = New Collection
        oStpoint.Set flangeThick, 0, 0
        oEnpoint.Set flangeThick, pipeDiam / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Elliptical Arc 1
        oCenter.Set flangeThick, pipeDiam / 2 + (parDiameter - pipeDiam) / 4, 0
        dMMRatio = ((parDiameter - pipeDiam) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, 1, _
                                                                                       1.6 * dHorizontalOffset, 0, 0, dMMRatio, -PI / 2, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Elliptical Arc 2
        oCenter.Set flangeThick + 3.2 * dHorizontalOffset, parDiameter / 2 - ((parDiameter - pipeDiam) / 4), oCenter.z
        dMMRatio = ((parDiameter - pipeDiam) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, -1, _
                                                                                       -1.6 * dHorizontalOffset, 0, 0, dMMRatio, 0, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Line 1
        oStpoint.Set flangeThick + 3.2 * dHorizontalOffset, parDiameter / 2, oCenter.z
        oEnpoint.Set 0.9 * parHeight, parDiameter / 2, oCenter.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Create small circular Arc
        'Arc 1
        Dim oCurve As IngrGeom3D.Arc3d
        oStpoint.Set 0.9 * parHeight, parDiameter / 2, oCenter.z
        oEnpoint.Set parHeight, (parDiameter / 2) - dHorizontalOffset, oStpoint.z
        oCenter.Set oStpoint.x, oEnpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'St Line
        oStpoint.Set parHeight, (parDiameter / 2) - dHorizontalOffset, oCenter.z
        oEnpoint.Set parHeight, 0, oCenter.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        oStpoint.Set flangeThick, 0, oStpoint.z
        Set oAxisVec = New DVector
        oAxisVec.Set -1, 0, 0
        Set oComplexStr = PlaceTrCString(oStpoint, objCollection)
        oStpoint.Set flangeThick, 0, oStpoint.z
        Set objCurve = PlaceRevolution(m_OutputColl, oComplexStr, oAxisVec, oStpoint, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurve
        Set objCurve = Nothing
        Set oAxisVec = Nothing
        Set oComplexStr = Nothing
        Set objCollection = Nothing

    Case 901
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        dHorizontalOffset = 0.1 * parHeight

        Set objCollection = New Collection
        oStpoint.Set flangeThick, 0, 0
        oEnpoint.Set flangeThick, pipeDiam / 2, 0
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Line 1
        oStpoint.Set flangeThick, pipeDiam / 2, 0
        oEnpoint.Set flangeThick + dHorizontalOffset, pipeDiam / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Elliptical Arc 1
        oCenter.Set flangeThick + dHorizontalOffset, pipeDiam / 2 + (parDiameter - pipeDiam) / 4, oStpoint.z
        dMMRatio = ((parDiameter - pipeDiam) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, 1, _
                                                                                       1.6 * dHorizontalOffset, 0, 0, dMMRatio, -PI / 2, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Elliptical Arc 2
        oCenter.Set flangeThick + 4.2 * dHorizontalOffset, parDiameter / 2 - ((parDiameter - pipeDiam) / 4), oCenter.z
        dMMRatio = ((parDiameter - pipeDiam) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, -1, _
                                                                                       -1.6 * dHorizontalOffset, 0, 0, dMMRatio, 0, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Line 1
        oStpoint.Set flangeThick + 4.2 * dHorizontalOffset, parDiameter / 2, oCenter.z
        oEnpoint.Set 0.9 * parHeight, parDiameter / 2, oCenter.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Create small circular Arc
        'Arc 1
        oStpoint.Set 0.9 * parHeight, parDiameter / 2, oCenter.z
        oEnpoint.Set parHeight, (parDiameter / 2) - dHorizontalOffset, oStpoint.z
        oCenter.Set oStpoint.x, oEnpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'St Line
        oStpoint.Set parHeight, (parDiameter / 2) - dHorizontalOffset, oCenter.z
        oEnpoint.Set parHeight, 0, oCenter.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        oStpoint.Set flangeThick, 0, oCenter.z
        Set oAxisVec = New DVector
        oAxisVec.Set -1, 0, 0
        Set oComplexStr = PlaceTrCString(oStpoint, objCollection)
        oStpoint.Set flangeThick, 0, oStpoint.z
        Set objCurve = PlaceRevolution(m_OutputColl, oComplexStr, oAxisVec, oStpoint, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurve
        Set objCurve = Nothing
        Set oAxisVec = Nothing
        Set oComplexStr = Nothing
        Set objCollection = Nothing

        oStpoint.Set flangeThick + 4.2 * dHorizontalOffset, -parDiameter / 2 + (1.4 * pipeDiam3 / 2), 0
        oEnpoint.Set parDrainPortXOffset + (0.7 * pipeDiam3), oStpoint.y, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, 1.4 * pipeDiam3, True)
        'Set the output
        m_OutputColl.AddOutput "ObjCylinder", objCylinder
        Set objCylinder = Nothing
        
        'Place elbow
        Dim objCircle As IngrGeom3D.Circle3d
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                          parDrainPortXOffset, -parDiameter / 2, 0, _
                                                                          0, -1, 0, _
                                                                          pipeDiam3 / 2)

        Dim CP As New AutoMath.DPosition    'revolution center point
        Dim CV As New AutoMath.DVector    'rotation vector for rotation
        CP.Set parDrainPortXOffset + (0.7 * pipeDiam3), -parDiameter / 2, 0
        CV.Set 0, 0, -1

        Dim ObjElbow As Object
        Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                                       CV, CP, PI / 2, True)

        'Set the output
        m_OutputColl.AddOutput "Elbow", ObjElbow
        Set ObjElbow = Nothing
        Set CP = Nothing
        Set CV = Nothing
        Set objCircle = Nothing

    Case 902
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        parDrainPortYOffset = arrayOfInputs(10)
        dHorizontalOffset = 0.1 * parHeight

        Set objCollection = New Collection
        'line 1
        oStpoint.Set flangeThick, 0, oStpoint.z
        oEnpoint.Set flangeThick, pipeDiam / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'line 2
        oStpoint.Set flangeThick, pipeDiam / 2, 0
        oEnpoint.Set 2 * dHorizontalOffset, pipeDiam / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Arc 1
        oStpoint.Set 2 * dHorizontalOffset, pipeDiam / 2, oStpoint.z
        oEnpoint.Set 4 * dHorizontalOffset, parDiameter / 2, oStpoint.z
        oCenter.Set oEnpoint.x, oStpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'line 2
        oStpoint.Set 4 * dHorizontalOffset, parDiameter / 2, oStpoint.z
        oEnpoint.Set 9 * dHorizontalOffset, parDiameter / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Arc 2
        oStpoint.Set 9 * dHorizontalOffset, parDiameter / 2, oStpoint.z
        oEnpoint.Set parHeight, (parDiameter / 2) - dHorizontalOffset, oStpoint.z
        oCenter.Set oStpoint.x, oEnpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'line 3
        oStpoint.Set parHeight, (parDiameter / 2) - dHorizontalOffset, oStpoint.z
        oEnpoint.Set parHeight, 0, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        oStpoint.Set flangeThick, 0, oStpoint.z
        Set oAxisVec = New DVector
        oAxisVec.Set -1, 0, 0
        Set oComplexStr = PlaceTrCString(oStpoint, objCollection)
        oStpoint.Set flangeThick, 0, 0
        Set objCurve = PlaceRevolution(m_OutputColl, oComplexStr, oAxisVec, oStpoint, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCurve
        Set objCurve = Nothing
        Set oAxisVec = Nothing
        Set oComplexStr = Nothing
        Set objCollection = Nothing

    Case 903
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        parDrainPortYOffset = arrayOfInputs(10)
        parInletPortXOffset = arrayOfInputs(11)
        
        oStpoint.Set parInletPortXOffset, 0, 0
        oEnpoint.Set parHeight, 0, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, parDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
        Set objCylinder = Nothing

    Case Else
        GoTo ErrorLabel:    'Invalid specification
    End Select
    Set oStpoint = Nothing
    Set oEnpoint = Nothing
    Set oGeomFactory = Nothing
    Set oCenter = Nothing
    
    If lPartDataBasis = 900 Or lPartDataBasis = 901 Or _
                lPartDataBasis = 902 Or lPartDataBasis = 903 Then
        ' Place Nozzle 1
        Dim dNozzle1Length As Double
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        oPlacePoint.Set -sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        If lPartDataBasis = 903 Then
            dNozzle1Length = parInletPortXOffset
        Else
            dNozzle1Length = flangeThick
        End If
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzle1Length)
    
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        ' Place Nozzle 2(Exhaust port)
        oPlacePoint.Set parHeight + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        ' Place Nozzle 3(Drain port)
        Dim dNozzle3Length As Double
        If lPartDataBasis = 900 Then
            oPlacePoint.Set parDrainPortXOffset - sptOffset + depth, -0.5 * (parDiameter - pipeDiam3), 0
            dNozzle3Length = 2 * parDrainPortXOffset
            oDir.Set -1, 0, 0
        ElseIf lPartDataBasis = 901 Then
            oPlacePoint.Set parDrainPortXOffset, -0.5 * parDiameter - sptOffset + depth, 0
            dNozzle3Length = flangeThick3
            oDir.Set 0, -1, 0
        ElseIf lPartDataBasis = 902 Then
            oPlacePoint.Set parDrainPortXOffset - (sptOffset - depth) * Cos(PI / 4), -parDrainPortYOffset - (sptOffset - depth) * Sin(PI / 4), 0
            oDir.Set -Cos(PI / 4), -Sin(PI / 4), 0
            'Calculate nozzle 3 length
            Dim a, b, c As Double
            a = Sqr(parDrainPortXOffset ^ 2 + parDrainPortYOffset ^ 2)
            b = 0.5 * pipeDiam / Sin(PI / 4)
            c = Sqr(((0.5 * parDiameter - 0.5 * pipeDiam) ^ 2) - (pipeDiam3 * 0.5) ^ 2)
            dNozzle3Length = a - (b + c)
        ElseIf lPartDataBasis = 903 Then
            oPlacePoint.Set parDrainPortXOffset, parDrainPortYOffset, 0
            dNozzle3Length = parInletPortXOffset - parDrainPortXOffset
            oDir.Set -1, 0, 0
        End If
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzle3Length)
        'Set the output
        m_OutputColl.AddOutput "DrainPort", objNozzle
        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
  End If
  
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

