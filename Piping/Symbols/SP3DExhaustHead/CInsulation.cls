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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          Babu Govindarajan
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the catalog available at URL http://www.nciweb.net/exhaust.htm
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   22.Aug.2005     Sundar(svsmylav)        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'   30.Jan.2006     Sundar(svsmylav)        RI-28367: Deleted reference to 'SN=S7AZ' of PDS in the
'                                           revision history and added hyper link to Anderson's site.
'                                           Also, modified symbol input parameters to match that of the
'                                           manufacturer's skecth.
'  24.Apr.2008      VRK                     CR-136579: Enhance exhaust head symbol
'                                            to provide more realistic dimensional basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:"    'Used for error messages
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

    Dim ObjExConeIns As Object
    Dim ObjNoz1Ins As Object
    Dim ObjPipeIns As Object
    Dim ObjNoz2Ins As Object

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
    Dim dInsulationDia As Double
    Dim dInsDia As Double
    Dim dInsulationDia3 As Double


    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '    parExhaustDiameter = arrayOfInputs(2) 'A
    '    parExhaustHeight = arrayOfInputs(3)   'B
    '    parPortVOffset = arrayOfInputs(4)     'C
    '    parPortHOffset = arrayOfInputs(5)     'E
        parInsulationThickness = arrayOfInputs(6)
    '    parDiameter = arrayOfInputs(7)
    '    parHeight = arrayOfInputs(8)
    '    parDrainPortXOffset =arrayOfInputs(9)
    '    parDrainPortYOffset =arrayOfInputs(10)
    '    parInletPortXOffset=arrayOfInputs(11)


    iOutput = 0

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
    Case Is <= 1
        parExhaustDiameter = arrayOfInputs(2) 'A
        parExhaustHeight = arrayOfInputs(3)   'B
        parPortVOffset = arrayOfInputs(4)     'C
        parPortHOffset = arrayOfInputs(5)     'E

        If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
        Dim baseCenter As AutoMath.DPosition
        Dim topCenter As AutoMath.DPosition
        Set baseCenter = New AutoMath.DPosition
        Set topCenter = New AutoMath.DPosition

        baseCenter.Set 0, flangeThick, 0
        topCenter.Set 0, parExhaustHeight + parInsulationThickness, 0
        Set ObjExConeIns = PlaceCone(m_OutputColl, baseCenter, topCenter, pipeDiam / 2 + parInsulationThickness, _
                                     parExhaustDiameter / 2 + parInsulationThickness, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExConeIns
        Set ObjExConeIns = Nothing
        Set baseCenter = Nothing
        Set topCenter = Nothing

        ' Insert your code for output 5(Nozzle 1 Insulation)
        oStpoint.Set 0, 0, 0
        oEnpoint.Set 0, flangeThick + parInsulationThickness, 0
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, dInsulationDia, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
        Set ObjNoz1Ins = Nothing

        ' Insert your code for output 6(Down Pipe Insulation)
        If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
        oStpoint.Set parPortHOffset, parExhaustHeight + parInsulationThickness, 0
        oEnpoint.Set parPortHOffset, parPortVOffset + flangeThick, 0
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        Set ObjPipeIns = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, dInsulationDia, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeIns
        Set ObjPipeIns = Nothing

        ' Insert your code for output 7(Place Nozzle 2 Insulation)
        oStpoint.Set parPortHOffset, parPortVOffset, 0
        oEnpoint.Set parPortHOffset, parPortVOffset + flangeThick + parInsulationThickness, 0
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Set ObjNoz2Ins = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, dInsulationDia, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2Ins
        Set ObjNoz2Ins = Nothing

    Case 900
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        dHorizontalOffset = 0.1 * parHeight
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        dInsDia = parDiameter + parInsulationThickness * 2
        Set objCollection = New Collection
        
        oStpoint.Set flangeThick, 0, 0
        oEnpoint.Set flangeThick, dInsulationDia / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Elliptical Arc 1
        oCenter.Set flangeThick, dInsulationDia / 2 + (dInsDia - dInsulationDia) / 4, 0
        dMMRatio = ((dInsDia - dInsulationDia) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, 1, _
                                                                                       1.6 * dHorizontalOffset, 0, 0, dMMRatio, -PI / 2, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Elliptical Arc 2
        oCenter.Set flangeThick + 3.2 * dHorizontalOffset, dInsDia / 2 - ((dInsDia - dInsulationDia) / 4), oCenter.z
        dMMRatio = ((dInsDia - dInsulationDia) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, -1, _
                                                                                       -1.6 * dHorizontalOffset, 0, 0, dMMRatio, 0, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Line 1
        oStpoint.Set flangeThick + 3.2 * dHorizontalOffset, dInsDia / 2, oCenter.z
        oEnpoint.Set 0.9 * parHeight, dInsDia / 2, oCenter.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Create small circular Arc
        'Arc 1
        Dim oCurve As IngrGeom3D.Arc3d
        oStpoint.Set 0.9 * parHeight, dInsDia / 2, oCenter.z
        oEnpoint.Set parHeight, (dInsDia / 2) - dHorizontalOffset, oStpoint.z
        oCenter.Set oStpoint.x, oEnpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'St Line
        oStpoint.Set parHeight, (dInsDia / 2) - dHorizontalOffset, oCenter.z
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
               
        'Nozzle Insulation:
        oStpoint.Set 0, 0, 0
        If CmpDblEqual(flangeThick, 0) Then
            oEnpoint.Set NEGLIGIBLE_THICKNESS, 0, 0
        Else
            oEnpoint.Set flangeThick + parInsulationThickness, 0, 0
        End If
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness
        Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, _
                                                        dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
        Set ObjNoz1Ins = Nothing
             
        dInsulationDia3 = flangeDiam3 + parInsulationThickness * 2
        If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then dInsulationDia3 = pipeDiam3 + 2 * parInsulationThickness
        oStpoint.Set parDrainPortXOffset, -0.5 * (parDiameter - pipeDiam3), 0
        oEnpoint.Set 3 * parDrainPortXOffset, oStpoint.y, oStpoint.z
        Set ObjPipeIns = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, _
                                                        dInsulationDia3, True)

        'Set the output
        m_OutputColl.AddOutput "InsDrainBody", ObjPipeIns
        Set ObjPipeIns = Nothing

    Case 901
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        dHorizontalOffset = 0.1 * parHeight
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        dInsDia = parDiameter + parInsulationThickness * 2

        Set objCollection = New Collection
        oStpoint.Set flangeThick, 0, 0
        oEnpoint.Set flangeThick, dInsulationDia / 2, 0
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Line 1
        oStpoint.Set flangeThick, dInsulationDia / 2, 0
        oEnpoint.Set flangeThick + dHorizontalOffset, dInsulationDia / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Elliptical Arc 1
        oCenter.Set flangeThick + dHorizontalOffset, dInsulationDia / 2 + (dInsDia - dInsulationDia) / 4, oStpoint.z
        dMMRatio = ((dInsDia - dInsulationDia) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, 1, _
                                                                                       1.6 * dHorizontalOffset, 0, 0, dMMRatio, -PI / 2, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Elliptical Arc 2
        oCenter.Set flangeThick + 4.2 * dHorizontalOffset, dInsDia / 2 - ((dInsDia - dInsulationDia) / 4), oCenter.z
        dMMRatio = ((dInsDia - dInsulationDia) / 4) / (1.6 * dHorizontalOffset)
        Set oArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                                       oCenter.x, oCenter.y, oCenter.z, 0, 0, -1, _
                                                                                       -1.6 * dHorizontalOffset, 0, 0, dMMRatio, 0, PI / 2)
        objCollection.Add oArc
        Set oArc = Nothing

        'Line 1
        oStpoint.Set flangeThick + 4.2 * dHorizontalOffset, dInsDia / 2, oCenter.z
        oEnpoint.Set 0.9 * parHeight, dInsDia / 2, oCenter.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Create small circular Arc
        'Arc 1
        oStpoint.Set 0.9 * parHeight, dInsDia / 2, oCenter.z
        oEnpoint.Set parHeight, (dInsDia / 2) - dHorizontalOffset, oStpoint.z
        oCenter.Set oStpoint.x, oEnpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'St Line
        oStpoint.Set parHeight, (dInsDia / 2) - dHorizontalOffset, oCenter.z
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
              
        'Nozzle Insulation:
        oStpoint.Set 0, 0, 0
        If CmpDblEqual(flangeThick, 0) Then
            oEnpoint.Set NEGLIGIBLE_THICKNESS, 0, 0
        Else
            oEnpoint.Set flangeThick + parInsulationThickness, 0, 0
        End If
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness
        Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, _
                                                        dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
        Set ObjNoz1Ins = Nothing
        
        dInsulationDia3 = pipeDiam3 + parInsulationThickness * 2
        oStpoint.Set flangeThick + 4.2 * dHorizontalOffset, -parDiameter / 2 + (1.4 * pipeDiam3 / 2), 0
        oEnpoint.Set parDrainPortXOffset + (0.7 * pipeDiam3), oStpoint.y, 0
        Set objCylinder = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, 1.4 * dInsulationDia3, True)
        'Set the output
        m_OutputColl.AddOutput "ObjInsCylinder", objCylinder
        Set objCylinder = Nothing
        
        'Place elbow
        Dim objCircle As IngrGeom3D.Circle3d
        Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                          parDrainPortXOffset, -parDiameter / 2, 0, _
                                                                          0, -1, 0, _
                                                                          dInsulationDia3 / 2)

        Dim CP As New AutoMath.DPosition    'revolution center point
        Dim CV As New AutoMath.DVector    'rotation vector for rotation
        CP.Set parDrainPortXOffset + (0.7 * pipeDiam3), -parDiameter / 2, 0
        CV.Set 0, 0, -1

        Dim ObjElbow As Object
        Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                                       CV, CP, PI / 2, True)

        'Set the output
        m_OutputColl.AddOutput "InsElbow", ObjElbow
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
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        dInsDia = parDiameter + parInsulationThickness * 2

        Set objCollection = New Collection
        'line 1
        oStpoint.Set flangeThick, 0, oStpoint.z
        oEnpoint.Set flangeThick, dInsulationDia / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'line 2
        oStpoint.Set flangeThick, dInsulationDia / 2, 0
        oEnpoint.Set 2 * dHorizontalOffset, dInsulationDia / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Arc 1
        oStpoint.Set 2 * dHorizontalOffset, dInsulationDia / 2, oStpoint.z
        oEnpoint.Set 4 * dHorizontalOffset, dInsDia / 2, oStpoint.z
        oCenter.Set oEnpoint.x, oStpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'line 2
        oStpoint.Set 4 * dHorizontalOffset, dInsDia / 2, oStpoint.z
        oEnpoint.Set 9 * dHorizontalOffset, dInsDia / 2, oStpoint.z
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         oStpoint.x, oStpoint.y, oStpoint.z, _
                                                         oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oLine
        Set oLine = Nothing

        'Arc 2
        oStpoint.Set 9 * dHorizontalOffset, dInsDia / 2, oStpoint.z
        oEnpoint.Set parHeight, (dInsDia / 2) - dHorizontalOffset, oStpoint.z
        oCenter.Set oStpoint.x, oEnpoint.y, oEnpoint.z
        Set oCurve = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                                oCenter.x, oCenter.y, oCenter.z, _
                                                                oStpoint.x, oStpoint.y, oStpoint.z, _
                                                                oEnpoint.x, oEnpoint.y, oEnpoint.z)
        objCollection.Add oCurve
        Set oCurve = Nothing

        'line 3
        oStpoint.Set parHeight, (dInsDia / 2) - dHorizontalOffset, oStpoint.z
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
        
       'Nozzle Insulation:
        oStpoint.Set 0, 0, 0
        If CmpDblEqual(flangeThick, 0) Then
            oEnpoint.Set NEGLIGIBLE_THICKNESS, 0, 0
        Else
            oEnpoint.Set flangeThick + parInsulationThickness, 0, 0
        End If
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness
        Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, _
                                                        dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
        Set ObjNoz1Ins = Nothing
        
        'Drain port insulation
        Dim a, b, c As Double
        Dim dCylLen As Double
        a = Sqr(parDrainPortXOffset ^ 2 + parDrainPortYOffset ^ 2)
        b = 0.5 * pipeDiam / Sin(PI / 4)
        c = Sqr(((0.5 * parDiameter - 0.5 * pipeDiam) ^ 2) - (pipeDiam3 * 0.5) ^ 2)
        dCylLen = a - (b + c)
        dInsulationDia3 = flangeDiam3 + parInsulationThickness * 2
        If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then dInsulationDia3 = pipeDiam3 + 2 * parInsulationThickness
        oStpoint.Set parDrainPortXOffset, -parDrainPortYOffset, 0
        oEnpoint.Set parDrainPortXOffset + dCylLen * Cos(PI / 4), -parDrainPortYOffset + dCylLen * Sin(PI / 4), oStpoint.z
        Set ObjPipeIns = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, _
                                                        dInsulationDia3, True)

        'Set the output
        m_OutputColl.AddOutput "InsDrainBody", ObjPipeIns
        Set ObjPipeIns = Nothing
        
   Case 903
        parDiameter = arrayOfInputs(7)
        parHeight = arrayOfInputs(8)
        parDrainPortXOffset = arrayOfInputs(9)
        parDrainPortYOffset = arrayOfInputs(10)
        parInletPortXOffset = arrayOfInputs(11)
        dInsulationDia = pipeDiam + parInsulationThickness * 2
        dInsDia = parDiameter + parInsulationThickness * 2

        Dim objInsCylinder As Object
        oStpoint.Set parInletPortXOffset, 0, 0
        oEnpoint.Set parHeight, 0, 0
        Set objInsCylinder = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, dInsDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
        Set objInsCylinder = Nothing
        
        'Nozzle insulation
        oStpoint.Set 0, 0, 0
        If CmpDblEqual(flangeThick, 0) Then
            oEnpoint.Set NEGLIGIBLE_THICKNESS, 0, 0
        Else
            oEnpoint.Set flangeThick + parInsulationThickness, 0, 0
        End If
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness
        Set ObjNoz1Ins = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, _
                                                        dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz1Ins
        Set ObjNoz1Ins = Nothing
        
        oStpoint.Set 0, 0, 0
        oEnpoint.Set parInletPortXOffset, 0, 0
        dInsulationDia = flangeDiam + parInsulationThickness * 2
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + 2 * parInsulationThickness
        Set objInsCylinder = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, dInsulationDia, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
        Set objInsCylinder = Nothing
        
        dInsulationDia3 = flangeDiam3 + parInsulationThickness * 2
        If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then dInsulationDia3 = pipeDiam3 + 2 * parInsulationThickness
        oStpoint.Set parDrainPortXOffset, parDrainPortYOffset, 0
        oEnpoint.Set parInletPortXOffset, parDrainPortYOffset, 0
        Set objInsCylinder = PlaceCylinder(m_OutputColl, oStpoint, oEnpoint, dInsulationDia3, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
        Set objInsCylinder = Nothing
      
    Case Else
        GoTo ErrorLabel:    'Invalid specification
    End Select
        Set oStpoint = Nothing
        Set oEnpoint = Nothing
        Set oGeomFactory = Nothing
        Set oCenter = Nothing
        
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
