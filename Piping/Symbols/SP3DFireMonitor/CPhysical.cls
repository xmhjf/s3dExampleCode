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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Monday, Jun 10 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: FS1A Page No. D-87 PDS Piping Component Data Reference Guide.
' SP3D3WFireMonitor Symbol is created with Three Outputs
' The three physical aspect outputs are created as follows:
'   ObjMonitorBody-Using 'PlaceBox; function,
'   ObjCannonBody- Using 'PlaceCylinder' function and
'   One ObjNozzle object by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   19.Dec.2003      svsmylav        TR-52639: Added logic to ensure minimum length of Nozzle
'                                           to be equal to flange Thickness (to avoid failure of nozzle).
'   28.May.2007      MA                   CR-113431: Implemented new part data basis '375' for SN=FS113B, H2WN3
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
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjMonitorBody As Object
    Dim ObjCannonBody As Object
    
    Dim parFacetoCenter As Double
    Dim parBodyTtoCenter As Double
    Dim parBodyBtoCenter As Double
    Dim parBodyLtoCenter As Double
    Dim parBodyRtoCenter As Double
    Dim parBodyWidth As Double
    Dim parCannonLength As Double
    Dim parCannonDiameter As Double
    Dim parHeight As Double
    Dim parDiameter As Double
    Dim parLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)          'D1
'    parBodyTtoCenter = arrayOfInputs(3)     'D2
'    parBodyBtoCenter = arrayOfInputs(4)     'D3
'    parBodyLtoCenter = arrayOfInputs(5)     'D4
'    parBodyRtoCenter = arrayOfInputs(6)     'D5
'    parBodyWidth = arrayOfInputs(7)             'D6
'    parCannonLength = arrayOfInputs(8)      'D7
'    parCannonDiameter = arrayOfInputs(9)    'D8
'    Above D1 to D8 are from SN=FS1A; FOR FS113B, parFacetoCenter is 'Par 3'
'    and other dimensions are as follows:
'    parHeight = arrayOfInputs(10)        'Par 4
'    parDiameter = arrayOfInputs(11)      'Par 5
'    parLength = arrayOfInputs(12)        'Par 6
    
    iOutput = 0
    
    Dim dNozzleLength As Double
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Dim oCenPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
    
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
'Prepare transformation matrix and get operator direction
    Dim oTransMatrix As IJDT4x4
    Dim oVector As AutoMath.DVector

    Set oVector = New AutoMath.DVector
    Set oTransMatrix = New DT4x4

    oTransMatrix.LoadIdentity
    oVector.Set 0, 0, 1
    oTransMatrix.Rotate PI / 2, oVector
    
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 347
        parBodyTtoCenter = arrayOfInputs(3)
        parBodyBtoCenter = arrayOfInputs(4)
        parBodyLtoCenter = arrayOfInputs(5)
        parBodyRtoCenter = arrayOfInputs(6)
        parBodyWidth = arrayOfInputs(7)
        parCannonLength = arrayOfInputs(8)
        parCannonDiameter = arrayOfInputs(9)
        
'       Compute nozzle length
        dNozzleLength = parFacetoCenter - parBodyBtoCenter

        ' Insert your code for output 1(Body of Monitor)
        oStPoint.Set parBodyRtoCenter, -parBodyWidth / 2, parBodyTtoCenter
        oEnPoint.Set -parBodyLtoCenter, parBodyWidth / 2, -parBodyBtoCenter
            
        Set ObjMonitorBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMonitorBody
        Set ObjMonitorBody = Nothing
        
        ' Insert your code for output 2(Body of Cannon)
        Dim CannonBodyAngle As Double
        CannonBodyAngle = PI / 4
    
        oStPoint.Set 0, 0, 0
        oEnPoint.Set parCannonLength * Cos(CannonBodyAngle), 0, parCannonLength * Sin(CannonBodyAngle)
        
        Set ObjCannonBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parCannonDiameter, True)
        ObjCannonBody.Transform oTransMatrix
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCannonBody
        Set ObjCannonBody = Nothing
        
     Case 375
'       Implemented new part data basis '375' for SN=FS113B, H2WN3
        parCannonLength = parFacetoCenter 'This is 'Par 3' in the sketch
        parCannonDiameter = arrayOfInputs(9) 'Par 7
        parHeight = arrayOfInputs(10)        'Par 4
        parDiameter = arrayOfInputs(11)      'Par 5
        parLength = arrayOfInputs(12)        'Par 6
        
        ' Insert your code for output 1(Body of Monitor)
        Dim objMonitor As Object
        Dim oAxisVec As New AutoMath.DVector
        Dim oCenter As New AutoMath.DPosition
        
        Dim oLine1   As IngrGeom3D.Line3d
        Dim oLine2   As IngrGeom3D.Line3d
        Dim oArc1  As IngrGeom3D.Arc3d
        
        'Line1
        oStPoint.Set -0.0001, 0, -parFacetoCenter
        oEnPoint.Set -0.5 * parDiameter - 0.0001, 0, -parFacetoCenter
        Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
        'Line2
        oStPoint.Set -0.5 * parDiameter - 0.0001, 0, -parFacetoCenter
        oEnPoint.Set -0.5 * parDiameter - 0.0001, 0, 0
        Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
        'Arc1
        oStPoint.Set -0.5 * parDiameter - 0.0001, 0, 0
        oEnPoint.Set -0.0001, 0, parHeight
        oCenPoint.Set 0, 0, 0
    
        Set oArc1 = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
        
        Dim oCurve         As IngrGeom3D.ComplexString3d
        Dim oCurveCol      As Collection
        Set oCurveCol = New Collection
        oCurveCol.Add oLine1
        oCurveCol.Add oLine2
        oCurveCol.Add oArc1
        oStPoint.Set -0.0001, 0, -parFacetoCenter
        
        Set oCurve = PlaceTrCString(oStPoint, oCurveCol)
        
        oAxisVec.Set 0, 0, 1
        oCenter.Set 0.0001, 0, -parFacetoCenter
        Set objMonitor = PlaceRevolution(m_OutputColl, oCurve, oAxisVec, oCenter, 2 * PI, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "MonitorBody", objMonitor
        Set objMonitor = Nothing
        Set oAxisVec = Nothing
        Set oCenter = Nothing
        Set oLine1 = Nothing
        Set oLine2 = Nothing
        Set oArc1 = Nothing
        Set oCurve = Nothing
        Set oCurveCol = Nothing
        
        ' Insert your code for output 2(Body of Cannon)
        Dim oHorLine   As IngrGeom3D.Line3d
        Dim oArc  As IngrGeom3D.Arc3d
        Dim oVerLine   As IngrGeom3D.Line3d
        
        'Horizontal line
        oStPoint.Set 0, 0, 0.3 * parCannonLength - parFacetoCenter
        oEnPoint.Set 0, parLength, 0.3 * parCannonLength - parFacetoCenter
        
        Set oHorLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
        'Arc
        oStPoint.Set 0, parLength, 0.3 * parCannonLength - parFacetoCenter
        oEnPoint.Set 0, parLength + (1.5 * parCannonDiameter), 0.6 * parCannonLength - parFacetoCenter
        oCenPoint.Set 0, parLength, 0.6 * parCannonLength - parFacetoCenter
    
        Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
        
        'Vertical line
        oStPoint.Set 0, parLength + (1.5 * parCannonDiameter), 0.6 * parCannonLength - parFacetoCenter
        oEnPoint.Set 0, parLength + (1.5 * parCannonDiameter), _
        ((0.6 * parCannonLength) + (0.25 * parDiameter) + (2 * parDiameter)) - parFacetoCenter
        
        Set oVerLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
        
        Dim oSweepCurve         As IngrGeom3D.ComplexString3d
        Dim oSweepCurveCol      As Collection
        Set oSweepCurveCol = New Collection
        oSweepCurveCol.Add oHorLine
        oSweepCurveCol.Add oArc
        oSweepCurveCol.Add oVerLine

        Dim oStartBC   As AutoMath.DPosition
        Set oStartBC = New AutoMath.DPosition
        oStartBC.Set 0, 0, 0.3 * parCannonLength - parFacetoCenter
        
        Set oSweepCurve = PlaceTrCString(oStartBC, oSweepCurveCol)
        Dim oDirProj As AutoMath.DVector
        Set oDirProj = New AutoMath.DVector
        
        Dim oCircle As IngrGeom3D.Circle3d
        
        Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                        0, 0, 0.3 * parCannonLength - parFacetoCenter, _
                        0, 1, 0, _
                        0.5 * parCannonDiameter)
        
        Dim Surfset   As IngrGeom3D.IJElements
        Dim stnorm() As Double
        Dim ednorm() As Double
        Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                         oSweepCurve, oCircle, CircularCorner, 0, stnorm, ednorm, False)
    
        Dim ObjSurface As Object
        For Each ObjSurface In Surfset
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
        Next ObjSurface
        
        Set oHorLine = Nothing
        Set oArc = Nothing
        Set oVerLine = Nothing
        Set oSweepCurve = Nothing
        Set oSweepCurveCol = Nothing
        Set oStartBC = Nothing
        Set oDirProj = Nothing
        Set oCircle = Nothing
        Set Surfset = Nothing
        Set oGeomFactory = Nothing

        'Insert your code for Cylinder 1
        Dim objCyl1 As Object
        oStPoint.Set 0, parLength + (1.5 * parCannonDiameter), 0.6 * parCannonLength - parFacetoCenter
        oEnPoint.Set 0, parLength + (1.5 * parCannonDiameter), _
            ((0.6 * parCannonLength) + (0.25 * parDiameter)) - parFacetoCenter
    
        Set objCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter, True)
        
        ' Set the output
'        iOutput = iOutput + 1
'        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl1

        m_OutputColl.AddOutput "CannonBody", objCyl1
        Set objCyl1 = Nothing
    
        'Insert your code for Sphere
        Dim objSphere As Object
        Dim oCentPoint   As AutoMath.DPosition
        Set oCentPoint = New AutoMath.DPosition
        oCentPoint.Set 0, parLength + (1.5 * parCannonDiameter), _
        ((0.6 * parCannonLength) + (0.25 * parDiameter) + (2 * parDiameter)) - parFacetoCenter
    
        Set objSphere = PlaceSphere(m_OutputColl, oCentPoint, 1.5 * parCannonDiameter)
    
        ' Set the output
        m_OutputColl.AddOutput "CannonBody", objSphere
        Set objSphere = Nothing
        Set oCentPoint = Nothing

        'Insert your code for Cylinder 2
        Dim objCyl2 As Object
        oStPoint.Set 0, parLength + (1.5 * parCannonDiameter), _
        ((0.6 * parCannonLength) + (0.25 * parDiameter) + (2 * parDiameter)) - parFacetoCenter
        oEnPoint.Set 0, parLength + (1.5 * parCannonDiameter) + (parCannonLength * Cos(PI / 4)), _
        ((0.6 * parCannonLength) + (0.25 * parDiameter) + _
            (2 * parDiameter) + (parCannonLength * Sin(PI / 4))) - parFacetoCenter
    
        Set objCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parCannonDiameter, True)
        
        ' Set the output
        m_OutputColl.AddOutput "CannonBody", objCyl2
        Set objCyl2 = Nothing

        'Insert your code for Cylinder 3
        Dim objCyl3 As Object
        oStPoint.Set 0, parLength + (1.5 * parCannonDiameter), _
        ((0.6 * parCannonLength) + (0.25 * parDiameter) + (2 * parDiameter)) - parFacetoCenter
        oEnPoint.Set 0, -0.5 * parDiameter, _
        ((0.6 * parCannonLength) + (0.25 * parDiameter) + (2 * parDiameter)) - parFacetoCenter
    
        Set objCyl3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.25 * parDiameter, True)
        
        ' Set the output
        m_OutputColl.AddOutput "CannonBody", objCyl3
        Set objCyl3 = Nothing
        
        Set oPipeComponent = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oCenPoint = Nothing
        
    Case Else
        GoTo ErrorLabel:
        
    End Select

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set 0, 0, -parFacetoCenter - sptOffset + depth
    oDir.Set 0, 0, -1

    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oVector = Nothing
    Set oTransMatrix = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
