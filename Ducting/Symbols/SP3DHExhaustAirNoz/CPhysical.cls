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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Monday, Dec 03 2007
'   Description:
'   This symbol is created for CR-129442 (Exhaust Air Nozzle)
'   Source: Toilet Exhaust Air Nozzle.pdf.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parWidth As Double
    Dim parBodyDiameter As Double
    Dim parBodyLength As Double
    Dim parRingDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBodyDiameter = arrayOfInputs(3)
    parBodyLength = arrayOfInputs(4)
    parRingDiameter = arrayOfInputs(5)
    
    'Origin is taken at the Port
    'For Horizontally routed Duct X - Along the flow and Y - towards UP
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for Output 1 (Body)
    Dim dLineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim objBody As Object
    
    'Point 1
    dLineStrPoints(0) = 0
    dLineStrPoints(1) = 0
    dLineStrPoints(2) = 0
    
    'Point 2
    dLineStrPoints(3) = dLineStrPoints(0)
    dLineStrPoints(4) = parWidth / 2
    dLineStrPoints(5) = dLineStrPoints(2)
    
    'Point 3
    dLineStrPoints(6) = 0.5 * parBodyLength
    dLineStrPoints(7) = dLineStrPoints(4)
    dLineStrPoints(8) = dLineStrPoints(2)
    
    'Point 4
    dLineStrPoints(9) = dLineStrPoints(6)
    dLineStrPoints(10) = parBodyDiameter / 2
    dLineStrPoints(11) = dLineStrPoints(2)
    
    'Point 5
    dLineStrPoints(12) = 0.6 * parBodyLength
    dLineStrPoints(13) = dLineStrPoints(10)
    dLineStrPoints(14) = dLineStrPoints(2)

    'Point 6
    dLineStrPoints(15) = dLineStrPoints(12)
    dLineStrPoints(16) = dLineStrPoints(1)
    dLineStrPoints(17) = dLineStrPoints(2)
    
    'Point 7
    dLineStrPoints(18) = dLineStrPoints(0)
    dLineStrPoints(19) = dLineStrPoints(1)
    dLineStrPoints(20) = dLineStrPoints(2)
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                            7, dLineStrPoints)
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 1, 0, 0
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0, -0.00001, 0
  
    Set objBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                        oCenPoint, 2 * PI, True)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
    Set objBody = Nothing
    Set oAxisVect = Nothing
    Set oCenPoint = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing

    'Insert your code for Output 2 (Cone)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set 0.1 * parBodyLength, 0, 0
    oEnPoint.Set 0.5 * parBodyLength, 0, 0
    
    Dim objCone As Object
    Set objCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, parRingDiameter / 2, _
                                                         parWidth / 2, True)
   
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Insert your code for Output 3 (RingBody)
    'Assumption : The height of the cylindrical portion is 10% of the Body Length
    'Points specifying the B-Spline curve are approximated to produce closer geometry.
    Dim oPoints(1 To 5) As New AutoMath.DPosition
    oPoints(1).Set 0.15 * parBodyLength, 0.5 * parRingDiameter, 0
    oPoints(2).Set 0.25 * parBodyLength, 0.5 * parRingDiameter, 0
    oPoints(3).Set 0.26 * parBodyLength, 0.4 * parRingDiameter, 0
    oPoints(4).Set 0.28 * parBodyLength, 0.35 * parRingDiameter, 0
    oPoints(5).Set 0.3 * parBodyLength, 0, 0
  
    Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
    Set objBsplineCurve = PlaceTrBspline(4, oPoints)
    
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 1, 0, 0
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0, 0, 0
   
    Dim objRingBody As Object
    Set objRingBody = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVect, _
                                                oCenPoint, 2 * PI, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRingBody
    Set objRingBody = Nothing
    Set oCenPoint = Nothing
    Set oAxisVect = Nothing
    Set objBsplineCurve = Nothing

    '-------------------------------------------------------------------------------------------------
    '==================================================
    'BUILD HVACNOZZLE 1
    '==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1) As Long
    Dim dThickness(1) As Double
    Dim dFlangeWidth(1) As Double
    Dim lFlowDir(1) As Long
    Dim dPortDepth(1) As Double
    Dim dCptOffSet(1) As Double
    Dim dNozzLength(1) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = NEGLIGIBLE_THICKNESS
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + NEGLIGIBLE_THICKNESS
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Dim dDepth As Double
    If CmpDblEqual(GSCADNozzleEntities.Round, 4) Then dDepth = parWidth
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Round, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)

    'Position of the Nozzle
    Dim x As Double, y As Double, z As Double
    x = 0#
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir
    
    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
