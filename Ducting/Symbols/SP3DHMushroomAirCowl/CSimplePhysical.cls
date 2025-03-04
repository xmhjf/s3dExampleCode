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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   File            :  CSimplePhysical.cls
'   Author          :  PK
'   Creation Date   :  Monday,August 6 2007
'   Description     :  Mushroom Type Air Cowl
'   Reference       :  STD-M1.pdf
'   CR number       :  CR-124217
'   Change History  :
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "SimplePhysical" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt            As PartFacelets.IJDPart
    Dim parWidth             As Double
    Dim parHeight1           As Double
    Dim parHeight2           As Double
    Dim parCowlDiameter      As Double
    Dim parHeight3           As Double
    Dim iOutput              As Integer
    Dim dDepth               As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)             'a
    parHeight1 = arrayOfInputs(3)           'F
    parHeight2 = arrayOfInputs(4)           'b
    parCowlDiameter = arrayOfInputs(5)      'd1
    parHeight3 = arrayOfInputs(6)           '200mm or 300mm
    
    iOutput = 0
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    'Insert your code for output 1(Cowl Body)
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oCStr           As IngrGeom3D.ComplexString3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oAxis           As New AutoMath.DVector
    Dim oRevolution     As Object
    Dim iCount          As Integer
    Dim oCentPoint      As New AutoMath.DPosition
    
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.1, -0.5 * parWidth, _
        0, 0.1, -0.5 * parCowlDiameter, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.1, _
        -0.5 * parCowlDiameter, 0, 0.1 + parHeight2 + parHeight3, _
        -0.5 * parCowlDiameter, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        0.1 + parHeight2 + parHeight3, -0.5 * parCowlDiameter, _
        0, 0.1 + parHeight1 + parHeight3, -0.0625, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        0.1 + parHeight1 + parHeight3, -0.0625, 0, _
        0.1 + parHeight1 + parHeight3, 0, 0)
    lines.Add oLine

    oStPoint.Set 0.1, -0.5 * parWidth, 0
    Set oCStr = PlaceTrCString(oStPoint, lines)
    
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    
    oAxis.Set 1, 0, 0
    oCentPoint.Set 0.5 * parHeight2, 0, 0
    Set oRevolution = PlaceRevolution(m_OutputColl, oCStr, oAxis, oCentPoint, _
                                            2 * PI, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRevolution
    
    'Insert your code for output 2(Cylinder)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0.1 + parHeight3, 0, 0

    'Set the Output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, _
                                            oEnPoint, parWidth
   
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxis = Nothing
    Set oCentPoint = Nothing
    Set oLine = Nothing
    Set oArc = Nothing
    Set lines = Nothing
    Set oCStr = Nothing
    Set oRevolution = Nothing
    Set oGeomFactory = Nothing
    
    'Place Hvac Nozzle 1(Output 3)
    Dim oHvacNozzle      As GSCADNozzleEntities.IJDNozzle
    Dim oNozzleFactory   As New GSCADNozzleEntities.NozzleFactory
    Dim oNozzle          As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort     As GSCADNozzleEntities.IJDistribPort
    Dim CP               As New AutoMath.DPosition
    Dim PortStatus       As DistribPortStatus
    Dim oPortDirection   As AutoMath.DVector
    Dim oHvacPort        As IJDHvacPort
    Dim oHvacColl        As IJDCollection
    
    Dim bDimBaseOuter    As Boolean
    Dim iPortIndex       As Integer
    Dim lEndPrep         As Long
    Dim dThickness       As Double
    Dim dFlangeWidth     As Double
    Dim lFlowDir         As Long
    Dim dPortDepth       As Double
    Dim dCptOffSet       As Double
    Dim dNozzLength      As Double
    Dim dCornerRadius    As Double
       
    dCornerRadius = 0#
    dNozzLength = 0.2 * parWidth
    iPortIndex = 1
    
    'Set HVAC nozzle parameters
    Set oPortDirection = New AutoMath.DVector
    Set oHvacColl = oPartFclt.GetNozzles()
    Set oHvacPort = oHvacColl.Item(1)
    lEndPrep = oHvacPort.EndPrep
    dThickness = oHvacPort.Thickness
    dFlangeWidth = oHvacPort.FlangeWidth
    lFlowDir = oHvacPort.FlowDirection
    dPortDepth = oHvacPort.PortDepth
    dCptOffSet = oHvacPort.CptOffset
    
    dDepth = parWidth
    bDimBaseOuter = True
    
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    If CmpDblLessThanOrEqualTo(dThickness, LINEAR_TOLERANCE) Then
            dThickness = NEGLIGIBLE_THICKNESS
    End If
    If CmpDblLessThan(dNozzLength, dThickness) Then
            dNozzLength = dThickness + NEGLIGIBLE_THICKNESS
    End If
     
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, lEndPrep, _
                                            dThickness, dFlangeWidth, lFlowDir, parWidth, _
                                            dDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                            "Hvac1", dPortDepth, dCptOffSet, False, m_OutputColl.ResourceManager)
    CP.Set 0, 0, 0
    
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation CP
     
    'Direction specified here of the nozzle should be the direction in which
    'pipe will be routed. Graphics of the nozzle will appear in opposite
    'direction to the direction specified on the nozzle.
    oPortDirection.Set -1, 0, 0
    oDistribPort.SetDirectionVector oPortDirection
    Set oPortDirection = Nothing
    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzLength

    'Set the output for Hvac Nozzzle
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    Set oNozzleFactory = Nothing
    Set oPortDirection = Nothing
    Set CP = Nothing
    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oDistribPort = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub

