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
'   Creation Date   :  Friday,August 10 2007
'   Description     :  HVAC Weather Tight Cover
'   Reference       :  STD-M4.pdf
'   CR number       :  CR-124219
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "SimplePhysical" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
        
        PI = 4 * Atn(1)
        
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), _
    arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt           As PartFacelets.IJDPart
    Dim parWidth            As Double
    Dim parDepth            As Double
    Dim parCoverLength      As Double
    Dim parCoverWidth       As Double
    Dim parCoverThickness   As Double
    Dim parCoverEdgeHeight  As Double
    Dim parHandletoCenter   As Double
    
    Dim iOutput             As Integer
   
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parCoverLength = arrayOfInputs(4)       'L
    parCoverWidth = arrayOfInputs(5)        'W
    parCoverThickness = arrayOfInputs(6)    't
    parCoverEdgeHeight = arrayOfInputs(7)   '30 mm
    parHandletoCenter = arrayOfInputs(8)    'F
   
    iOutput = 0
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    'Insert your code for output 1(CoverPlate)
    
    oStPoint.Set 0, 0.5 * parCoverLength - parCoverThickness, _
                    -0.5 * parCoverWidth + parCoverThickness
    oEnPoint.Set -parCoverThickness, -0.5 * parCoverLength + parCoverThickness, _
                    0.5 * parCoverWidth - parCoverThickness
    
    Dim oCoverPlate As Object
    Set oCoverPlate = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCoverPlate
    Set oCoverPlate = Nothing
    
    'Insert your code for output 2(CoverEdge1)
    Dim oCoverEdge1 As Object
    Dim oCoverEdge2 As Object
    Dim oCoverEdge3 As Object
    Dim oCoverEdge4 As Object
    
    oStPoint.Set -parCoverThickness, 0.5 * parCoverLength _
        - parCoverThickness, -0.5 * parCoverWidth
    oEnPoint.Set parCoverEdgeHeight - parCoverThickness, _
        0.5 * parCoverLength, 0.5 * parCoverWidth
    
    Set oCoverEdge1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCoverEdge1
    Set oCoverEdge1 = Nothing
    
    'Insert your code for output 3(CoverEdge2)
    oStPoint.Set -parCoverThickness, 0.5 * parCoverLength _
        - parCoverThickness, 0.5 * parCoverWidth
    oEnPoint.Set parCoverEdgeHeight - parCoverThickness, _
                -0.5 * parCoverLength + parCoverThickness, _
                0.5 * parCoverWidth - parCoverThickness
    
    Set oCoverEdge2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCoverEdge2
    Set oCoverEdge2 = Nothing
    
    'Insert your code for output 4(CoverEdge3)
    oStPoint.Set parCoverEdgeHeight - parCoverThickness, _
        -0.5 * parCoverLength + parCoverThickness, 0.5 * parCoverWidth
    oEnPoint.Set -parCoverThickness, -0.5 * parCoverLength, -0.5 * parCoverWidth
    
    Set oCoverEdge3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCoverEdge3
    Set oCoverEdge3 = Nothing
    
    'Insert your code for output 5(CoverEdge4)
    oStPoint.Set -parCoverThickness, -0.5 * _
        parCoverLength + parCoverThickness, -0.5 * parCoverWidth
    oEnPoint.Set parCoverEdgeHeight - parCoverThickness, _
                0.5 * parCoverLength - parCoverThickness, _
                -0.5 * parCoverWidth + parCoverThickness

    Set oCoverEdge4 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCoverEdge4
    Set oCoverEdge4 = Nothing
    
    'Insert your code for output 6(Handle)
    Dim oCsCircle       As IngrGeom3D.Circle3d
    Dim oTrHandle       As IngrGeom3D.ComplexString3d
    Dim oHandle         As Object
    Dim oLines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oSurfset        As IngrGeom3D.IJElements
    Dim iCount          As Integer
    Dim stnorm()        As Double
    Dim ednorm()        As Double
     
    Set oLines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oCsCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
        -parCoverThickness, 0.1 * parCoverWidth, 0.35 * parCoverLength, 1, 0, 0, _
        0.015 * parCoverWidth)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
        -parCoverThickness, 0.1 * parCoverWidth, 0.35 * parCoverLength, -0.05 * _
        parCoverWidth - parCoverThickness, 0.1 * parCoverWidth, 0.35 * _
        parCoverLength)
    oLines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -0.05 * _
        parCoverWidth - parCoverThickness, 0.1 * parCoverWidth, 0.35 * _
        parCoverLength, -0.05 * parCoverWidth - parCoverThickness, -0.1 * _
        parCoverWidth, 0.35 * parCoverLength)
    oLines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -0.05 * _
        parCoverWidth - parCoverThickness, -0.1 * parCoverWidth, 0.35 * _
        parCoverLength, -parCoverThickness, -0.1 * parCoverWidth, 0.35 * _
        parCoverLength)
    oLines.Add oLine
    oStPoint.Set -parCoverThickness, 0.1 * parCoverWidth, 0.35 * parCoverLength

    Set oTrHandle = PlaceTrCString(oStPoint, oLines)
    Set oSurfset = _
        oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
        oTrHandle, oCsCircle, CircularCorner, 0, stnorm, ednorm, False)
    
    Dim oRotMat As DT4x4
    Set oRotMat = New DT4x4
    Dim oVec As DVector
    Set oVec = New DVector
    
    If CmpDblGreaterthanOrEqualTo(parCoverWidth, 1.05) Then
        Dim oTransMat As DT4x4
        Set oTransMat = New DT4x4
        oTransMat.LoadIdentity
        oVec.Set 0, parHandletoCenter, 0
        oTransMat.Translate oVec
        
        oRotMat.LoadIdentity
        oVec.Set 1, 0, 0
        oRotMat.Rotate PI / 2, oVec
        
        For Each oHandle In oSurfset
            oHandle.Transform oTransMat
            oHandle.Transform oRotMat
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHandle
        Next oHandle
        Set oSurfset = Nothing
        oRotMat.LoadIdentity
        oVec.Set 1, 0, 0
        oRotMat.Rotate PI / 2, oVec
        Set oSurfset = _
            oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
            oTrHandle, oCsCircle, CircularCorner, 0, stnorm, ednorm, False)
        oTransMat.LoadIdentity
        oVec.Set 0, -parHandletoCenter, 0
        oTransMat.Translate oVec
        For Each oHandle In oSurfset
            oHandle.Transform oTransMat
            oHandle.Transform oRotMat
            m_OutputColl.AddOutput "Handle2_", oHandle
        Next oHandle
        Set oTransMat = Nothing
    Else
        oRotMat.LoadIdentity
        oVec.Set 1, 0, 0
        oRotMat.Rotate PI / 2, oVec
        For Each oHandle In oSurfset
            oHandle.Transform oRotMat
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHandle
        Next oHandle
    End If
    Set oRotMat = Nothing
    Set oVec = Nothing
    Set oSurfset = Nothing
    For iCount = 1 To oLines.Count
        oLines.Remove 1
    Next iCount
    Set oCsCircle = Nothing
    Set oTrHandle = Nothing
    Set oHandle = Nothing
    Set oLine = Nothing
    Set oLines = Nothing
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Place Hvac Nozzle 1
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
    dNozzLength = 0.001
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
    
    bDimBaseOuter = True
    
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
    If CmpDblLessThanOrEqualTo(dThickness, LINEAR_TOLERANCE) Then
            dThickness = 0.0001
    End If
    If CmpDblLessThan(dNozzLength, dThickness) Then
            dNozzLength = dThickness + 0.001
    End If
     
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
        GSCADNozzleEntities.Rectangular, lEndPrep, dThickness, dFlangeWidth, _
        lFlowDir, parWidth, parDepth, dCornerRadius, bDimBaseOuter, PortStatus, _
        "Hvac1", dPortDepth, dCptOffSet, False, m_OutputColl.ResourceManager)
    CP.Set 0, 0, 0
    
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation CP
     
    'Direction specified here of the nozzle should be the direction in which
    'pipe will be routed. Graphics of the nozzle will appear in opposite
    'direction to the direction specified on the nozzle.
    oPortDirection.Set 1, 0, 0
    oDistribPort.SetDirectionVector oPortDirection
    
    oPortDirection.Set 0, 1, 0
    oDistribPort.SetRadialOrient oPortDirection
    
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
