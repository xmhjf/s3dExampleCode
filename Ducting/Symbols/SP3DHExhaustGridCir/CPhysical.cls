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
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls

'   Author:         kkk & svsmylav
'   Creation Date:  Saturday 4, June 2005
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Hvac Fincantieri Circular Duct (Reference: Fincantieri drawing).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDuctLength As Double
    Dim parPlateCutDimension As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDuctLength = arrayOfInputs(3)
    parPlateCutDimension = arrayOfInputs(4)

    iOutput = 0


'   Origin is taken at port center

 ' Insert your code for output 1(Bottom Plate)
    Dim oStPoint As IJDPosition
    Dim oEnPoint As IJDPosition
    Dim oAxis As IJDVector

    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oAxis = New DVector

    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim objBottomPlate As IJDObject
    Dim oEleCollection           As Collection
    Set oEleCollection = New Collection
    Dim oAxisVect As New AutoMath.DVector
    
    Dim dlinepts() As Double
    ReDim dlinepts(0 To 17) As Double

    dlinepts(0) = parDuctLength - 0.05
    dlinepts(1) = (parWidth / 2) * Cos(3 * PI / 4)
    dlinepts(2) = (parWidth / 2) * Sin(3 * PI / 4)

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = (parWidth / 2) * Cos(PI)
    dlinepts(5) = (parWidth / 2) * Sin(PI)

    dlinepts(6) = dlinepts(0)
    dlinepts(7) = (parWidth / 2) * Cos(5 * PI / 4)
    dlinepts(8) = (parWidth / 2) * Sin(5 * PI / 4)

    dlinepts(9) = dlinepts(0)
    dlinepts(10) = (parWidth / 2 + 0.01) * Cos(3 * PI / 4)
    dlinepts(11) = (parWidth / 2 + 0.01) * Sin(3 * PI / 4)

    dlinepts(12) = dlinepts(0)
    dlinepts(13) = (parWidth / 2 + 0.01) * Cos(PI)
    dlinepts(14) = (parWidth / 2 + 0.01) * Sin(PI)

    dlinepts(15) = dlinepts(0)
    dlinepts(16) = (parWidth / 2 + 0.01) * Cos(5 * PI / 4)
    dlinepts(17) = (parWidth / 2 + 0.01) * Sin(5 * PI / 4)


    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, dlinepts(0), 0, 0, _
                                                dlinepts(0), dlinepts(1), dlinepts(2), _
                                                dlinepts(6), dlinepts(7), dlinepts(8))
    oEleCollection.Add oArc
    Set oArc = Nothing
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                              dlinepts(15), dlinepts(16), dlinepts(17))
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(15), dlinepts(16), dlinepts(17), _
                                                   dlinepts(12), dlinepts(13), dlinepts(14), _
                                                   dlinepts(9), dlinepts(10), dlinepts(11))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                              dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oLine
    Set oLine = Nothing

    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    oAxisVect.Set -1, 0, 0
    Set objBottomPlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVect, parPlateCutDimension, True)
                                           
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomPlate
    Set objBottomPlate = Nothing
    
    Dim Count As Integer
    For Count = 1 To oEleCollection.Count
    oEleCollection.Remove 1
    Next Count

' Place Circular Duct Port
'=====================================
'BUILD HVAC NOZZLE: Circular Duct Port
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim Depth As Double
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim cptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double

    'Set HVAC nozzle parameters
    iPortIndex = 1
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = parDuctLength
    PortDepth = 0#
    cptOffset = 0#
    Depth = parWidth
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness


    FlowDir = DistribFlow_IN

    ' Depth and parWidth of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, parWidth, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "CircularDuctPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir

    dir.Set 0, -1, 0
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxis = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

