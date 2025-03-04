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
'   Creation Date:  Monday 6, June 2005
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Fincantieri Hvac Exhaust Grid for Rectangular Duct (Reference: Fincantieri drawing).
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   01.Aug.2006     KKK       TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
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
    Dim parDepth As Double
    Dim parDuctLength As Double
    Dim parPlateCutDimension As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parDuctLength = arrayOfInputs(4)
    parPlateCutDimension = arrayOfInputs(5)
    

    iOutput = 0
    
'   Origin is taken at port center

'   Dimensions are taken from Fincantieri drawing wherever necessary

' Insert your code for output 1(Rectangular Duct Tapered)
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oRuledSurface As IngrGeom3D.RuledSurface3d
    Dim oBaseLinestr As IngrGeom3D.LineString3d
    Dim oTopLinestr As IngrGeom3D.LineString3d
    Dim dCornerpts(0 To 14) As Double
    Dim dNoCorpts As Double
    
    Dim oBaseCenterPt As AutoMath.DPosition
    Dim oTopCenterPt As AutoMath.DPosition

    Set oBaseCenterPt = New AutoMath.DPosition
    Set oTopCenterPt = New AutoMath.DPosition

    oBaseCenterPt.Set parDuctLength - parPlateCutDimension - 0.03, 0, 0
    oTopCenterPt.Set parDuctLength, -parDepth / 2 + 0.025, 0
    
    dCornerpts(0) = oBaseCenterPt.X
    dCornerpts(1) = oBaseCenterPt.Y - parDepth / 2
    dCornerpts(2) = oBaseCenterPt.Z + parWidth / 2
    
    dCornerpts(3) = oBaseCenterPt.X
    dCornerpts(4) = dCornerpts(1) + parDepth
    dCornerpts(5) = dCornerpts(2)
    
    dCornerpts(6) = oBaseCenterPt.X
    dCornerpts(7) = dCornerpts(4)
    dCornerpts(8) = dCornerpts(5) - parWidth
    
    dCornerpts(9) = oBaseCenterPt.X
    dCornerpts(10) = dCornerpts(7) - parDepth
    dCornerpts(11) = dCornerpts(8)
    
    dCornerpts(12) = dCornerpts(0)
    dCornerpts(13) = dCornerpts(1)
    dCornerpts(14) = dCornerpts(2)

    Set oBaseLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
    
    ''Center of the Top Rectangle differs only in Z-coordinate by just adding height to Z coordinate of base center.
    ''Remaining coordinates are same.
    dCornerpts(0) = oTopCenterPt.X
    dCornerpts(1) = oTopCenterPt.Y - 0.05 / 2
    dCornerpts(2) = oTopCenterPt.Z + parWidth / 2
    
    dCornerpts(3) = oTopCenterPt.X
    dCornerpts(4) = dCornerpts(1) + 0.05
    dCornerpts(5) = dCornerpts(2)
    
    dCornerpts(6) = oTopCenterPt.X
    dCornerpts(7) = dCornerpts(4)
    dCornerpts(8) = dCornerpts(5) - parWidth
    
    dCornerpts(9) = oTopCenterPt.X
    dCornerpts(10) = dCornerpts(7) - 0.05
    dCornerpts(11) = dCornerpts(8)
    
    dCornerpts(12) = dCornerpts(0)
    dCornerpts(13) = dCornerpts(1)
    dCornerpts(14) = dCornerpts(2)
    
    Set oTopLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
    Dim objRectangularDuctTapered As Object
    Set objRectangularDuctTapered = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                    oTopLinestr, oBaseLinestr, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRectangularDuctTapered
    Set objRectangularDuctTapered = Nothing

' Insert your code for output 2(Bottom Plate)
    Dim oStPoint As IJDPosition
    Dim oEnPoint As IJDPosition

    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    oStPoint.Set oBaseCenterPt.X - 0.025, oBaseCenterPt.Y - parDepth / 2 - 0.002, oBaseCenterPt.Z - parWidth / 2
    oEnPoint.Set oStPoint.X + parPlateCutDimension + 0.05, oStPoint.Y + 0.002, oStPoint.Z + parWidth
    Dim objBottomPlate As Object
    Set objBottomPlate = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomPlate
    Set objBottomPlate = Nothing

'=====================================
'BUILD HVAC NOZZLE: Rectangular Duct Port
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
    NozzleLength = parDuctLength - parPlateCutDimension - 0.03
    PortDepth = 0#
    cptOffset = 0#

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness


    FlowDir = DistribFlow_IN

    ' Depth and parWidth of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, parWidth, _
                                            parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "RectangularDuctPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
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
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

