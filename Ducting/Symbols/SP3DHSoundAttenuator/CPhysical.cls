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
'   Creation Date:  Monday, Nov 26 2007
'   Description:
'   This symbol is created for CR-129441 (Sound Attenuator)
'   Source: Sound Attenuator.pdf.
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
    Dim parLength As Double
    Dim parThickness As Double
    Dim dInsulationDiameter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parLength = arrayOfInputs(3)
    parThickness = arrayOfInputs(4)

    'Origin is taken at the Center of the Damper
    'For Horizontally routed Duct X - Along the flow and Y - towards UP
    iOutput = 0
    
    'Assumption: If duct size is greater than 1000mm(i.e. 1 meter) then assume length of the
    '            attenuator equals to size of the duct.
    If CmpDblGreaterthan(parWidth, 1) Then parLength = parWidth
    
    'Assumption: Insulation Diameter is calculated for thickness values 50mm(0.05 meter) and 100mm(0.1 meter).
    If CmpDblEqual(parThickness, 0.05) Then
        dInsulationDiameter = parWidth + 2 * parThickness
    ElseIf CmpDblEqual(parThickness, 0.1) Then
        dInsulationDiameter = parWidth + 2 * parThickness
    End If
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Insert your code for Output 1 (Attenuator)
    Dim dLineStrPoints(0 To 26)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim objAttenuator As Object
    
    'Point 1
    dLineStrPoints(0) = -parLength / 2
    dLineStrPoints(1) = 0
    dLineStrPoints(2) = 0
    
    'Point 2
    dLineStrPoints(3) = dLineStrPoints(0)
    dLineStrPoints(4) = parWidth / 2
    dLineStrPoints(5) = dLineStrPoints(2)
    
    'Point 3
    dLineStrPoints(6) = -0.45 * parLength
    dLineStrPoints(7) = dLineStrPoints(4)
    dLineStrPoints(8) = dLineStrPoints(2)
    
    'Point 4
    dLineStrPoints(9) = dLineStrPoints(6)
    dLineStrPoints(10) = dInsulationDiameter / 2
    dLineStrPoints(11) = dLineStrPoints(2)
    
    'Point 5
    dLineStrPoints(12) = 0.45 * parLength
    dLineStrPoints(13) = dLineStrPoints(10)
    dLineStrPoints(14) = dLineStrPoints(2)

    'Point 6
    dLineStrPoints(15) = dLineStrPoints(12)
    dLineStrPoints(16) = parWidth / 2
    dLineStrPoints(17) = dLineStrPoints(2)
    
    'Point 7
    dLineStrPoints(18) = parLength / 2
    dLineStrPoints(19) = dLineStrPoints(16)
    dLineStrPoints(20) = dLineStrPoints(2)
    
    'Point 8
    dLineStrPoints(21) = dLineStrPoints(18)
    dLineStrPoints(22) = 0
    dLineStrPoints(23) = dLineStrPoints(2)
    
    'Point 9
    dLineStrPoints(24) = dLineStrPoints(0)
    dLineStrPoints(25) = dLineStrPoints(1)
    dLineStrPoints(26) = dLineStrPoints(2)
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                            9, dLineStrPoints)
    Set oAxisVect = New AutoMath.DVector
    oAxisVect.Set 1, 0, 0
    Set oCenPoint = New AutoMath.DPosition
    oCenPoint.Set 0, -0.00001, 0
   
    Set objAttenuator = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                        oCenPoint, 2 * PI, True)
   
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objAttenuator
    Set objAttenuator = Nothing
    Set oAxisVect = Nothing
    Set oCenPoint = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing

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
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    Dim dNozzLength(1 To 2) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    dNozzLength(2) = 0.001
    
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
    x = -parLength / 2
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

    '==================================================
    'BUILD HVACNOZZLE 2
    '==================================================
    iPortIndex = 2
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Round, lEndPrep(2), _
                                dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                                dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                                m_OutputColl.ResourceManager)

    'Position of the Nozzle
    x = parLength / 2
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir
    
    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)

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
