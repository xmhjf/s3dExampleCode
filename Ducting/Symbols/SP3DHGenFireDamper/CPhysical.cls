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
'   Author:          RUK
'   Creation Date:  Monday, Aug 27 2007
'   Description:
'    CR-124211 Generic Fire Damper.
'       Source STD-M13.pdf.
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   27-Aug-2007     RUK     Created
'   07.Nov.2007     RUK     DI-CP·56531 Ability for custom weight and cg calculation for parametric standard components.
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
    Dim iOutput     As Double
    
    Dim parHVACShape As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBladetoPort1 As Double
    Dim parPPtoBlade As Double
    Dim parPPtoPort2 As Double
    Dim parDamperThick As Double
    Dim parBladeAngle As Double
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parBladetoPort1 = arrayOfInputs(5)
    parPPtoBlade = arrayOfInputs(6)
    parPPtoPort2 = arrayOfInputs(7)
    parDamperThick = arrayOfInputs(8)   'Not used in the code
    parBladeAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)

    'Origin is taken at the Center of the Damper
    'For Horizontally routed Duct X - Along the flow and Y - towards UP
    iOutput = 0
    
    If parHVACShape = 4 Then 'Round
        parDepth = parWidth
    End If
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oGeomFactory = New GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    'Create the Damper (Output 1)
    Dim objDamper As Object
    If parHVACShape = Rectangular Then
        oStPoint.Set -(parPPtoBlade + parBladetoPort1), -parDepth / 2, -parWidth / 2
        oEnPoint.Set parPPtoPort2, parDepth / 2, parWidth / 2
        Set objDamper = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    ElseIf parHVACShape = 4 Then 'Round
        oStPoint.Set -(parPPtoBlade + parBladetoPort1), 0, 0
        oEnPoint.Set parPPtoPort2, 0, 0
        Set objDamper = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
    End If
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamper
    Set objDamper = Nothing
    
    'Create Blade (Output 2)
    Dim objBlade As Object
    If parHVACShape = Rectangular Then
        Dim dPoints(0 To 14) As Double
        
        'Point 1
        dPoints(0) = 0
        dPoints(1) = -parDepth / 2
        dPoints(2) = parWidth / 2
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = dPoints(1)
        dPoints(5) = -dPoints(2)
        'Point 3
        dPoints(6) = dPoints(0)
        dPoints(7) = -dPoints(4)
        dPoints(8) = dPoints(5)
        'Point 4
        dPoints(9) = dPoints(0)
        dPoints(10) = dPoints(7)
        dPoints(11) = -dPoints(8)
        'Point 5
        dPoints(12) = dPoints(0)
        dPoints(13) = dPoints(1)
        dPoints(14) = dPoints(2)
        
        Set objBlade = oGeomFactory.LineStrings3d.CreateByPoints(m_OutputColl.ResourceManager, _
                                    5, dPoints)
        Dim oTransMat As AutoMath.DT4x4
        Dim oVector As AutoMath.DVector
        Set oTransMat = New DT4x4
        Set oVector = New DVector
        
        'Rotate the Blade
        oTransMat.LoadIdentity
        oVector.Set 0, 0, 1
        oTransMat.Rotate PI / 2 - parBladeAngle, oVector
        objBlade.Transform oTransMat
        'Transform the Blade
        oTransMat.LoadIdentity
        oVector.Set -parPPtoBlade, 0, 0
        oTransMat.Translate oVector
        objBlade.Transform oTransMat
        
        Set oTransMat = Nothing
        Set oVector = Nothing
    
    ElseIf parHVACShape = 4 Then 'Round
        Dim oCenter As AutoMath.DPosition
        Dim oNormal As AutoMath.DVector
        
        Set oCenter = New DPosition
        Set oNormal = New DVector
        
        oCenter.Set -parPPtoBlade, 0, 0
        oNormal.Set -Sin(parBladeAngle), -Cos(parBladeAngle), 0
        Set objBlade = oGeomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                    oCenter.x, oCenter.y, oCenter.z, _
                                    oNormal.x, oNormal.y, oNormal.z, parWidth / 2)
        Set oCenter = Nothing
        Set oNormal = Nothing
    End If
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBlade
    Set objBlade = Nothing
    
    'Remove references
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
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
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                parHVACShape, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)

    'Position of the Nozzle
    Dim x As Double, y As Double, z As Double
    x = -(parPPtoBlade + parBladetoPort1)
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
                    parHVACShape, lEndPrep(2), _
                    dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                    parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                    "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                    m_OutputColl.ResourceManager)

    
    'Position of the Nozzle
    x = parPPtoPort2
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
    
    'New code for WCGOutputs DI-CP·56531
    Dim oGenBOFactory As IJGeneralBusinessObjectsFactory
    Dim oAttributes As IJDAttributes
    Dim oAttribCol As IJDAttributesCol
    Set oGenBOFactory = New GeneralBusinessObjectsFactory
    
    Set oAttributes = oGenBOFactory.CreateWCGPropHolder(m_OutputColl.ResourceManager)

    'In V8 we should use the IJWCGPropHolder interface to set these values
    Set oAttribCol = oAttributes.CollectionOfAttributes(CVar("IJWCGPropHolder"))
    
    'Calculate the Volume
        'Assumed that the thickness of duct is negligible
    Dim dTotVolume As Double
    Dim dHeaderVolume As Double
    Dim dBladeVolume As Double
    Dim dCGX, dCGY, dCGZ As Double
    
    If parHVACShape = Rectangular Then
        dHeaderVolume = 2 * (parWidth + parDepth) * (parBladetoPort1 + parPPtoBlade + parPPtoPort2)
        'Assumed that blade has negligible thickness
        dBladeVolume = (parWidth * parDepth)
    ElseIf parHVACShape = 4 Then 'Round
        dHeaderVolume = PI * parWidth * (parBladetoPort1 + parPPtoBlade + parPPtoPort2)
        'Assumed that blade has negligible thickness
        dBladeVolume = PI * ((parWidth / 2) ^ 2)
    End If
    dTotVolume = dHeaderVolume + dBladeVolume
        
    dCGX = -((dBladeVolume * (((parBladetoPort1 + parPPtoBlade + parPPtoPort2) / 2) _
                                            - parBladetoPort1) / dTotVolume) + _
            (((parBladetoPort1 + parPPtoBlade + parPPtoPort2) / 2) - parPPtoPort2))
    dCGY = 0
    dCGZ = 0
    
    oAttribCol.Item("Volume").Value = dTotVolume '(cubic meters)
    oAttribCol.Item("CGX").Value = dCGX
    oAttribCol.Item("CGY").Value = dCGY
    oAttribCol.Item("CGZ").Value = dCGZ
    
    m_OutputColl.AddOutput "StrippedWCG", oAttributes
    Set oAttribCol = Nothing
    Set oAttributes = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
