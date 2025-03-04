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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   ProgID:         SP3DHEccentricTee.EccentricTee
'   Author:         KKK
'   Creation Date:  Thursday, Jan 19 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect.
'   HVAC Eccentric Tee. TU - (Tangential –with hand made T-piece TSTU)
'   Symbol Taken from the http://www.lindab.com/Sweden/products/pdf/vent_eng/kapitel/eng_safe.pdf - page no.88
'
'   Change History:

'   08.SEP.2006    KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   11.Sep.2006    dkl  TR-105007 In the Else condition for testing if 0.25<parWidth<0.4, replaced & with AND.
'   21.Jan-2008    VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oPos As New AutoMath.DPosition
    Dim NozzleLength As Double
    
    Dim iOutput     As Double
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parLength As Double
    Dim parBLength As Double
    Dim parCornerRadius As Double
    Dim parHVACShape As Integer
    Dim dblOffset As Double
    Dim parInsulationThickness As Double
    Dim PortDirection As AutoMath.DVector
    Set PortDirection = New AutoMath.DVector
    Dim RadialDirection As AutoMath.DVector
    Set RadialDirection = New AutoMath.DVector
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim ObjElbow As Object
    
    Dim parAngle As Double
    parAngle = PI / 2
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parLength = arrayOfInputs(4) ' 'l' - Face to center for Port1 &2
    parBLength = arrayOfInputs(5) ' 'l3' - Face to center for Port3
    
    iOutput = 0
    
    dblOffset = (parWidth - parBWidth) / 2
           
' Place Nozzle 1
    parHVACShape = 4
    parCornerRadius = 0
    PortDirection.Set -1, 0, 0

    oPos.Set -parLength / 2, 0, 0
    RadialDirection.Set 0, 0, 1
    NozzleLength = parLength / 2

    Dim oHvacNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort

' HVAC parameters for Nozzle
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim iPortIndex As Integer
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 3) As Long
    Dim dThickness(1 To 3) As Double
    Dim dFlangeWidth(1 To 3) As Double
    Dim lFlowDir(1 To 3) As Long
    Dim dPortDepth(1 To 3) As Double
    Dim dCptOffSet(1 To 3) As Double
    
    parCornerRadius = 0#
    
    'Set HVAC nozzle parameters from Catalog
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
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
                            parWidth, parCornerRadius, DimBaseOuter, PortStatus, _
                            "HvacNozzle1", dPortDepth(1), dCptOffSet(1), False, _
                            m_OutputColl.ResourceManager)

    Set NozzleFactory = Nothing
    
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation oPos
     
'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    iDistribPort.SetDirectionVector PortDirection

    iDistribPort.SetRadialOrient RadialDirection

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    

' Place Nozzle 2
    PortDirection.Set 1, 0, 0
    NozzleLength = parLength / 2
    oPos.Set parLength / 2, 0, 0
    RadialDirection.Set 0, 0, 1
    
    iPortIndex = 2
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            parHVACShape, lEndPrep(2), _
                            dThickness(2), dFlangeWidth(2), lFlowDir(2), parWidth, _
                            parWidth, parCornerRadius, DimBaseOuter, PortStatus, _
                            "HvacNozzle2", dPortDepth(2), dCptOffSet(2), False, _
                            m_OutputColl.ResourceManager)

     Set NozzleFactory = Nothing
     Set iDistribPort = oHvacNozzle
     iDistribPort.SetPortLocation oPos
     
'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    iDistribPort.SetDirectionVector PortDirection
    iDistribPort.SetRadialOrient RadialDirection
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
' Place Nozzle 3
    PortDirection.Set 0, 1, 0
    NozzleLength = parBLength
    oPos.Set 0, parBLength, dblOffset
    RadialDirection.Set 0, 0, 1
    iPortIndex = 3
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                            parHVACShape, lEndPrep(3), _
                            dThickness(3), dFlangeWidth(3), lFlowDir(3), parBWidth, _
                            parBWidth, parCornerRadius, DimBaseOuter, PortStatus, _
                            "HvacNozzle3", dPortDepth(3), dCptOffSet(3), False, _
                            m_OutputColl.ResourceManager)

     Set NozzleFactory = Nothing

     Set iDistribPort = oHvacNozzle
     iDistribPort.SetPortLocation oPos

'     Direction specified here of the nozzle should be the direction in which
'     pipe will be routed. Graphics of the nozzle will appear in opposite
'     direction to the direction specified on the nozzle.
    iDistribPort.SetDirectionVector PortDirection
    iDistribPort.SetRadialOrient RadialDirection
    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set oPos = Nothing
    Set PortDirection = Nothing
    Set RadialDirection = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
