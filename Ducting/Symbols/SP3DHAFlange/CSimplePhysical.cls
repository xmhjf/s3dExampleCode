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
'   CSimplePhysical.cls
'   Author:         VRK
'   Creation Date:  Monday, June 4 2007
'   Description:
'       Tangential Tee
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages

Private m_oGeomHelper As IJSymbolGeometryHelper
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
        
    '{<(InputParamDec)>}
    Dim parHVACShape As Integer
    Dim parWidth As Double
    Dim parNominalFlatLength As Double
    Dim parDepth As Double
    Dim dHalfWidth As Double
    Dim dHalfDepth As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)                     'D
    parNominalFlatLength = arrayOfInputs(4)         'L
    parDepth = arrayOfInputs(5)                     't

    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    dHalfWidth = parWidth / 2
    dHalfDepth = parDepth / 2
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim objBody As Object
    Dim oAxisVec As New AutoMath.DVector
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oCenter As New AutoMath.DPosition
    
    Dim iPortIndex As Integer
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    Dim dNozzLength(1 To 2) As Double
    
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
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
    
If CmpDblGreaterthan(parWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(dFlangeWidth(2), LINEAR_TOLERANCE) And _
         CmpDblGreaterthan(parNominalFlatLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(dThickness(2), LINEAR_TOLERANCE) Then
    'Insert your code for Cylinder(Output 1)
    If parHVACShape = 4 Then 'Round
    
        dLineStrPoints(0) = -parNominalFlatLength
        dLineStrPoints(1) = 0
        dLineStrPoints(2) = dHalfWidth
    
        dLineStrPoints(3) = -parNominalFlatLength
        dLineStrPoints(4) = 0
        dLineStrPoints(5) = dHalfWidth + dThickness(2)
    
        dLineStrPoints(6) = 0
        dLineStrPoints(7) = 0
        dLineStrPoints(8) = dHalfWidth + dThickness(2)
    
        dLineStrPoints(9) = 0
        dLineStrPoints(10) = 0
        dLineStrPoints(11) = dHalfWidth
        
        dLineStrPoints(12) = -parNominalFlatLength
        dLineStrPoints(13) = 0
        dLineStrPoints(14) = dHalfWidth
    
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 1, 0, 0
        oCenter.Set 0, 0, 0
        Set objBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVec, oCenter, 2 * PI, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjBody_", objBody
        Set objBody = Nothing
        Set oLineString = Nothing
        Set oAxisVec = Nothing
        Set oCenter = Nothing
  End If
End If

If CmpDblGreaterthan(parWidth, LINEAR_TOLERANCE) And CmpDblGreaterthan(dFlangeWidth(2), LINEAR_TOLERANCE) And _
         CmpDblGreaterthan(parNominalFlatLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(dThickness(2), LINEAR_TOLERANCE) And _
                CmpDblGreaterthan(parDepth, LINEAR_TOLERANCE) Then

  If parHVACShape = Rectangular Then
  
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = dHalfDepth
        dLineStrPoints(2) = dHalfWidth + dThickness(2)
    
        dLineStrPoints(3) = 0
        dLineStrPoints(4) = dHalfDepth + dThickness(2)
        dLineStrPoints(5) = dHalfWidth + dThickness(2)
    
        dLineStrPoints(6) = 0
        dLineStrPoints(7) = dHalfDepth + dThickness(2)
        dLineStrPoints(8) = -(dHalfWidth + dThickness(2))
    
        dLineStrPoints(9) = 0
        dLineStrPoints(10) = dHalfDepth
        dLineStrPoints(11) = -(dHalfWidth + dThickness(2))
    
        dLineStrPoints(12) = 0
        dLineStrPoints(13) = dHalfDepth
        dLineStrPoints(14) = dHalfWidth + dThickness(2)
    
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set -1, 0, 0
        Set objBody = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parNominalFlatLength, True)
        'Set the outputs
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjBody_", objBody
        Set oLineString = Nothing
        
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = -dHalfDepth
        dLineStrPoints(2) = dHalfWidth + dThickness(2)

        dLineStrPoints(3) = 0
        dLineStrPoints(4) = -(dHalfDepth + dThickness(2))
        dLineStrPoints(5) = dHalfWidth + dThickness(2)

        dLineStrPoints(6) = 0
        dLineStrPoints(7) = -(dHalfDepth + dThickness(2))
        dLineStrPoints(8) = -(dHalfWidth + dThickness(2))

        dLineStrPoints(9) = 0
        dLineStrPoints(10) = -dHalfDepth
        dLineStrPoints(11) = -(dHalfWidth + dThickness(2))

        dLineStrPoints(12) = 0
        dLineStrPoints(13) = -dHalfDepth
        dLineStrPoints(14) = dHalfWidth + dThickness(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set -1, 0, 0
        Set objBody = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parNominalFlatLength, True)
        'Set the outputs
        m_OutputColl.AddOutput "ObjBody_", objBody
        Set oLineString = Nothing

        dLineStrPoints(0) = 0
        dLineStrPoints(1) = dHalfWidth + dThickness(2)
        dLineStrPoints(2) = dHalfDepth

        dLineStrPoints(3) = 0
        dLineStrPoints(4) = dHalfWidth + dThickness(2)
        dLineStrPoints(5) = dHalfDepth + dThickness(2)

        dLineStrPoints(6) = 0
        dLineStrPoints(7) = -(dHalfWidth + dThickness(2))
        dLineStrPoints(8) = dHalfDepth + dThickness(2)

        dLineStrPoints(9) = 0
        dLineStrPoints(10) = -(dHalfWidth + dThickness(2))
        dLineStrPoints(11) = dHalfDepth

        dLineStrPoints(12) = 0
        dLineStrPoints(13) = dHalfWidth + dThickness(2)
        dLineStrPoints(14) = dHalfDepth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set -1, 0, 0
        Set objBody = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parNominalFlatLength, True)
        'Set the outputs
        m_OutputColl.AddOutput "ObjBody_", objBody
        Set oLineString = Nothing

        dLineStrPoints(0) = 0
        dLineStrPoints(1) = dHalfWidth + dThickness(2)
        dLineStrPoints(2) = -dHalfDepth

        dLineStrPoints(3) = 0
        dLineStrPoints(4) = dHalfWidth + dThickness(2)
        dLineStrPoints(5) = -(dHalfDepth + dThickness(2))

        dLineStrPoints(6) = 0
        dLineStrPoints(7) = -(dHalfWidth + dThickness(2))
        dLineStrPoints(8) = -(dHalfDepth + dThickness(2))

        dLineStrPoints(9) = 0
        dLineStrPoints(10) = -(dHalfWidth + dThickness(2))
        dLineStrPoints(11) = -dHalfDepth

        dLineStrPoints(12) = 0
        dLineStrPoints(13) = dHalfWidth + dThickness(2)
        dLineStrPoints(14) = -dHalfDepth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set -1, 0, 0
        Set objBody = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parNominalFlatLength, True)
        'Set the outputs
        m_OutputColl.AddOutput "ObjBody_", objBody
        Set oLineString = Nothing
        Set oLineString = Nothing
        Set oAxisVec = Nothing
  End If
End If
'-------------------------------------------------------------------------------------------------
'==================
'BUILD HVACNOZZLE1
'==================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim oDir As New AutoMath.DVector
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    
    Dim dWidth As Double
    Dim dDepth As Double
    Dim CornerRadius As Double
    Dim lCSType As Long
            
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    iPortIndex = 1
    If parHVACShape = Rectangular Then
         dWidth = parWidth
         dDepth = parDepth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width
    If parHVACShape = 4 Then 'Round
         dWidth = parWidth
         dDepth = dWidth
        lCSType = GSCADNozzleEntities.Round
    End If
 
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
'    dCptOffSet(1) = parNominalFlatLength
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0
    Y = 0#
    Z = 0#
    
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    oDistribPort.SetDirectionVector dir
    If parHVACShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

'==================
'BUILD HVACNOZZLE2
'==================
    iPortIndex = 2
    
    If parHVACShape = Rectangular Then
        dWidth = parWidth
        dDepth = parDepth
        lCSType = GSCADNozzleEntities.Rectangular
    End If
    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the Width2
    If parHVACShape = 4 Then 'Round
        dWidth = parWidth
        dDepth = parWidth
        lCSType = GSCADNozzleEntities.Round
    End If

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), dWidth, _
                                            dDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
    
    'Position of the nozzle should be the conenct point of the nozzle
    X = dThickness(2)
    Y = 0
    Z = 0#
    
    pos.Set X, Y, Z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
    dir.Set 1, 0, 0
    oDistribPort.SetDirectionVector dir
    
     If parHVACShape = Rectangular Then
        oDir.Set 0, 1, 0
        oDistribPort.SetRadialOrient oDir
    End If
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oDir = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
    
    
