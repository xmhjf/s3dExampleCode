VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSplitAttentor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    RSplitAttentor.cls
'
'Author
'    NN 10 Mars 2002
'
'Description
'    Definition of HVAC AirFilterR.
'
'Notes
'
'    <notes>
'
'History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   03.Nov.2005     kkk     CR-87366  Create TROX HVAC parts and symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Private Const nOutputs = 11

Private m_GeomFactory As IngrGeom3D.GeometryFactory

'Private Form As frmEditOccurrence

Implements IJDUserSymbolServices

Private Sub Class_Initialize()

Set m_GeomFactory = New IngrGeom3D.GeometryFactory

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHRSplitAttentor.RSplitAttentor"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHRSplitAttentor.RSplitAttentor"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed SP3DHFlatFlange.HFF Definition
    ' Inputs:
    '          1. "Width"
    '          2. "Depth"
    '          3. "FilterLength"
    '          4. "Thickness"
    '          5. "FlangeThickness"
    
    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)
  
    On Error GoTo ErrorHandler
  
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    ' Create a new input by new operator
    Dim Inputs(1 To 9) As IMSSymbolEntities.DInput
    Dim index As Integer
  
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
    
    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).Name = "Width"
    Inputs(2).Description = "Width of the Rect Duct"
    Inputs(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.4
    Inputs(2).DefaultParameterValue = PC
    
    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).Name = "Depth"
    Inputs(3).Description = "Depth of the Rect Duct"
    Inputs(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.2
    Inputs(3).DefaultParameterValue = PC
    
    
    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).Name = "Length"
    Inputs(4).Description = "FilterLength between ducts"
    Inputs(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.35
    Inputs(4).DefaultParameterValue = PC
    
    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).Name = "Thickness"
    Inputs(5).Description = "Thickness of the Flat Flange of AirFilterR"
    Inputs(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.01
    Inputs(5).DefaultParameterValue = PC
    
    Set Inputs(6) = New IMSSymbolEntities.DInput
    Inputs(6).Name = "FlangeWidth"
    Inputs(6).Description = "Width of the Flat Flange of AirFilterR"
    Inputs(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.02
    Inputs(6).DefaultParameterValue = PC
    
    Set Inputs(7) = New IMSSymbolEntities.DInput
    Inputs(7).Name = "AttentuatorWidth"
    Inputs(7).Description = "AttentuatorWidth"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.05
    Inputs(7).DefaultParameterValue = PC
    
    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "PlateThickness"
    Inputs(8).Description = "PlateThickness"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.005
    Inputs(8).DefaultParameterValue = PC
    
    Set Inputs(9) = New IMSSymbolEntities.DInput
    Inputs(9).Name = "AttentuatorLength"
    Inputs(9).Description = "AttentuatorLength"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.5
    Inputs(9).DefaultParameterValue = PC
    
  
    For index = 2 To 9
        InputsIf.SetInput Inputs(index), index
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To nOutputs) As IMSSymbolEntities.DOutput
    
    For index = 1 To nOutputs
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).Properties = 0
    Next

   
    O(1).Name = "HvacNozzle1"
    O(1).Description = "HvacPort1 of Rectangular Splitter Attentuator"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of Rectangular Splitter Attentuator"
    
    
    O(3).Name = "Hebel"
    O(3).Description = "Hebel"
    
    For index = 4 To nOutputs
      O(index).Name = "Feature" & (index - 3)
      O(index).Description = "Feature" & (index - 3)
    Next index
    
    
    
    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation
  
    rep1.Name = "Physical"
    rep1.Description = "Physical Represntation of the Air Distrib Assembly"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
        
    For index = 1 To nOutputs
        oRepPhysicalOutputs.SetOutput O(index)
        Set O(index) = Nothing
    Next

    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

   
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As DRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "script for the Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DHRSplitAttentor.RSplitAttentor"

    
    'Set the evaluations for the Physical representation on the definition
    Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepEvalsIf = pSymbolDefinition
    RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
    
    Set PhysicalRepEval = Nothing
    
    Set RepEvalsIf = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    Set oRepPhysicalOutputs = Nothing
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False
  Exit Sub
  Resume
End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    'If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
        'End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HFF
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())
    
    Const METHOD = "SP3DHRSplitAttentor-Physical"
    
    On Error GoTo ErrorLabel
    
    
    Dim Width As Double, Depth As Double
    Dim FilterLength As Double, Thickness As Double
    Dim FlangeWidth As Double
    
    Dim AttentuatorWidth As Double
    Dim PlateThickness As Double
    Dim AttentuatorLength As Double

    
    Dim HFilterLength As Double
    Dim NozzleLength As Double
    Dim NozzleFlangeThickness As Double
    Dim CornerRadius As Double
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    'assign to meaningful variables from the input array
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    FilterLength = arrayOfInputs(4)
    Thickness = arrayOfInputs(5)
    FlangeWidth = arrayOfInputs(6)
    
    AttentuatorWidth = arrayOfInputs(7)
    PlateThickness = arrayOfInputs(8)
    AttentuatorLength = arrayOfInputs(9)
    
  
    
'=====================================
'BUILD HVAC NOZZLE ON BASIC ASSEMBLY
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim nozzleIndex As Integer
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iLogicalDistPort As GSCADNozzleEntities.IJLogicalDistPort
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    
    'Set HVAC nozzle parameters
    iPortIndex = 1
    PortDepth = 0#
    CptOffset = 0
    
    EndPrep = 11
    CornerRadius = 0#
    HFilterLength = FilterLength / 2
    'NozzleFlangeThickness = Thickness + HFilterLength
    NozzleLength = FilterLength - Thickness
    'NozzleLength Has to be GREATER than NozzleFlangeThickness
 
    

    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    
    
    FlowDir = DistribFlow_IN

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -AttentuatorLength * 0.5 - CptOffset '  -HFilterLength - CptOffset
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 0.26

    m_outputColl.AddOutput "HvacNozzle1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
'=====================================
'BUILD HVACNOZZLE2 Of AirFilterR
'=====================================
 

    iPortIndex = 2
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    
    
    X = AttentuatorLength * 0.5 + CptOffset ' HFilterLength + CptOffset
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set 1, 0, 0
    iDistribPort.SetDirectionVector dir
    
    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir
'    NozzleLength = Thickness * 1.001
    
 
    Set iNozzle = oHvacNozzle
    iNozzle.Length = AttentuatorLength - 0.26

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    ' GW: Hebel:
    Dim objStart As New AutoMath.DPosition
    Dim objP1 As New AutoMath.DPosition
    Dim objP2 As New AutoMath.DPosition
    Dim objP3 As New AutoMath.DPosition
    Dim objP4 As New AutoMath.DPosition
    Dim objDir As New AutoMath.DVector
    Dim objUp As New AutoMath.DVector
    Dim objQuer As New AutoMath.DVector
''
''    If TroxSide = 1 Then
''        objStart.Set 0, Depth * 0.5 + FlangeWidth, Width * 0.5 - TroxH
''    ElseIf TroxSide = 2 Then
''        objStart.Set 0, Depth * 0.5 - TroxH, Width * 0.5 + FlangeWidth
''    ElseIf TroxSide = 3 Then
''        objStart.Set 0, -Depth * 0.5 - FlangeWidth, Width * 0.5 - TroxH
''    ElseIf TroxSide = 4 Then
''        objStart.Set 0, Depth * 0.5 - TroxH, -Width * 0.5 - FlangeWidth
''
''    ElseIf TroxSide = 5 Then
''        objStart.Set 0, Depth * 0.5 + FlangeWidth, -Width * 0.5 + TroxH
''    ElseIf TroxSide = 6 Then
''        objStart.Set 0, -Depth * 0.5 + TroxH, Width * 0.5 + FlangeWidth
''    ElseIf TroxSide = 7 Then
''        objStart.Set 0, -Depth * 0.5 - FlangeWidth, -Width * 0.5 + TroxH
''    Else
''        objStart.Set 0, -Depth * 0.5 + TroxH, -Width * 0.5 - FlangeWidth
''    End If
''    objDir.Set 0.16, 0, 0
''    objUp.Set 0, 1, 0
''    objQuer.Set 0, 0, 1
''
''
''    Set objP1 = vecDir2(objStart, objDir, -0.01, objUp, 0.01, objQuer, 0#)
''    Set objP2 = vecDir2(objStart, objDir, -0.01, objUp, 0.01, objQuer, 0.01)
''    Set objP3 = vecDir2(objStart, objDir, -0.01, objUp, -0.01, objQuer, 0.01)
''    Set objP4 = vecDir2(objStart, objDir, -0.01, objUp, -0.01, objQuer, 0#)
''
''
''
''    Call createProjectedLines(m_outputColl, "Hebel", -1, _
''        objP1, objP2, objP3, objP4, objDir, True)
    
    'GW: create Kulissen
    Dim lngIndex As Long
    Dim dirDown As New AutoMath.DVector
    Dim dblX As Double
    Dim i As Long
    
    For i = 1 To 2
        If i = 1 Then
            dblX = -AttentuatorLength * 0.5 + 0.15
        Else
            dblX = AttentuatorLength * 0.5 - 0.15
        End If
        objP1.Set dblX, Depth * 0.5 + AttentuatorWidth, Width * 0.5 + AttentuatorWidth
        objP2.Set dblX, -Depth * 0.5 - AttentuatorWidth, Width * 0.5 + AttentuatorWidth
        objP3.Set dblX, -Depth * 0.5 - AttentuatorWidth, -Width * 0.5 - AttentuatorWidth
        objP4.Set dblX, Depth * 0.5 + AttentuatorWidth, -Width * 0.5 - AttentuatorWidth
        
        dirDown.Set -PlateThickness, 0, 0
        
        Call createRect(m_outputColl, "Feature", lngIndex, objP1, objP2, objP3, objP4, AttentuatorWidth, dirDown)
    
    Next i
    
'''    ' GW: add Inspectionsöffnungen.
'''    Dim ObjBodyA As Object
'''
'''    objP1.Set -AttentuatorLength * 0.5 + 0.25 - 0.01, 0.1, Width * 0.5
'''    objP3.Set -AttentuatorLength * 0.5 + FlangeWidth + 0.01, -0.06, Width * 0.5 + 0.03
'''
'''    Set ObjBodyA = PlaceBox(m_outputColl, objP1, objP3)
'''
'''    lngIndex = lngIndex + 1
'''    m_outputColl.AddOutput "Feature" & lngIndex, ObjBodyA
'''    Set ObjBodyA = Nothing
'''
'''    ' GW: add Hebel für Inspectionsöffnungen.
'''
'''
'''    objP1.Set -AttentuatorLength * 0.5 + 0.25 - 0.01, 0.08, Width * 0.5 + 0.03
'''    objP3.Set -AttentuatorLength * 0.5 + 0.25 - 0.01 - 0.16, 0.06, Width * 0.5 + 0.05
'''
'''    Set ObjBodyA = PlaceBox(m_outputColl, objP1, objP3)
'''
'''    lngIndex = lngIndex + 1
'''    m_outputColl.AddOutput "Feature" & lngIndex, ObjBodyA
'''    Set ObjBodyA = Nothing
    
    
    Exit Sub
    
ErrorLabel:
    LogErrorGer METHOD & ": " & Err.Description
    Debug.Print "Error encountered while computing Physical representation", vbExclamation, "RSplitAttentor definition"
        Debug.Assert False
    Exit Sub
    Resume
    
End Sub



Private Function createProjectedLines(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    objVec As AutoMath.DVector, _
                    Optional dblClosed As Boolean = True) As Long
                    
    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP1.X, objP1.Y, objP1.Z, objP2.X, objP2.Y, objP2.Z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP2.X, objP2.Y, objP2.Z, objP3.X, objP3.Y, objP3.Z)
    complex.AddCurve oLine, True
    
    If objP4 Is Nothing Then
        
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.X, objP3.Y, objP3.Z, objP1.X, objP1.Y, objP1.Z)
        complex.AddCurve oLine, True
        
    Else
    
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.X, objP3.Y, objP3.Z, objP4.X, objP4.Y, objP4.Z)
        complex.AddCurve oLine, True
        
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.X, objP4.Y, objP4.Z, objP1.X, objP1.Y, objP1.Z)
        complex.AddCurve oLine, True
    
    End If
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, objVec.X, objVec.Y, objVec.Z, objVec.Length, dblClosed)
                                                    
    If lngIndex >= 0 Then
        lngIndex = lngIndex + 1
        objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
    Else
        objOutputColl.AddOutput strName, Projection
    End If
    
     
     
End Function
        


     
Private Function vecDir2(Pin As AutoMath.DPosition, _
                         Dir1 As AutoMath.DVector, dblSize1 As Double, _
                         Optional Dir2 As AutoMath.DVector = Nothing, _
                            Optional dblSize2 As Double = 0#, _
                         Optional Dir3 As AutoMath.DVector = Nothing, _
                            Optional dblSize3 As Double = 0#) As AutoMath.DPosition

Dim Dir1X As AutoMath.DVector
Dim Dir2X As AutoMath.DVector

Set Dir1X = Dir1.Clone
Dir1X.Length = dblSize1
Set vecDir2 = Pin.Offset(Dir1X)

If Not Dir2 Is Nothing Then
    Set Dir2X = Dir2.Clone
    Dir2X.Length = dblSize2
    Set vecDir2 = vecDir2.Offset(Dir2X)
End If
If Not Dir3 Is Nothing Then
    Set Dir2X = Dir3.Clone
    Dir2X.Length = dblSize3
    Set vecDir2 = vecDir2.Offset(Dir2X)
End If


End Function



Private Function createRect(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    dblToInner As Double, _
                    dirDown As AutoMath.DVector) As Long
                    
' create 4 quader resulting in a rectangle:
'               1             2
'               +-------------+
'               |             |
'               |  +-------+  |
'               |  |       |  |
'               |  |       |  |
'               |  |       |  |
'               |  +-------+  |
'               |             |
'               +-------------+
'               4             3
'
'  P1 - P4 are the 4 outer corner points
'  dblToInner is the size from the outer points to the inner points
'  The size is measured for each side.
'  DirDown is the size and direction in down position
'


' Compute the 4 inner Points
Dim objVec1 As AutoMath.DVector
Dim objVec2 As AutoMath.DVector
 

Dim objP1I As AutoMath.DPosition
Dim objP2I As AutoMath.DPosition
Dim objP3I As AutoMath.DPosition
Dim objP4I As AutoMath.DPosition

 

Set objVec1 = objP2.Subtract(objP1)
Set objVec2 = objP4.Subtract(objP1)
Set objP1I = vecDir2(objP1, objVec1, dblToInner, objVec2, dblToInner)
 
Set objVec1 = objP1.Subtract(objP2)
Set objVec2 = objP3.Subtract(objP2)
Set objP2I = vecDir2(objP2, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP2.Subtract(objP3)
Set objVec2 = objP4.Subtract(objP3)
Set objP3I = vecDir2(objP3, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP3.Subtract(objP4)
Set objVec2 = objP1.Subtract(objP4)
Set objP4I = vecDir2(objP4, objVec1, dblToInner, objVec2, dblToInner)

' Compute the 4 curves (which will later be projected)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP1, objP2, objP2I, objP1I, dirDown)
        
         
        
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP2, objP3, objP3I, objP2I, dirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP3, objP4, objP4I, objP3I, dirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP4, objP1, objP1I, objP4I, dirDown)

End Function



