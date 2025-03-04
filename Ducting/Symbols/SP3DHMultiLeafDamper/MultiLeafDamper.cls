VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiLeafDamper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'File
'    MultiLeafDamper.cls
'
'Author
'    NN 10 March 2002
'
'Description
'    Definition of HVAC MultiLeafDamper.
'
'Notes
'
'    <notes>
'
'History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   03.Nov.2005     KKK     CR-87366  Create TROX HVAC parts and symbols
'   06.Oct.2006     dkl     TR-106671 Rectified the direction vectors for placing the Handle.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Private Const MODULE = "MultiLeafDamper:"  'Used for error messages
  Const METHOD = "PlaceTrCString:"

Const NEGLIGIBLE_THICKNESS = 0.0001

Private Const nOutputs = 3

Private m_GeomFactory As IngrGeom3D.GeometryFactory

Implements IJDUserSymbolServices

Private Sub Class_Initialize()

Set m_GeomFactory = New IngrGeom3D.GeometryFactory

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHMultiLeafDamper.MultiLeafDamper"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHMultiLeafDamper.MultiLeafDamper"
  oSymbolDefinition.CodeBase = CodeBase
  
  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
  Exit Function

ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD
End Function
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

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
  
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
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
    Inputs(7).Name = "DamperLength"
    Inputs(7).Description = "DamperLength"
    Inputs(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.158
    Inputs(7).DefaultParameterValue = PC
    
  
    Set Inputs(8) = New IMSSymbolEntities.DInput
    Inputs(8).Name = "HandleOffset"
    Inputs(8).Description = "HandleOffset"
    Inputs(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0.9
    Inputs(8).DefaultParameterValue = PC
  
    Set Inputs(9) = New IMSSymbolEntities.DInput
    Inputs(9).Name = "HandlePosition"
    Inputs(9).Description = "HandlePosition"
    Inputs(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
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
    O(1).Description = "HvacPort1 of MultiLeafDamper"
    
    O(2).Name = "HvacNozzle2"
    O(2).Description = "HvacPort2 of MultiLeafDamper"
    
    O(3).Name = "Hebel"
    O(3).Description = "Hebel"
    
    
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
    PhysicalRepEval.ProgId = "SP3DHMultiLeafDamper.MultiLeafDamper"

    
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
  ReportUnanticipatedError MODULE, METHOD
End Sub
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
        Physical arrayOfInputs
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HFF
'=========================================================================

Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorLabel
    Dim Width As Double, Depth As Double
    Dim FilterLength As Double, Thickness As Double
    Dim FlangeWidth As Double
    Dim DamperLength As Double
    Dim HandleOffset As Double
    Dim HandlePosition As Long
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
    DamperLength = arrayOfInputs(7)
    HandleOffset = arrayOfInputs(8)
    HandlePosition = arrayOfInputs(9)
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
    NozzleLength = FilterLength - Thickness
    FlowDir = DistribFlow_IN
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)
   
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = -DamperLength * 0.5 - CptOffset '  -HFilterLength - CptOffset
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
    iNozzle.Length = DamperLength * 0.5 ' NozzleLength

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
    X = DamperLength * 0.5 + CptOffset ' HFilterLength + CptOffset
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
    Set iNozzle = oHvacNozzle
    iNozzle.Length = DamperLength * 0.5 ' NozzleLength

    m_outputColl.AddOutput "HvacNozzle2", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    
    Dim objStart As New AutoMath.DPosition
    Dim objEnd As New AutoMath.DPosition
    Dim objP1 As New AutoMath.DPosition
    Dim objP2 As New AutoMath.DPosition
    Dim objP3 As New AutoMath.DPosition
    Dim objP4 As New AutoMath.DPosition
    Dim objDir As New AutoMath.DVector
    Dim objUp As New AutoMath.DVector
    Dim objQuer As New AutoMath.DVector
    
   objUp.Set 0, 1, 0
   objQuer.Set 0, 0, 1
   
    If HandlePosition = 1 Then
        objStart.Set DamperLength / 2, Depth * 0.5 + FlangeWidth, Width * 0.5 - HandleOffset
    ElseIf HandlePosition = 2 Then
        objStart.Set 0, Depth * 0.5 - HandleOffset, Width * 0.5 + FlangeWidth
    ElseIf HandlePosition = 3 Then
        objStart.Set 0, -Depth * 0.5 - FlangeWidth, Width * 0.5 - HandleOffset
        objUp.Set 0, -1, 0
    ElseIf HandlePosition = 4 Then
        objStart.Set 0, Depth * 0.5 - HandleOffset, -Width * 0.5 - FlangeWidth
        objQuer.Set 0, 0, -1
    ElseIf HandlePosition = 5 Then
        objStart.Set 0, Depth * 0.5 + FlangeWidth, -Width * 0.5 + HandleOffset
    ElseIf HandlePosition = 6 Then
        objStart.Set 0, -Depth * 0.5 + HandleOffset, Width * 0.5 + FlangeWidth
    ElseIf HandlePosition = 7 Then
        objStart.Set 0, -Depth * 0.5 - FlangeWidth, -Width * 0.5 + HandleOffset
        objUp.Set 0, -1, 0
    Else
        objStart.Set 0, -Depth * 0.5 + HandleOffset, -Width * 0.5 - FlangeWidth
        objQuer.Set 0, 0, -1
    End If
        
    objDir.Set 0, -0.16, 0
    

    Set objP1 = vecDir2(objStart, objDir, 0.001, objUp, 0#, objQuer, 0#)
    Set objP2 = vecDir2(objStart, objDir, 0.001, objUp, 0#, objQuer, 0.02)
    Set objP3 = vecDir2(objStart, objDir, 0.001, objUp, 0.02, objQuer, 0.02)
    Set objP4 = vecDir2(objStart, objDir, 0.001, objUp, 0.02, objQuer, 0#)
    
    objDir.Set -0.16, 0, 0   'Direction of projection is -ve X
    Call createProjectedLines(m_outputColl, "Hebel", -1, _
        objP1, objP2, objP3, objP4, objDir, True)


    Exit Sub
    
ErrorLabel:
  ReportUnanticipatedError MODULE, METHOD
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
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.X, objP3.Y, objP3.Z, objP4.X, objP4.Y, objP4.Z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.X, objP4.Y, objP4.Z, objP1.X, objP1.Y, objP1.Z)
    complex.AddCurve oLine, True
    
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


