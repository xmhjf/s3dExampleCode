VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDivisionW3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'File
'    CDivisionW3.cls
'
'Author
'    NN 20020324
'
'Description
'    Definition of HVAC Division Width divided into three cells.
'
'Notes
'   The port 2 is closest from reference (cylinder) and designed the cell 2,
'   port 3 and cell 3 are in the middle, port4 and cell 4 are far from the reference.
'   The Cell dimensions define the dimension along the original Width and dimension one cell
'   along the Width measured from Reference 2 0r 4 (first digit)and the middle cell 3, or cell 4 (second digit)
'   Cell values are: 23, 24, 43
'   CellDimension defines that of cell 2, or cell 4 (first digit of cell value)
'   CellDimension3or4 defines that of cell 3 or cell 4 (second digit of cell value)
'History:
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   05.April.2004   NN           Validate Inputs and report error to middle tier error log.
'   15.DEC.2004      MS  CR 67971 Corrected port indices and flow-directions in HVAC Symbols
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Const NEGLIGIBLE_THICKNESS = 0.0001

Private m_oErrors As IJEditErrors

Private Const MODULE = "CDivisionW3:"
Private Const E_FAIL = &H80004005

'Private Form As frmEditOccurrence

Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DDivisionW3.CDivisionW3"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"

  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition

  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DDivisionW3.CDivisionW3"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.name = oSymbolDefinition.ProgId

  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition

  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing
  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.description
    Debug.Assert False
    ReportUnanticipatedError MODULE, METHOD
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    ' Feed GSCADHRoundFlatFlange.HROUNDFF Definition
    ' Inputs:
    '          1. "Width"
    '          2. "Depth"
    '          3. "Cell"
    '          4. "CellDimension"
    '          5. "CellDimension3or4"
    '
    ' The representation "Physical" is set to the definition
    ' Physical :    (RepresentationId = 1)

    On Error GoTo ErrorHandler
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"

    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 5) As IMSSymbolEntities.DInput
    Dim index As Integer

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC

    PC.Type = igValue

    Set Inputs(1) = New IMSSymbolEntities.DInput
    Inputs(1).name = "Width"
    Inputs(1).description = "Width of the Rect Duct"
    Inputs(1).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.4
    Inputs(1).DefaultParameterValue = PC


    Set Inputs(2) = New IMSSymbolEntities.DInput
    Inputs(2).name = "Depth"
    Inputs(2).description = "Depth of the Rect Duct"
    Inputs(2).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.2
    Inputs(2).DefaultParameterValue = PC

    Set Inputs(3) = New IMSSymbolEntities.DInput
    Inputs(3).name = "Cell"
    Inputs(3).description = "Cells value"
    Inputs(3).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 23
    Inputs(3).DefaultParameterValue = PC

    Set Inputs(4) = New IMSSymbolEntities.DInput
    Inputs(4).name = "CellDimension"
    Inputs(4).description = "Width Cell 2 or 4"
    Inputs(4).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.1
    Inputs(4).DefaultParameterValue = PC

    Set Inputs(5) = New IMSSymbolEntities.DInput
    Inputs(5).name = "CellDimension3or4"
    Inputs(5).description = "Dimension of Cell 3 or 4"
    Inputs(5).properties = igINPUT_IS_A_PARAMETER
    PC.uomValue = 0.15
    Inputs(5).DefaultParameterValue = PC

    For index = 1 To 5
        InputsIf.SetInput Inputs(index), index + 1
        Set Inputs(index) = Nothing
    Next

    'Define the outputs
    Dim O(1 To 5) As IMSSymbolEntities.DOutput

    For index = 1 To 5
        Set O(index) = New IMSSymbolEntities.DOutput
        O(index).properties = 0
    Next


    O(1).name = "GrandPort1"
    O(1).description = "Grand Port1 HvacPort1 of Division"

    O(2).name = "ClosetoRefCell2"
    O(2).description = "Close to Reference, at Cell2 HvacPort2 of Division"

    O(3).name = "MiddleCell3"
    O(3).description = "Middle Cell3 HvacPort3 of Division"

    O(4).name = "FarfromRefCell4"
    O(4).description = "Far from Reference, at Cell4 HvacPort4 of Division"

    O(5).name = "DivisionReference"
    O(5).description = "Reference of Division"


    'Define the representation "Physical"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.name = "Physical"
    rep1.description = "Physical Represntation of the Air Distrib Assembly"
    rep1.properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1

    For index = 1 To 5
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
    PhysicalRepEval.name = "Physical"
    PhysicalRepEval.description = "script for the Physical representation"
    PhysicalRepEval.properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "SP3DDivisionW3.CDivisionW3"


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
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.description
  Debug.Assert False
  ReportUnanticipatedError MODULE, METHOD

End Sub


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    'If StrComp(repName, "Physical") = 0 Then
        Physical arrayOfInputs
        'End If

End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF HRECTFF
'=========================================================================


Private Sub Physical(ByRef arrayOfInputs())

    On Error GoTo ErrorLabel
    Const METHOD = "Physical"

    Dim oPart As PartFacelets.IJDPart
    Dim Width As Double
    Dim Depth As Double, Cell As Long
    Dim CellDimension As Double
    Dim CellDimension3or4 As Double
    Dim NozzleLength As Double
    Dim CornerRadius As Double
    Dim CellWidth(2 To 4) As Double
    Dim NWidth(2 To 4) As Double
    Dim NDepth(2 To 4) As Double
    Dim CellPortnumber As Long

    Dim strDescritption As String

    'assign to meaningful variables from the input array
    Set oPart = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    Cell = arrayOfInputs(4)
    CellDimension = arrayOfInputs(5)
    CellDimension3or4 = arrayOfInputs(6)

'   Input validation
    Dim blnInputErrorFlag As Boolean

'   Collect errors in input values and report them
    If Width <= 0 Then
        strDescritption = vbCrLf & " - Width value is less than or equal to Zero."
        blnInputErrorFlag = True
    End If

    If Depth <= 0 Then
        strDescritption = strDescritption & vbCrLf & " - Depth value is less than or equal to Zero."
        blnInputErrorFlag = True
    End If

    If CellDimension <= 0 Then
        strDescritption = strDescritption & vbCrLf & " - CellDimension value is less than or equal to Zero."
        blnInputErrorFlag = True
    End If

    If CellDimension3or4 <= 0 Then
        strDescritption = strDescritption & vbCrLf & " - CellDimension3or4 value is less than or equal to Zero."
        blnInputErrorFlag = True
    End If

    If CellDimension > Width Then
        strDescritption = strDescritption & vbCrLf & " - CellDimension is greater than the width."
        blnInputErrorFlag = True
    End If

    If CellDimension3or4 > Width Then
        strDescritption = strDescritption & vbCrLf & " - CellDimension3or4 is greater than the width."
        blnInputErrorFlag = True
    End If

    'Check if the computed nozzle dimension is less than the NEGLIGIBLE_THICKNESS
    Dim dNozzleDimension As Double
    dNozzleDimension = Width - CellDimension - CellDimension3or4
    If dNozzleDimension < NEGLIGIBLE_THICKNESS Then
        strDescritption = strDescritption & _
            vbCrLf & " - The computed nozzle dimension (Width - CellDimension - CellDimension3or4) is less than the minimum: " & _
            CStr(dNozzleDimension)
    End If

    If blnInputErrorFlag = True Then GoTo ErrorLabel

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
    Dim pos As IJDPosition
    Dim dir As IJDVector
    Dim CelldirOrientY(2 To 4) As Double
    Dim CelldirOrientZ(2 To 4) As Double
    Dim CelldirOrient As IJDVector
    Set pos = New DPosition
    Set dir = New DVector
    Set CelldirOrient = New DVector
    Dim iPortIndex As Integer

    'Set HVAC nozzle parameters
    iPortIndex = 1
    PortDepth = 0#
    CptOffset = 0#
    EndPrep = 301
    CornerRadius = 0#

    NozzleLength = NEGLIGIBLE_THICKNESS * 2
    FlowDir = DistribFlow_IN
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE

    Dim NThickness As Double
    Dim NFlangeWidth As Double
    NThickness = NEGLIGIBLE_THICKNESS
    NFlangeWidth = NEGLIGIBLE_THICKNESS

        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                            NThickness, NFlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

    If oHvacNozzle Is Nothing Then
        strDescritption = "Port Index = 1: Hvac Nozzle object is nothing"
'        GoTo ErrorHandler
        GoTo ErrorLabel
        'Err.Raise E_FAIL, , "Hvac Nozzle object is nothing"
    End If

    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0
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
    iNozzle.Length = NozzleLength

    m_outputColl.AddOutput "GrandPort1", oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Define arrangement of division cells and construct all ports: 2, 3 and 4 '
'   within each arrangement; three arrangements are considered 23, 24 and 43.'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Cell = 23 Then
            CellWidth(2) = CellDimension
            CellWidth(3) = CellDimension3or4
            CellWidth(4) = Width - CellDimension - CellDimension3or4
        Else
            If Cell = 24 Then
                    CellWidth(2) = CellDimension
                    CellWidth(4) = CellDimension3or4
                    CellWidth(3) = Width - CellDimension - CellDimension3or4
                Else
                    CellWidth(4) = CellDimension
                    CellWidth(3) = CellDimension3or4
                    CellWidth(2) = Width - CellDimension - CellDimension3or4
            End If
    End If
    For CellPortnumber = 2 To 4
        If CellWidth(CellPortnumber) >= Depth Then
                NWidth(CellPortnumber) = CellWidth(CellPortnumber)
                NDepth(CellPortnumber) = Depth
                CelldirOrientY(CellPortnumber) = 1
                CelldirOrientZ(CellPortnumber) = 0
            Else
                NWidth(CellPortnumber) = Depth
                NDepth(CellPortnumber) = CellWidth(CellPortnumber)
                CelldirOrientY(CellPortnumber) = 0
                CelldirOrientZ(CellPortnumber) = 1
        End If
    Next CellPortnumber
    '=====================================
    'BUILD HVACNOZZLE2 Of Division
    '=====================================
    iPortIndex = 2
         FlowDir = DistribFlow_OUT
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                    NThickness, NFlangeWidth, FlowDir, NWidth(2), _
                                                    NDepth(2), CornerRadius, DimBaseOuter, PortStatus, _
                                                    "HNoz2", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

            If oHvacNozzle Is Nothing Then
                strDescritption = "Port Index = 2: Hvac Nozzle object is nothing"
                GoTo ErrorLabel
            End If
            'Position of the nozzle should be the connect point of the nozzle: the Port


            X = NEGLIGIBLE_THICKNESS
            Y = 0
            Z = -(Width / 2 - CellWidth(2) / 2)

            pos.Set X, Y, Z
            Set iDistribPort = oHvacNozzle
            iDistribPort.SetPortLocation pos

            'Direction specified here of the nozzle should be the direction in which pipe will be routed.
            'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
            dir.Set 1, 0, 0
            iDistribPort.SetDirectionVector dir

            CelldirOrient.Set 0, CelldirOrientY(2), CelldirOrientZ(2)
            iDistribPort.SetRadialOrient CelldirOrient

            Set iNozzle = oHvacNozzle
            iNozzle.Length = NozzleLength


            m_outputColl.AddOutput "ClosetoRefCell2", oHvacNozzle
            Set oHvacNozzle = Nothing
            Set iNozzle = Nothing
            Set iDistribPort = Nothing

    '=====================================
    'BUILD HVACNOZZLE3 Of Division
    '=====================================
    iPortIndex = 3
                FlowDir = DistribFlow_OUT
                Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                    NThickness, NFlangeWidth, FlowDir, NWidth(3), _
                                                    NDepth(3), CornerRadius, DimBaseOuter, PortStatus, _
                                                    "HNoz3", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

            If oHvacNozzle Is Nothing Then
                strDescritption = "Port Index = 3: Hvac Nozzle object is nothing"
                GoTo ErrorLabel
            End If

            X = NEGLIGIBLE_THICKNESS
            Y = 0
            Z = -(Width / 2 - CellWidth(2) - CellWidth(3) / 2)

            pos.Set X, Y, Z
            Set iDistribPort = oHvacNozzle
            iDistribPort.SetPortLocation pos
            dir.Set 1, 0, 0
            iDistribPort.SetDirectionVector dir
            CelldirOrient.Set 0, CelldirOrientY(3), CelldirOrientZ(3)
            iDistribPort.SetRadialOrient CelldirOrient

            Set iNozzle = oHvacNozzle
            iNozzle.Length = NozzleLength
            m_outputColl.AddOutput "MiddleCell3", oHvacNozzle
            Set oHvacNozzle = Nothing
            Set iNozzle = Nothing
            Set iDistribPort = Nothing


    '=====================================
    'BUILD HVACNOZZLE4 Of Division
    '=====================================
    iPortIndex = 4
                FlowDir = DistribFlow_OUT
                Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                    NThickness, NFlangeWidth, FlowDir, NWidth(4), _
                                                    NDepth(4), CornerRadius, DimBaseOuter, PortStatus, _
                                                    "HNoz4", PortDepth, CptOffset, False, m_outputColl.ResourceManager)

            If oHvacNozzle Is Nothing Then
                strDescritption = "Port Index = 4: Hvac Nozzle object is nothing"
                GoTo ErrorLabel
            End If

            X = NEGLIGIBLE_THICKNESS
            Y = 0
            Z = Width / 2 - CellWidth(4) / 2

            pos.Set X, Y, Z
            Set iDistribPort = oHvacNozzle
            iDistribPort.SetPortLocation pos
            dir.Set 1, 0, 0
            iDistribPort.SetDirectionVector dir
            CelldirOrient.Set 0, CelldirOrientY(4), CelldirOrientZ(4)
            iDistribPort.SetRadialOrient CelldirOrient

            Set iNozzle = oHvacNozzle
            iNozzle.Length = NozzleLength
            m_outputColl.AddOutput "FarfromRefCell4", oHvacNozzle
            Set oHvacNozzle = Nothing
            Set iNozzle = Nothing
            Set iDistribPort = Nothing

    Dim ObjDivisionRef As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ldiam     As Double

    stPoint.Set 0, -Depth / 2, -Width / 2
    enPoint.Set 0.001, -Depth / 2, -Width / 2
    ldiam = 0.02

    Set ObjDivisionRef = PlaceCylinder(m_outputColl, stPoint, enPoint, ldiam, True)
    m_outputColl.AddOutput "DivisionReference", ObjDivisionRef
    Set ObjDivisionRef = Nothing

    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing
    Set CelldirOrient = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    'Debug.Print "Error encountered while computing Physical representation", vbExclamation, "HRECTFF definition"
    '    Debug.Assert False
    If Len(strDescritption) = 0 Then
        strDescritption = Err.description
    End If
    m_oErrors.Add E_FAIL, MODULE & METHOD, strDescritption, "UserSymbol"
    Err.Raise E_FAIL
End Sub

