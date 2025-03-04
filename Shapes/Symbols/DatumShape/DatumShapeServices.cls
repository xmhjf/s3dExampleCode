VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatumShapeServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
'Copyright (c) 2004,2005 Intergraph Corporation. All rights reserved.
'
'File
'    DatumShapeServices.cls
'
'Author
'    JFF
'
'Description
'    Definition of Datum Shape Symbol
'    A Datum Shape is the equivlent of PDS Datum Point of PDS.
'    It is a reference and not a geometrical shape.
'    It allows user to place Nozzles relative to it.
'
'History
'   23 Dec 2004 JFF Creation
'   12 Jan 2005 JFF DI 69531 Need to add the base (edge + plane)
'   10 Jun 2006 GUK TR-CP·99196 removed the Edge + Plane + Point from definition. did not bumpup the Version of the DLL
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const E_FAIL = -2147467259

Private m_outputColl As Object

Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "DatumShape.DatumShapeServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    'Feed RightCircularCone Defnition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "ConeLength"
    ' Representations :
    '           Physical
    
    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
   
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 2) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'Define inputs for DatumShape
    Dim Index As Integer
    For Index = 1 To 1
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.12
            Inputs(Index).Name = "ConeLength"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 1
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next
    
    'Define the outputs
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput
    
    O1.Name = "Cone"
    O1.Description = "Cone"
    O1.Properties = 0
    
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of Cone"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    'Set the outputs
    IJDOutputs.SetOutput O1
    Set O1 = Nothing
        
    'Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation rep1

    Set rep1 = Nothing
    Set RepsIf = Nothing
    Set IJDOutputs = Nothing
    
    'Define evaluation for Physical representation
    Dim PhysicalRepEval As IJDRepresentationEvaluation
    Set PhysicalRepEval = New DRepresentationEvaluation
    PhysicalRepEval.Name = "Physical"
    PhysicalRepEval.Description = "Physical representation"
    PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
    PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
    PhysicalRepEval.ProgId = "DatumShape.DatumShapeServices"
    
    'Set the evaluations for the Symbolic and Physical representations on the definition
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

    Exit Sub

ErrorHandler:
  Err.Raise E_FAIL
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "DatumShape.DatumShapeServices"
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal RepName As String, ByVal OutputColl As Object, arrayOfInputs() As Variant)
On Error GoTo ErrorHandler
    Set m_outputColl = OutputColl
    If StrComp(RepName, "Physical") = 0 Then
        Physical arrayOfInputs
    End If
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'=========================================================================
'CREATION OF PHYSICAL REPRESENTATION OF DatumShape
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    
    On Error GoTo ErrorHandler
    
    Dim Offset As Double
    Dim ConeLength As Double
    Dim ConeMajorDiameter As Double
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oPart As PartFacelets.IJDPart
    
    Set oPart = arrayOfInputs(1)
    
    ConeLength = arrayOfInputs(2)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If ConeLength <= 0 Then
        Dim strx As String
        strx = "ConeLength=" & CStr(ConeLength)
        oErrors.Add E_FAIL, "DatumShape.DatumShapeServices", strx, "ZeroOrNegative"
        GoTo ErrorHandler
    End If
    
    ConeMajorDiameter = ConeLength / ((Sqr(5) + 1) / 2)
    
    Dim CentreBx As Double, CentreBy As Double, CentreBz As Double
    CentreBx = 0#
    CentreBy = 0#
    CentreBz = 0#
    
    Dim CentreTx As Double, CentreTy As Double, CentreTz As Double
    CentreTx = ConeLength
    CentreTy = 0#
    CentreTz = 0#
    
    Dim StartBx As Double, StartBy As Double, StartBz As Double
    StartBx = 0#
    StartBy = 0#
    StartBz = ConeMajorDiameter / 2
    
    Dim StartTx As Double, StartTy As Double, StartTz As Double
    StartTx = ConeLength
    StartTy = 0#
    StartTz = 0#

    '=====================================
    ' CONSTRUCTION OF  Cone
    '=====================================
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim Cone As IngrGeom3D.Cone3d
    Dim BaseEdge As IngrGeom3D.Circle3d
    Dim BasePlane As IngrGeom3D.Plane3d
    
    Set Cone = geomFactory.Cones3d.CreateBy4Pts(Nothing, _
                                CentreBx, CentreBy, CentreBz, _
                                CentreTx, CentreTy, CentreTz, _
                                StartBx, StartBy, StartBz, _
                                StartTx, StartTy, StartTz, _
                                True)
    Cone.Solid = True
    
    m_outputColl.AddOutput "Cone", Cone
            
    Set Cone = Nothing
    Set geomFactory = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub





