VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SemiEllipticalServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    SemiEllipticalServices.cls
'
'Author
'    MU
'
'Description
'    Definition of SemiEllipticalServices Symbol
'History
' 26 march MU Creation
'   02Jun2003   VS     Change the creation of end planes to use the CreateByOuterBdry method,
'                           this allows to check if the boundary has been created properly
'   24Jul2003   VS      Do not clear middle tier errors. This may clear out other errors too.
'   16Oct2003   VS      Added circular edges at the faces for CR36485. SmartSketch will locate points on these edges
'                           Also, set the bSolid flag to false since we are adding planes at start and end. (TR45200)
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const E_FAIL = -2147467259
Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "SemiElliptical.SemiEllipticalServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    ' Feed SemiEllipticalHead Definition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "SemiEllipticalDiameter"
    '          3. "SemiEllipticalThickness"
    ' Representations :
    '           Physical
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    ' Create a new input by new operator
    Dim Inputs(1 To 10) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'Define inputs for SemiElliptical Head
    Dim Index As Integer
    For Index = 1 To 2
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.4
            Inputs(Index).Name = "SemiEllipticalDiameter"
            Inputs(Index).Description = "Diameter of the SemiElliptical Head"
        Else
            PC.UomValue = 0.1
            Inputs(Index).Name = "SemiEllipticalThickness"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 2
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next
 
    
    'Define the outputs
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput
    
    O1.Name = "SemiEllipticalHead"
    O1.Description = "SemiEllipticalHead"
    O1.Properties = 0
    
    Dim O2 As IMSSymbolEntities.DOutput
    Set O2 = New IMSSymbolEntities.DOutput
    
    O2.Name = "Point1"
    O2.Description = "SemiEllipticalHead"
    O2.Properties = 0
    
    Dim E1 As IMSSymbolEntities.DOutput
    Set E1 = New IMSSymbolEntities.DOutput
    E1.Name = "Edge1"
    E1.Description = "Edge1 of SemiEllipticalHead"
    E1.Properties = 0
    
    Dim oLocatePlane As IMSSymbolEntities.DOutput
    Set oLocatePlane = New IMSSymbolEntities.DOutput

    oLocatePlane.Name = "DefaultSurface"
    oLocatePlane.Description = "DefaultSurface used for relationships"
    oLocatePlane.Properties = 0

    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of SemiEllipticalHead"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    ' Set the output
    IJDOutputs.SetOutput O1
    Set O1 = Nothing
    
    IJDOutputs.SetOutput O2
    Set O2 = Nothing

    IJDOutputs.SetOutput E1
    Set E1 = Nothing
    
    IJDOutputs.SetOutput oLocatePlane
    Set oLocatePlane = Nothing
    
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
    PhysicalRepEval.ProgId = "SemiElliptical.SemiEllipticalServices"
    
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
    oSymbolDefinition.ProgId = "SemiElliptical.SemiEllipticalServices"
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
'CREATION OF PHYSICAL REPRESENTATION OF SemiEllipticalHead
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    
    Dim dSemiEllipticalDiameter  As Double
    Dim dSemiEllipticalThickness As Double
    Dim dRadius As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    
    dSemiEllipticalDiameter = arrayOfInputs(2)
    dSemiEllipticalThickness = arrayOfInputs(3)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If dSemiEllipticalDiameter <= 0 Or dSemiEllipticalThickness <= 0 Then
        oErrors.Add E_FAIL, "SemiElliptical.SemiEllipticalServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    End If
   
    '=====================================
    ' CONSTRUCTION OF  SemiEllipticalHead
    '=====================================
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oSphere As IngrGeom3D.Sphere3d
    Dim dCircleCenterX As Double, dCircleCenterY As Double, dCircleCenterZ As Double
    
    dCircleCenterX = 0#
    dCircleCenterY = 0#
    dCircleCenterZ = 0#
    
    Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
    Dim oRevolution As IngrGeom3D.Revolution3d
    Dim PI As Double
   
    PI = 4 * Atn(1)
    dRadius = dSemiEllipticalDiameter / 2
    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                              dCircleCenterX, dCircleCenterY, dCircleCenterZ, 0, 0, 1, 0, dRadius, 0, _
                              (dSemiEllipticalThickness * 2) / dSemiEllipticalDiameter, PI * 1.5, PI / 2)
    Set oRevolution = oGeomFactory.Revolutions3d.CreateByCurve(Nothing, _
                                                    oEllipticalArc, 1, 0, 0, 0, 0, 0, 2 * PI, False)
    
    m_outputColl.AddOutput "SemiEllipticalHead", oRevolution
    
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oPoint3d As IngrGeom3D.Point3d
    Dim oPlane  As IngrGeom3D.Plane3d
    Dim oComplexString3d As IngrGeom3D.ComplexString3d
    
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d

    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, 0, -1, 0, 0, dRadius)
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oCircle)
    m_outputColl.AddOutput "DefaultSurface", oPlane
    m_outputColl.AddOutput "Edge1", oCircle
    Set oPoint3d = Nothing
    Set oCircle = Nothing
    Set oComplexString3d = Nothing
    Set oPlane = Nothing
    Set oRevolution = Nothing
    Set oEllipticalArc = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

'Following method will be removed very soon. This has been
'included to improve the performance. These lines will be removed as soon as
'symbol subsytem allows this to be removed from Symbol Definition .- Raju 05/05/99.

' CMCache custom method to cache the input argument into a parameter contend and the reverse conversion
'It is up to you to find a way to convert your reference data object to a parameter content
Public Sub CMCacheForPart(pInputCM As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)

 If bArgToCache Then
       
        'Need to convert the graphic input pToConvert into a Parameter ( pOutput)
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        
        Dim oPart As IJDPart
        Set oPart = pToConvert
        'MsgBox "Partnumber :" & oPart.PartNumber
        
        ' Create a PC whose value is an identifier of the input
        ' Raju,
        ' the property Part_Number must be retrieved form from the pToConvert argument.
        ' I am hardcoding it
        '
        oPC.Type = igString
        oPC.String = oPart.PartNumber

        Set oPart = Nothing

        ' Always return this PC
        Set pOutput = oPC
        Set oPC = Nothing
 Else
        'Need to convert the cached Parameter pToConvert into your reference data object pOutput
        Dim oPCout As IJDParameterContent
        Dim oPCCache As IJDParameterContent
        Set oPCout = New DParameterContent
        
        ' Here there is three options
        ' o Return a parameter contents, that containts the part_number stored by the pToConvert argument
        '   In this case the edit command will have to retrieve the Part object when needed.
        '   Note : It is better to return a copy of the cached object instead the cached object itself.
        '          This allow to avoid the edition of the cached object.
        ' o Retrieve the Part_number and retrieve your part object with it.
        '   With this solution you can have assoc assertion while when this method is called you are in
        '   in a compute process.
        ' o Get the symbol (or equipment) from the pInputCM, query for the IJDReferencesArg interface,
        '   then get the argument at index 1 it is the SiteProxy.
        '   With this last method, the part has to be passed by reference to the symbol,
        '   but that is what you doing with your design.
        
        ' returning NULL means that the cache method is unable to resolve the cache.
        ' as of now anyway the part is passed as an input argument. this issue of caching
        ' has to be resolved in cycle2.
        
        Set pOutput = Nothing
        
        ' Here is the implementation of the option 1
'        Set oPCCache = pToConvert
'        oPCout.Type = oPCCache.Type
'        oPCout.String = oPCCache.String
'        Set pOutput = oPCout
'        Set oPCout = Nothing
        
        ' Here is the implementation of the option 2
'        Dim oRefDB As GSCADRefDataServices.IJDRefDBGlobal
'        Dim oPart As IJDPart

'        Set oRefDB = New GSCADRefDataServices.RefDBGlobal
'        Set oPart = oRefDB.GetCatalogPart("Storage Tanks", oPCCache.String)
'        Set pOutput = oRefDB
'        Set oRefDB = Nothing
'        Set oPart = Nothing
End If

End Sub






