VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransElementServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'******************************************************************
' Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'File
'    TransElementServices.cls
'
'Author
'    MU
'
'Description
'    Definition of TransElementServices Symbol
'History
'   12 April Mule Creation
'   02Jun2003   VS     Change the creation of end planes to use the CreateByOuterBdry method,
'                           this allows to check if the boundary has been created properly
'   24Jul2003   VS     Do not clear middle tier errors. This may clear out other errors too.
'   15Oct2003   VS      Added outputs for more edges
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   20.Sep.2005     kkk         TR: 78321(Changing the Eccentiricity from Z-Direction
'                                                     to Y-Direction)
'   20.Sep.2005     kkk         TR: 69313(Generating rulesurfaces based on consistent
'                                                     nature.The curves are broken down in elements
'                                                     which are consistent and in same number(4)
'                                                     on the 2 curves. The start points are also
'                                                     consistent to generate uniform quality.)
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices
Private Const E_FAIL = -2147467259

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "TransElement.TransElementServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    'Feed TransitionElement Defnition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "A"
    '          3. "B"
    '          4. "C"
    '          5. "D"
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
    Dim Inputs(1 To 10) As IMSSymbolEntities.DInput
    
    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
    PC.Type = igValue
    
    'Define inputs for TransitionElement
    Dim Index As Integer
    For Index = 1 To 4
        Set Inputs(Index) = New IMSSymbolEntities.DInput
        Inputs(Index).Properties = igINPUT_IS_A_PARAMETER
        If (Index = 1) Then
            PC.UomValue = 0.12
            Inputs(Index).Name = "A"
        ElseIf (Index = 2) Then
            PC.UomValue = 0.12
            Inputs(Index).Name = "B"
        ElseIf (Index = 3) Then
            PC.UomValue = 0.1
            Inputs(Index).Name = "C"
        Else
            PC.UomValue = 0.08
            Inputs(Index).Name = "D"
        End If
        Inputs(Index).DefaultParameterValue = PC
    Next
    
    For Index = 1 To 4
        InputsIf.SetInput Inputs(Index), Index + 1
        Set Inputs(Index) = Nothing
    Next
    
    'Define the outputs
    Dim oPoint1 As IMSSymbolEntities.DOutput
    Set oPoint1 = New IMSSymbolEntities.DOutput
    
    oPoint1.Name = "Point1"
    oPoint1.Description = "TransitionElement"
    oPoint1.Properties = 0
    
    Dim oPoint2 As IMSSymbolEntities.DOutput
    Set oPoint2 = New IMSSymbolEntities.DOutput
    
    oPoint2.Name = "Point2"
    oPoint2.Description = "TransitionElement"
    oPoint2.Properties = 0
    
    Dim oSurface1 As IMSSymbolEntities.DOutput
    Set oSurface1 = New IMSSymbolEntities.DOutput
    
    oSurface1.Name = "Surface1"
    oSurface1.Description = "Surface1 of TransitionElement"
    oSurface1.Properties = 0
    
    Dim oSurface2 As IMSSymbolEntities.DOutput
    Set oSurface2 = New IMSSymbolEntities.DOutput
    
    oSurface2.Name = "Surface2"
    oSurface2.Description = "Surface2 of TransitionElement"
    oSurface2.Properties = 0
    
    Dim oSurface3 As IMSSymbolEntities.DOutput
    Set oSurface3 = New IMSSymbolEntities.DOutput
    
    oSurface3.Name = "Surface3"
    oSurface3.Description = "Surface3 of TransitionElement"
    oSurface3.Properties = 0
    
    Dim oSurface4 As IMSSymbolEntities.DOutput
    Set oSurface4 = New IMSSymbolEntities.DOutput
    
    oSurface4.Name = "Surface4"
    oSurface4.Description = "Surface4 of TransitionElement"
    oSurface4.Properties = 0
    
    Dim oSurface5 As IMSSymbolEntities.DOutput
    Set oSurface5 = New IMSSymbolEntities.DOutput
    
    oSurface5.Name = "Surface5"
    oSurface5.Description = "Surface5 of TransitionElement"
    oSurface5.Properties = 0
    
    Dim oSurface6 As IMSSymbolEntities.DOutput
    Set oSurface6 = New IMSSymbolEntities.DOutput
    
    oSurface6.Name = "Surface6"
    oSurface6.Description = "Surface6 of TransitionElement"
    oSurface6.Properties = 0

    Dim E1 As IMSSymbolEntities.IJDOutput
    Set E1 = New IMSSymbolEntities.DOutput
    E1.Name = "Edge1"
    E1.Description = "Edge1 of TransitionElement"
    E1.Properties = 0

    Dim E2 As IMSSymbolEntities.IJDOutput
    Set E2 = New IMSSymbolEntities.DOutput
    E2.Name = "Edge2"
    E2.Description = "Edge2 of TransitionElement"
    E2.Properties = 0


    Dim E3 As IMSSymbolEntities.IJDOutput
    Set E3 = New IMSSymbolEntities.DOutput
    E3.Name = "Edge3"
    E3.Description = "Edge3 of TransitionElement"
    E3.Properties = 0

    Dim E4 As IMSSymbolEntities.IJDOutput
    Set E4 = New IMSSymbolEntities.DOutput
    E4.Name = "Edge4"
    E4.Description = "Edge4 of TransitionElement"
    E4.Properties = 0

    Dim E5 As IMSSymbolEntities.IJDOutput
    Set E5 = New IMSSymbolEntities.DOutput
    E5.Name = "Circle5"
    E5.Description = "Circle5 of TransitionElement"
    E5.Properties = 0

    
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of TransitionElement"
    rep1.Properties = igREPRESENTATION_ISVBFUNCTION
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    rep1.RepresentationId = SimplePhysical

    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Set IJDOutputs = rep1

    ' Set the output
    IJDOutputs.SetOutput oPoint1
    Set oPoint1 = Nothing
    
    IJDOutputs.SetOutput oPoint2
    Set oPoint2 = Nothing
    
    IJDOutputs.SetOutput oSurface1
    Set oSurface1 = Nothing
    
    IJDOutputs.SetOutput oSurface2
    Set oSurface2 = Nothing
    
    IJDOutputs.SetOutput oSurface3
    Set oSurface3 = Nothing
    
    IJDOutputs.SetOutput oSurface4
    Set oSurface4 = Nothing
    
    IJDOutputs.SetOutput oSurface5
    Set oSurface5 = Nothing
    
    IJDOutputs.SetOutput oSurface6
    Set oSurface6 = Nothing
    
    IJDOutputs.SetOutput E1
    Set E1 = Nothing

    IJDOutputs.SetOutput E2
    Set E2 = Nothing

    IJDOutputs.SetOutput E3
    Set E3 = Nothing

    IJDOutputs.SetOutput E4
    Set E4 = Nothing

    IJDOutputs.SetOutput E5
    Set E5 = Nothing

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
    PhysicalRepEval.ProgId = "TransElement.TransElementServices"
    
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
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "TransElement.TransElementServices"
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

Private Sub Physical(ByRef arrayOfInputs())
On Error GoTo ErrorHandler
    Dim a As Double, b As Double, c As Double, d As Double, dDiagonal As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim arrPoints(0 To 16) As Double
    Dim oPart As PartFacelets.IJDPart
    Set oPart = arrayOfInputs(1)
    Dim iOutput As Integer
    iOutput = 0
    Dim PI As Double
    PI = 4 * Atn(1)
    a = arrayOfInputs(2)
    b = arrayOfInputs(3)
    c = arrayOfInputs(4)
    d = arrayOfInputs(5)
    dDiagonal = Sqr(a * a + b * b)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If a <= 0 Or b <= 0 Or c <= 0 Or d <= 0 Then
        oErrors.Add E_FAIL, "TransElement.TransElementServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
        GoTo ErrorHandler
    End If
    
    Dim CentreBx As Double, CentreBy As Double, CentreBz As Double
    CentreBx = 0#
    CentreBy = 0#
    CentreBz = 0#
    
    Dim CentreTx As Double, CentreTy As Double, CentreTz As Double
    CentreTx = c
    CentreTy = 0#
    CentreTz = 0#
    
    Dim StartBx As Double, StartBy As Double, StartBz As Double
    StartBx = 0#
    StartBy = 0#
    StartBz = dDiagonal / 2
    
    Dim StartTx As Double, StartTy As Double, StartTz As Double
    StartTx = c
    StartTy = 0#
    StartTz = d / 2

    '=====================================
    ' CONSTRUCTION OF  Transition Element
    '=====================================
    Dim geomFactory As New GeometryFactory
    Dim oSurface As RuledSurface3d
    Dim oCircle3d As Circle3d
    Dim oFrontFace As Plane3d
    Dim oFrontBndry As Circle3d
    Dim oPlane As Plane3d
    Dim oLineString As Line3d
    Dim oPoint3d As Point3d
    Dim line As Line3d
    Dim oArc3d As Arc3d
    Dim ang As Double
    Dim Cx As Double
    Dim Cy As Double
    ang = PI / 4
    Cx = (d / 2) * Cos(ang)
    Cy = (d / 2) * Sin(ang)
        
    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d   'OutPut 1
    
    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, c, 0, 0)
    m_outputColl.AddOutput "Point2", oPoint3d   'OutPut 2

    arrPoints(0) = 0
    arrPoints(1) = -b / 2
    arrPoints(2) = -a / 2
    arrPoints(3) = 0
    arrPoints(4) = -b / 2
    arrPoints(5) = a / 2
    arrPoints(6) = 0
    arrPoints(7) = b / 2
    arrPoints(8) = a / 2
    arrPoints(9) = 0
    arrPoints(10) = b / 2
    arrPoints(11) = -a / 2
    arrPoints(12) = 0
    arrPoints(13) = -b / 2
    arrPoints(14) = -a / 2
    'Creation of The Vertical - Right Hand Surface
    Set oLineString = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
    Set oArc3d = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, c, 0, 0, _
                                                        c, -Cx, -Cy, _
                                                        c, -Cx, Cy)
    Set oSurface = geomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oLineString, oArc3d, False)
    m_outputColl.AddOutput "Surface1", oSurface  'OutPut 3
     
    
'Creation of The Horizontal - Top Surface
    Set oLineString = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
    Set oArc3d = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, c, 0, 0, _
                                                        c, -Cx, Cy, _
                                                        c, Cx, Cy)
    Set oSurface = geomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oLineString, oArc3d, False)
    m_outputColl.AddOutput "Surface2", oSurface  'OutPut 4
     
'Creation of The Vertical - Left Surface
    Set oLineString = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
    Set oArc3d = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, c, 0, 0, _
                                                        c, Cx, Cy, _
                                                        c, Cx, -Cy)
    Set oSurface = geomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oLineString, oArc3d, False)
    m_outputColl.AddOutput "Surface3", oSurface  'OutPut 5
   
'Creation of The Horizontal - Bottom Surface
    Set oLineString = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), arrPoints(12), arrPoints(13), arrPoints(14))
    Set oArc3d = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, c, 0, 0, _
                                                        c, Cx, -Cy, _
                                                        c, -Cx, -Cy)
    Set oSurface = geomFactory.RuledSurfaces3d.CreateByCurves(Nothing, oLineString, oArc3d, False)
    
    m_outputColl.AddOutput "Surface4", oSurface  'OutPut 6
 
'Creation of Rectangular Surface

    arrPoints(0) = 0
    arrPoints(1) = -b / 2
    arrPoints(2) = -a / 2
    arrPoints(3) = 0
    arrPoints(4) = b / 2
    arrPoints(5) = -a / 2
    arrPoints(6) = 0
    arrPoints(7) = b / 2
    arrPoints(8) = a / 2
    arrPoints(9) = 0
    arrPoints(10) = -b / 2
    arrPoints(11) = a / 2
    arrPoints(12) = 0
    arrPoints(13) = -b / 2
    arrPoints(14) = -a / 2
    Set oPlane = geomFactory.Planes3d.CreateByPoints(Nothing, 5, arrPoints)
    
    m_outputColl.AddOutput "Surface5", oPlane   'Output 7
     
''Creation of Circular Face at front
    Set oCircle3d = geomFactory.Circles3d.CreateBy3Points(Nothing, c, 0, d / 2, _
                                                                                     c, -d / 2, 0, _
                                                                                     c, 0, -d / 2)
    Set oFrontFace = geomFactory.Planes3d.CreateByOuterBdry(Nothing, oCircle3d)
    m_outputColl.AddOutput "Surface6", oFrontFace   'Output 8
  
  'Creation of Edges
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(0), arrPoints(1), arrPoints(2), arrPoints(3), arrPoints(4), arrPoints(5))
    m_outputColl.AddOutput "Edge1", line    'Output 9
    Set line = Nothing

    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(3), arrPoints(4), arrPoints(5), arrPoints(6), arrPoints(7), arrPoints(8))
    m_outputColl.AddOutput "Edge2", line    'Output 10
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(6), arrPoints(7), arrPoints(8), arrPoints(9), arrPoints(10), arrPoints(11))
    m_outputColl.AddOutput "Edge3", line    'Output 11
    Set line = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPoints(9), arrPoints(10), arrPoints(11), arrPoints(0), arrPoints(1), arrPoints(2))
    m_outputColl.AddOutput "Edge4", line    'Output 12
    Set line = Nothing
    'Creation of the Circular Front Face
    Set oFrontBndry = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                     c, 0, 0, 1, 0, 0, d / 2)
    m_outputColl.AddOutput "Circle5", oFrontBndry    'Output 13
    
    Set oSurface = Nothing
    Set oPlane = Nothing
    Set oLineString = Nothing
    Set oArc3d = Nothing
    Set oCircle3d = Nothing
    Set geomFactory = Nothing
    Set oFrontFace = Nothing
     
    Exit Sub


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






