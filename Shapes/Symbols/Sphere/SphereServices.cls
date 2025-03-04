VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SphereServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '******************************************************************
' Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    SphereServices.cls
'
'Author
'    MU
'
'Description
'    Definition of SphereServices Symbol
'History
'  4  Nov  MU Creation
'   24Jul2003   VS      Do not clear middle tier errors. This may clear out other errors too.
'
'Notes
'
'    <notes>
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Private Const E_FAIL = -2147467259
Implements IJDUserSymbolServices


Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = "Sphere.SphereServices"
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    On Error GoTo ErrorHandler
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    ' Feed Pump Definition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. "SphereDiameter"
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
    
    'Define inputs for Sphere
     Set Inputs(1) = New IMSSymbolEntities.DInput
     Inputs(1).Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.2
     Inputs(1).Name = "SphereDiameter"
     Inputs(1).Description = "Diameter of the Sphere"
     Inputs(1).DefaultParameterValue = PC

     InputsIf.SetInput Inputs(1), 2
     Set Inputs(1) = Nothing
 
    
    'Define the outputs
    Dim O1 As IMSSymbolEntities.DOutput
    Set O1 = New IMSSymbolEntities.DOutput
    
    O1.Name = "Sphere"
    O1.Description = "Sphere"
    O1.Properties = 0
    
    Dim O2 As IMSSymbolEntities.DOutput
    Set O2 = New IMSSymbolEntities.DOutput
    
    O2.Name = "Point1"
    O2.Description = "Sphere"
    O2.Properties = 0
    
    'Define the representation "Symbolic"
    Dim rep1 As IMSSymbolEntities.DRepresentation
    Set rep1 = New IMSSymbolEntities.DRepresentation

    rep1.Name = "Physical"
    rep1.Description = "Physical Representation of Sphere"
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
    PhysicalRepEval.ProgId = "Sphere.SphereServices"
    
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
    Error.Raise E_FAIL
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' Set definition progId and codebase
    oSymbolDefinition.ProgId = "Sphere.SphereServices"
    oSymbolDefinition.CodeBase = CodeBase
    
    ' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
    
    'return symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    Set oSymbolDefinition = Nothing
    
    Exit Function

ErrorHandler:
    Error.Raise E_FAIL
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
'CREATION OF PHYSICAL REPRESENTATION OF Sphere
'=========================================================================
Private Sub Physical(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Dim Offset As Double
    Dim SphereDia As Double
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oPart As PartFacelets.IJDPart
    
    Set oPart = arrayOfInputs(1)
    SphereDia = arrayOfInputs(2)
    If SphereDia <= 0 Then
        GoTo ErrorHandler
    End If
    '=====================================
    ' CONSTRUCTION OF  Sphere
    '=====================================
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim sphere As IngrGeom3D.Sphere3d
    Dim oPoint3d As IngrGeom3D.Point3d
    Dim circleCenterX As Double, circleCenterY As Double, circleCenterZ As Double
    
    Set oPoint3d = geomFactory.Points3d.CreateByPoint(Nothing, 0, 0, 0)
    m_outputColl.AddOutput "Point1", oPoint3d
    
    circleCenterX = 0#
    circleCenterY = 0#
    circleCenterZ = 0#
    Set sphere = geomFactory.Spheres3d.CreateByCenterRadius(Nothing, _
                                circleCenterX, circleCenterY, circleCenterZ, _
                                SphereDia / 2, True)
                                
    m_outputColl.AddOutput "Sphere", sphere
    
    Set oPoint3d = Nothing
    Set sphere = Nothing
    Set geomFactory = Nothing
    Exit Sub
    
ErrorHandler:
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
    If SphereDia <= 0 Then
        oErrors.Add E_FAIL, "Sphere.SphereServices", "Shape Dimensions should be greater than zero", "ZeroOrNegative"
    End If
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




