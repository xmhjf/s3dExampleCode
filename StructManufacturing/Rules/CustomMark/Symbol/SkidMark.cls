VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SkidMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project:
'
'File:
'
'Abstract: implementation of skid mark symbol
'
'Author:
'
'Revision:
'*******************************************************************

Option Explicit

Implements IJDUserSymbolServices
Implements IJMfgGenericMarkInputs

Private Const MODULE = "MfgGenericMarkSymbol.SkidMark(SkidMark.cls)"
Private Const m_progID = "MfgGenericMarkSymbol.SkidMark"

'---------------------------------------------------
'   INPUT Names
'---------------------------------------------------
Private Const INPUT_1 = "PlateSystem"
Private Const INPUT_2 = "UpperFrame"
Private Const INPUT_3 = "LowerFrame"
Private Const INPUT_4 = "LongitudinalReference"
Private Const INPUT_5 = "PlaneReference"
Private Const INPUT_6 = "StartingPlanes"
Private Const INPUT_7 = "GirthOffsets"

'---------------------------------------------------
'   OUTPUT and REPRESENTATION Names
'---------------------------------------------------
Private Const OUTPUT = "SkidMarkCurve"
Private Const REPRESENTATION1 = "Simple" ' used to named the simple physical representation

' Below user defined types are commented as the sending the objects of these types to command causing
' failure in DLLGetClassObject(in SP3DCreateObject) with symbol share configurartion. So collections are adopted
'---------------------------------------------------
'   User Defined Structures
'--------------- ------------------------------------
'Public Type GirthOffsetInput
'    StartingPlane   As Object
'    dOffsetvalue    As Double
'End Type
'
'Public Type InputInfo
'    InputType       As tagGenericMarkInputType
'    DisplayName     As String
'    Prompt          As String
'    DefaultValue    As String
'End Type

Private m_strINPUT_1                            As String
Private m_strINPUT_2                            As String
Private m_strINPUT_3                            As String
Private m_strINPUT_4                            As String
Private m_strINPUT_5                            As String
Private m_strINPUT_6                            As String
Private m_strINPUT_7                            As String

Private m_strPlateSysPrompt                     As String
Private m_strBoundary1_Prompt                   As String
Private m_strBoundary2_Prompt                   As String
Private m_strLONGITUDINALREFERENCE_PROMPT       As String
Private m_strOFFSETVALUE_PROMPT                 As String
Private m_strPLANEREFERENCE_PROMPT              As String
Private m_strOFFSET_DISPLAYNAME                 As String
Private m_strSTARTFRAME_PROMPT                  As String
Private m_strINPUTOFFSETVALUE_PROMPT            As String
Private m_strGirthOffset                        As String
Private m_strStartingPlane                      As String


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    InitializeLocalizeVariables
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: EditOccurence
'
'Interface: IJDUserSymbolServices
'
'Abstract: performs the edition of the passed symbol occurrence
'********************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, _
                                                     ByVal pTransactionMgr As Object) As Boolean
    'TrackText "SkidMark::IJDUserSymbolServices_EditOccurence called", 0
    
    Const METHOD = "IJDUserSymbolServices_EditOccurence"
    On Error GoTo ErrorHandler

    IJDUserSymbolServices_EditOccurence = False
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: GetDefinitionName
'
'Interface: IJDUserSymbolServices
'
'Abstract: get the symbol definition name (progID)
'********************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    'TrackText "SkidMark::IJDUserSymbolServices_GetDefinitionName called", 0
    
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler
    
    IJDUserSymbolServices_GetDefinitionName = m_progID

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InitializeSymbolDefinition
'
'Interface: IJDUserSymbolServices
'
'Abstract: constructs the symbol definition by (re)defining the inputs, outputs,
'          representations, and representation
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(ppSymbolDef As IMSSymbolEntities.IJDSymbolDefinition)
    'TrackText "SkidMark::IJDUserSymbolServices_InitializeSymbolDefinition called", 0

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    'remove all previous Symbol Definition information
    ppSymbolDef.IJDInputs.RemoveAllInput
    ppSymbolDef.IJDRepresentations.RemoveAllRepresentation
    ppSymbolDef.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    '----------------
    'define inputs
    '----------------
    
    'get the input interface of the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = ppSymbolDef

    '-----------------
    'PlateSystem Input
    '-----------------
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput

    oInput.name = INPUT_1
    oInput.Description = INPUT_1
    oInput.Properties = 0
    oInputs.SetInput oInput, 1
    oInput.Reset

    Set oInput = Nothing
    '----------------
    'UpperFrame Input
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_2
    oInput.Description = INPUT_2
    oInput.Properties = 0
    oInputs.SetInput oInput, 2
    oInput.Reset

    Set oInput = Nothing
    '----------------
    'LowerFrame Input
    '----------------
    Set oInput = New IMSSymbolEntities.DInput

    oInput.name = INPUT_3
    oInput.Description = INPUT_3
    oInput.Properties = 0
    oInputs.SetInput oInput, 3
    oInput.Reset

    Set oInput = Nothing
    '----------------------
    'LongitudinalReference
    '----------------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_4
    oInput.Description = INPUT_4
    oInput.Properties = 0 ' igDESCRIPTION_OPTIONAL '0
    oInputs.SetInput oInput, 4
    oInput.Reset

    Set oInput = Nothing
    '----------------
    'PlaneReference
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_5
    oInput.Description = INPUT_5
    oInput.Properties = 0 'igDESCRIPTION_OPTIONAL '0
    oInputs.SetInput oInput, 5
    oInput.Reset

    Set oInput = Nothing
    '----------------
    'StartingPlanes
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = INPUT_6
    oInput.Description = INPUT_6
    oInput.Properties = 0 ' igCOLLECTION_VARIABLE Or igDESCRIPTION_OPTIONAL
    oInputs.SetInput oInput, 6
    oInput.Reset

    Set oInput = Nothing
    '----------------
    'GirthOffsets
    '----------------
    Set oInput = New IMSSymbolEntities.DInput
    
    oInput.name = "GirthOffsets"
    oInput.Description = INPUT_7
    oInput.Properties = igINPUT_IS_A_PARAMETER

    'create an input parameter description
    Dim oPC As IJDParameterContent
    Set oPC = New IMSSymbolEntities.DParameterContent
  
    oPC.Type = igString
    oPC.String = m_strGirthOffset

    oInput.DefaultParameterValue = oPC
    oInputs.SetInput oInput, 7
    oInput.Reset

    Set oInput = Nothing
    Set oPC = Nothing
    
    Set oInputs = Nothing
    '-------------------------------------
    'define outputs and representation
    '-------------------------------------
    'set the outputs
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = ppSymbolDef
    
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation
    
    oRep.RepresentationId = 1  'define a aspect 0 (Simple_physical)
    oRep.name = REPRESENTATION1
    oRep.Description = REPRESENTATION1
    oRep.Properties = igCOLLECTION_VARIABLE

    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep

    oOutputs.Property = igCOLLECTION_VARIABLE

    Set oOutputs = Nothing
    
    oReps.SetRepresentation oRep
    
    Set oRep = Nothing
    Set oReps = Nothing
    
    '------------------------
    'define evaluations
    '------------------------
    Dim oRepEvals As IJDRepresentationEvaluations
    Set oRepEvals = ppSymbolDef
    
    Dim oRepEval As IJDRepresentationEvaluation
    Set oRepEval = New DRepresentationEvaluation
    
    oRepEval.name = REPRESENTATION1
    oRepEval.Description = "Evaluation function identification for the simple representation"
    oRepEval.Properties = igREPRESENTATION_HIDDEN
    oRepEval.Type = igREPRESENTATION_VBFUNCTION
    oRepEval.ProgId = m_progID

    oRepEvals.AddRepresentationEvaluation oRepEval

    Set oRepEval = Nothing
    Set oRepEvals = Nothing
    
CleanUp:

    Set oRep = Nothing
    Set oReps = Nothing
    Set oRepEval = Nothing
    Set oRepEvals = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InstanciateDefinition
'
'Interface: IJDUserSymbolServices
'
'Abstract: instantiates and returns an initialized symbol definition
'********************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
    'TrackText "IJDUserSymbolServices_InstanciateDefinition called", 0

    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IJDSymbolDefinition

    'instanciate the symbol definition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, pResourceMgr)

    'set the definition ProgId and CodeBase
    oSymbolDefinition.ProgId = m_progID
    oSymbolDefinition.CodeBase = CodeBase

    'feed the symbol definition with dynamic information (Inputs, Representation, Outputs)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    'set the definition name
    oSymbolDefinition.name = IJDUserSymbolServices_GetDefinitionName(defParams)

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition

CleanUp:
    Set oSymbolFactory = Nothing
    Set oSymbolDefinition = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InvokeRepresentation
'
'Interface: IJDUserSymbolServices
'
'Abstract: performs the evaluation function associated with the representation
'          with the set of passed inputs as arguments. the function has to compute
'          the symbol graphic outputs that will be attached to the specified
'          Output Collection object

'********************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, _
                                                       ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    'TrackText "SkidMark::IJDUserSymbolServices_InvokeRepresentation called", 0

    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"
    On Error GoTo ErrorHandler
    
    Dim oPlateSystem    As Object
    Dim oUpperFrame     As Object
    Dim oLowerFrame     As Object
    Dim oLongRef        As Object
    Dim oPlaneRef       As Object
    Dim dGirthLength    As Double
        
    '-------------------
    'gather all inputs
    '-------------------
    'TrackText "gather all inputs", 0
    
    On Error Resume Next
    'get the required inputs
    Set oPlateSystem = arrayOfInputs(1)
    
    If oPlateSystem Is Nothing Then
        'MsgBox "No inputs;  exit from symbol"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Set oUpperFrame = arrayOfInputs(2)
    Set oLowerFrame = arrayOfInputs(3)
    Set oLongRef = arrayOfInputs(4)
    Set oPlaneRef = arrayOfInputs(5)
    
    On Error Resume Next
    Dim oRefColl    As Object
    Set oRefColl = arrayOfInputs(6) ' collection of starting planes as reference collection
    
    Dim strCollectionOfgirthOffset  As String
    strCollectionOfgirthOffset = arrayOfInputs(7) ' collection of offset values and starting plane names concatenated into string
    
    Dim oCollection             As Collection
    Dim oRefCollArgs            As IEnumJDArgument
    Dim oStartingPlanesColl     As IJDObjectCollection
    
    Set oStartingPlanesColl = New JObjectCollection
    Set oRefCollArgs = oRefColl
    
    Dim arg1    As IJDArgument
    Dim found   As Long
    
    Do
        oRefCollArgs.Next 1, arg1, found
        If found = 0 Then Exit Do
        
        oStartingPlanesColl.Add arg1.Entity, arg1.Index
        
        Set arg1 = Nothing
    Loop
    
    Set oCollection = GetGirthOffsetInfoCollection(oStartingPlanesColl, strCollectionOfgirthOffset)
    
    If oCollection.Count = 0 Then
        ' exit when collection count is zero
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    '--------------------------------------------------
    '  compute the skid mark for the given inputs
    '--------------------------------------------------
    Dim oResourceManager    As Object
    Dim oOutPutCollection   As IJDOutputCollection
    Set oOutPutCollection = pOutputColl
    
    Set oResourceManager = oOutPutCollection.ResourceManager
    
    RemoveOutput pOutputColl
    
    CreateSymbolOuputs oPlateSystem, oUpperFrame, oLowerFrame, oLongRef, oPlaneRef, oCollection, oResourceManager, pOutputColl
    
    'TrackText "clean up and return", 0

CleanUp:

    Set oPlateSystem = Nothing
    Set oUpperFrame = Nothing
    Set oLowerFrame = Nothing
    Set oLongRef = Nothing
    Set oPlaneRef = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Sub

'--------------------------------------------------
'  compute the skid mark for the given inputs
'--------------------------------------------------
Private Sub CreateSymbolOuputs(oPlateSystem As IJPlateSystem, oUpperFrame As IJPlane, oLowerFrame As IJPlane, oLongRef As Object, oPlaneRef As IJPlane, collOffsets As Collection, oResourceManager As Object, ByRef pOutputColl As IJDOutputCollection)
On Error GoTo ErrorHandler
Const METHOD = "CreateSymbolOuputs"

    '    for each offset input in the collection
    '    get the starting plane
    '    get the next plane
    '    check if the next plane is in the limits of the upper and lower frame
    '    if within the limits
    '    then create skid mark
    
    Dim i       As Long
    
    For i = 1 To collOffsets.Count

        ' In each collection of the collOffsets, first item is Starting plane and second item is OffsetValue
        Dim dOffset As Double
        dOffset = collOffsets.Item(i).Item(2)

        Dim oBoundary1  As IJPlane
        Dim oBoundary2  As IJPlane

        Set oBoundary1 = collOffsets.Item(i).Item(1)

        If oBoundary1 Is Nothing Then
            GoTo NextItem
        End If
        
        Dim oISPGNavigate As ISPGNavigate
        Set oISPGNavigate = oBoundary1
        
        oISPGNavigate.GetReference ReferenceType.Next, NestingLevelType.Same, oBoundary2
        
        ' ??? should we validate if the second boundary is within the limits of upper and lower frame
        
        Dim oComplexString  As IJComplexString

        Set oComplexString = CreateSkidMark(oPlateSystem, oBoundary1, oBoundary2, oLongRef, oPlaneRef, dOffset, oResourceManager)
        '--------------------------------------------------
        'set the output to the complex string
        '--------------------------------------------------
        
        'TrackText "set symbol output", 0
    
        If Not oComplexString Is Nothing Then
            
            InitNewOutput pOutputColl, OUTPUT & i
            
            pOutputColl.AddOutput OUTPUT & i, oComplexString
        End If

        Set oComplexString = Nothing
NextItem:
    Next
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub


Private Function CreateSkidMark(oPlateSystem As IJPlateSystem, oBoundary1 As IJPlane, oBoundary2 As IJPlane, oLongRef As Object, oPlaneRef As IJPlane, dGirthLength As Double, oResourceManager As Object) As IJComplexString
On Error GoTo ErrorHandler
Const METHOD = "CreateSkidMark"

    '---------------------------------------------------
    ' Calculate Mid Plane For Intersection - MidPlane
    '---------------------------------------------------
    Dim oMidPlane   As IJPlane
    Set oMidPlane = GetMidPlane(oBoundary1, oBoundary2)
    
    '-----------------------------------------------
    ' Intersect MidPlane With PlateSystem - Wire
    '-----------------------------------------------
    Dim oWire   As IJWireBody
    Set oWire = PlanePlateSystemIntersection(oMidPlane, oPlateSystem)
    
    '-----------------------------------------------------------------------------------
    ' Get Point Of Intersection Between Wire And Logitudunal Reference - ReferencePoint
    '-----------------------------------------------------------------------------------
    Dim oReferencePoint As IJDPosition
    Set oReferencePoint = GetReferencePoint(oWire, oLongRef)
    
    '-----------------------------------------------------------------------------
    ' From ReferencePoint Get Point At Offset Distance Along Girth - FinalPoint
    '-----------------------------------------------------------------------------
    Dim oFinalPoint As IJDPosition
    Set oFinalPoint = PointAlongGirthFromReference(oReferencePoint, oWire, dGirthLength)
    
    ' If the FinalPoint is nothing skip the execution of remaining code
    If oFinalPoint Is Nothing Then
        Dim oNamedItem As IJNamedItem
        Set oNamedItem = oBoundary1
        StrMfgLogError Err, MODULE, METHOD, "Cannot find a point on the plate with the given offset"
        StrMfgLogError Err, MODULE, METHOD, "Could not create Marking line for Starting Plane " & oNamedItem.name & " and offset " & dGirthLength
        Exit Function
    End If
    
    '-------------------------------------------------------------------------------------------------
    ' Create A Plane At FinalPoint Using Reference Plane - MarkingPlane
    '-------------------------------------------------------------------------------------------------
    Dim oMarkingInputPlane As IJPlane
    Set oMarkingInputPlane = CreateIntersectionMarkInputPlane(oPlaneRef, oFinalPoint)

    '---------------------------------------------------------------------------------------------------
    ' Intersect PlateSystem With MarkingPlane With Boundaries As Boundary1 And Boundary2 - ResultWire
    '---------------------------------------------------------------------------------------------------
    Dim oSkidWire   As IJWireBody
    Set oSkidWire = CreateWireBodyOnPlate(oPlateSystem, oMarkingInputPlane, oBoundary1, oBoundary2)

    ' If the oSkidWire is nothing skip the execution of remaining code
    If oSkidWire Is Nothing Then
        Dim oNI As IJNamedItem
        Set oNI = oBoundary1
        StrMfgLogError Err, MODULE, METHOD, "Could not create Marking line for Starting Plane " & oNI.name & " and offset " & dGirthLength
        Exit Function
    End If
    
    '------------------------------------
    'Convert ResultWire To ComplexString
    '------------------------------------
    Dim oCS As IJComplexString
    Set oCS = ConvertWireBodyToCS(oSkidWire, oResourceManager)

    Set CreateSkidMark = oCS
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Create and plane given the orrt point and the normal. The first three double arguments
'            represent the root point of the plane and the last three double arguments represent the
'            normal of the plane.
'            Use the geometry factory to create a transient plane 3d object and return the plane
'--------------------------------------------------------------------------------------------------
Private Function CreatePlane(x As Double, y As Double, z As Double, nx As Double, ny As Double, nz As Double) As IJPlane
On Error GoTo ErrorHandler
Const METHOD = "CreatePlane"
        
    Dim oGeometryFactory    As IngrGeom3D.GeometryFactory
    Dim oPlane3D            As IngrGeom3D.IPlanes3d
        
    ' create persistent point
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oPlane3D = oGeometryFactory.Planes3d
    Set CreatePlane = oPlane3D.CreateByPointNormal(Nothing, x, y, z, nx, ny, nz)
    
    Set oGeometryFactory = Nothing
    Set oPlane3D = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Given two planes, create a plane at the mid of these two planes. Both the input
'            planes must be parallel to each other. Check for normals to ensure planes are parallel
'--------------------------------------------------------------------------------------------------
Private Function GetMidPlane(oFrame1 As IJPlane, oFrame2 As IJPlane) As IJPlane
On Error GoTo ErrorHandler
Const METHOD = "GetMidPlane"

    Dim dUpperNorX As Double
    Dim dUpperNorY As Double
    Dim dUpperNorZ As Double
    
    Dim dLowerNorX As Double
    Dim dLowerNorY As Double
    Dim dLowerNorZ As Double
    
    oFrame1.GetNormal dUpperNorX, dUpperNorY, dUpperNorZ
    oFrame2.GetNormal dLowerNorX, dLowerNorY, dLowerNorZ
    
    'TODO :create two vectors and compute the angle between them to ensure that the two given planes are co-planar
    
    Dim dUpperX As Double
    Dim dUpperY As Double
    Dim dUpperZ As Double
    
    Dim dLowerX As Double
    Dim dLowerY As Double
    Dim dLowerZ As Double
    
    oFrame1.GetRootPoint dUpperX, dUpperY, dUpperZ
    oFrame2.GetRootPoint dLowerX, dLowerY, dLowerZ
        
    Dim oNorX As Double
    Dim oNorY As Double
    Dim oNorZ As Double
    
    oFrame1.GetNormal oNorX, oNorY, oNorZ
    
    Dim oMidX As Double
    Dim oMidY As Double
    Dim oMidZ As Double
    
    oMidX = (dUpperX + dLowerX) / 2
    oMidY = (dUpperY + dLowerY) / 2
    oMidZ = (dUpperZ + dLowerZ) / 2
    
    Set GetMidPlane = CreatePlane(oMidX, oMidY, oMidZ, oNorX, oNorY, oNorZ)

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Find the intersection between the PlateSystem and the input plane object. Expect result
'           to be a wirebody
'--------------------------------------------------------------------------------------------------
Private Function PlanePlateSystemIntersection(oMidPlane As IJPlane, oPlateSystem As IJPlateSystem) As IJWireBody
On Error GoTo ErrorHandler
Const METHOD = "PlanePlateSystemIntersection"

    Set PlanePlateSystemIntersection = GetIntersection(oPlateSystem, oMidPlane)
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : The reference point is considered to be the point which is a result of intersection
'            between the wire body ( resulting from platesystem and plane intersection) and the
'            logitudinal reference input. Expected that these two objects intersect and return the
'            intersection point
'--------------------------------------------------------------------------------------------------
Private Function GetReferencePoint(oWireBody As IJWireBody, oLongRef As Object) As IJDPosition
On Error GoTo ErrorHandler
Const METHOD = "GetReferencePoint"

    Dim oPoints As IJPointsGraphBody
    Set oPoints = GetIntersection(oWireBody, oLongRef)

    Dim oSGOPointsGraphUtilities As SGOPointsGraphUtilities
    Set oSGOPointsGraphUtilities = New SGOPointsGraphUtilities
    
    Dim oPositionCol As Collection
    
    Set oPositionCol = oSGOPointsGraphUtilities.GetPositionsFromPointsGraph(oPoints)
    Set GetReferencePoint = oPositionCol.Item(1)

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function


'--------------------------------------------------------------------------------------------------
' Abstract : Given a Reference position, a point on the input curve, this routine would return a
'            position at the offset distance along the girth of the input curve
'--------------------------------------------------------------------------------------------------
Private Function PointAlongGirthFromReference(oReferencePosition As IJDPosition, oWireBody As IJWireBody, dGirthOffset As Double) As IJDPosition
On Error GoTo ErrorHandler
Const METHOD = "PointAlongGirthFromReference"

    Dim oSeamUtils  As IJSeamGeomUtils
    Set oSeamUtils = New SeamUtils

    Dim oFinalPos   As IJDPosition
    On Error Resume Next ' Calling the below routine can fail if the girth offset is large
    Set oFinalPos = oSeamUtils.ProjectPointAndMoveAlongWire(oReferencePosition, oWireBody, dGirthOffset, lcomGirth)
    On Error GoTo ErrorHandler
    
    Set PointAlongGirthFromReference = oFinalPos
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : This routine would create a plane at the give root point with normal aligned to the
'            input reference plane
'--------------------------------------------------------------------------------------------------
Private Function CreateIntersectionMarkInputPlane(oPlaneRef As IJPlane, oRootPoint As IJDPosition) As IJPlane
On Error GoTo ErrorHandler
Const METHOD = "CreateIntersectionMarkInputPlane"

    '------------------------------------------
    ' Get PlaneRef Normal   - ReferenceNormal
    '------------------------------------------
    Dim oRefNorX As Double
    Dim oRefNorY As Double
    Dim oRefNorZ As Double
    
    oPlaneRef.GetNormal oRefNorX, oRefNorY, oRefNorZ
    
    Dim dPosX   As Double
    Dim dPosy   As Double
    Dim dPosz   As Double

    oRootPoint.Get dPosX, dPosy, dPosz
    Set CreateIntersectionMarkInputPlane = CreatePlane(dPosX, dPosy, dPosz, oRefNorX, oRefNorY, oRefNorZ)

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
'
' ideally plate system would not be needed
' get the port from plate part
' intersect port with plane
' apply boudaries to resultant intersection curve
' return the bouded wirebody
'--------------------------------------------------------------------------------------------------
Private Function CreateWireBodyOnPlate(oPlateSystem As IJPlateSystem, oMarkingInputPlane As IJPlane, oBoundary1 As IJPlane, oBoundary2 As IJPlane) As IJWireBody
On Error GoTo ErrorHandler
Const METHOD = "CreateWireBodyOnPlatePart"
    
    '------------------------------------------
    ' Intersect Port with oMarkingInputPlane
    '------------------------------------------
    Dim oWireBody As IJWireBody
    Set oWireBody = GetIntersection(oPlateSystem, oMarkingInputPlane)
    
    If oWireBody Is Nothing Then
        StrMfgLogError Err, MODULE, METHOD, "Intersection does not exist between the Plate part and the Plane"
        Exit Function
    End If
    
    '------------------------------------------
    ' Apply boudaries to resultant intersection curve
    '------------------------------------------
    
    Dim oMfgGeomHelper As GSCADStrMfgUtilities.MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oBoundaryColl  As IJDObjectCollection
    Set oBoundaryColl = New JObjectCollection
    
    oBoundaryColl.Add oBoundary1
    oBoundaryColl.Add oBoundary2
    
    Dim dPosX As Double
    Dim dPosy As Double
    Dim dPosz As Double
    
    oMarkingInputPlane.GetRootPoint dPosX, dPosy, dPosz
    
    Dim oRootPoint As IJDPosition
    Set oRootPoint = New DPosition
    oRootPoint.Set dPosX, dPosy, dPosz

    Dim oTrimmedCurve As IJWireBody
    Set oTrimmedCurve = oMfgGeomHelper.TrimCurveByBoundaries(Nothing, oWireBody, oBoundaryColl, oRootPoint)

    Set CreateWireBodyOnPlate = oTrimmedCurve
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract :
'--------------------------------------------------------------------------------------------------
Private Function ConvertWireBodyToCS(oSkidWire As IJWireBody, oResourceManager As Object) As IJComplexString
On Error GoTo ErrorHandler
Const METHOD = "ConvertWireBodyToCS"

    Dim oMfgMGhelper As New MfgMGHelper

    Dim oCS As IJComplexString
    oMfgMGhelper.WireBodyToComplexString oSkidWire, oCS

    ' create persistent complex string
    Dim oGeometryFactory    As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Dim oComplexStrings3d As IComplexStrings3d
    Set oComplexStrings3d = oGeometryFactory.ComplexStrings3d

    Dim oCrvElemets As IJElements
    oCS.GetCurves oCrvElemets

    Dim oCS3d As ComplexString3d
    Set oCS3d = oComplexStrings3d.CreateByCurves(oResourceManager, oCrvElemets)
    
    Set ConvertWireBodyToCS = oCS3d
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : This method returns the types of Input to the command
'--------------------------------------------------------------------------------------------------

Private Sub IJMfgGenericMarkInputs_GetInputs(InputTypes As JCmnShp_CollectionAlias, DisplayNames As JCmnShp_CollectionAlias, Prompts As JCmnShp_CollectionAlias, DefaultValues As JCmnShp_CollectionAlias, InputOptionalInfo As JCmnShp_CollectionAlias)
On Error GoTo ErrorHandler
Const METHOD = "IJMfgGenericMarkInputs_GetInputs"
    'This method returns 4 collections which give information about
        ' collection 1 - Type of the Input - object, codelistedvalue, simple data type, dimensioned value
        ' collection 2 - strings that are to be shown as prompt when the object is selected
        ' collection 3 - Default values/filter criteria for the input objects.
        ' collection 4 - Info about Optional input or not
        
    Set InputTypes = New Collection
    Set Prompts = New Collection
    Set DefaultValues = New Collection
    Set InputOptionalInfo = New Collection
    
    Dim strCriteria As String
    
    '----------------
    'PlateSystem
    '----------------
    InputTypes.Add ObjectType
    Prompts.Add m_strPlateSysPrompt
    DisplayNames.Add m_strINPUT_1
    'get filter criteria for the selection of plate system
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_1)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '----------------
    'UpperFrame
    '----------------
    InputTypes.Add ObjectType
    Prompts.Add m_strBoundary1_Prompt
    DisplayNames.Add m_strINPUT_3
    'get filter criteria for the selection of plate system
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_2)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '----------------
    'LowerFrame
    '----------------
    InputTypes.Add ObjectType
    Prompts.Add m_strBoundary2_Prompt
    DisplayNames.Add m_strINPUT_2
    'get filter criteria for the selection of plate system
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_3)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '----------------------
    'LongitudinalReference
    '----------------------
    InputTypes.Add ObjectType
    Prompts.Add m_strLONGITUDINALREFERENCE_PROMPT
    DisplayNames.Add m_strINPUT_4
    'get filter criteria for the selection of plate system
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_4)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '----------------
    'PlaneReference
    '----------------
    InputTypes.Add ObjectType
    Prompts.Add m_strPLANEREFERENCE_PROMPT
    DisplayNames.Add m_strINPUT_5
    'get filter criteria for the selection of plate system
    strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_5)
    DefaultValues.Add strCriteria
    InputOptionalInfo.Add False 'False means Not optional
    
    '----------------
    'StartingPlanes
    '----------------
    InputTypes.Add ReferenceCollType
    Prompts.Add ""
    DisplayNames.Add ""
    'get filter criteria for the selection of plate system
    DefaultValues.Add ""
    InputOptionalInfo.Add False 'False means Not optional
    
    '----------------
    'GirthOffsets
    '----------------
    InputTypes.Add CollectionType
    Prompts.Add m_strOFFSETVALUE_PROMPT
    DisplayNames.Add m_strOFFSET_DISPLAYNAME
    'get default girth value
    'strCriteria = GetFilterCriteriaOrDefaultvalue(INPUT_6)
    'DefaultValues.Add strCriteria 'CLng(strCriteria)
    InputOptionalInfo.Add False 'False means Not optional
    
    Dim colOffsetInputs As Collection
    Set colOffsetInputs = New Collection
    Dim col1 As Collection
    Set col1 = New Collection
    Dim col2 As Collection
    Set col2 = New Collection

    Dim eInputType As GenericMarkInputType
    eInputType = RangeDataType
        
    ' While adding items to the Col1 and Col2, Add in the order of InputType, DisplayName, Prompt, DefaultValue
    
    col1.Add eInputType
    col1.Add m_strStartingPlane
    col1.Add m_strSTARTFRAME_PROMPT
    col1.Add "2:3"

    eInputType = DimensionedValue
    
    col2.Add eInputType
    col2.Add m_strOFFSET_DISPLAYNAME
    col2.Add m_strINPUTOFFSETVALUE_PROMPT
    col2.Add 0.5

    colOffsetInputs.Add col1
    colOffsetInputs.Add col2
    
    DefaultValues.Add colOffsetInputs
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

'--------------------------------------------------------------------------------------------------
' Abstract : Method to get the filter criteria or the default value of an input
'--------------------------------------------------------------------------------------------------

Private Function GetFilterCriteriaOrDefaultvalue(strInputString As String) As String
On Error GoTo ErrorHandler
Const METHOD = "GetFilterCriteriaOrDefaultvalue"
    
    Select Case strInputString
        Case INPUT_1
            GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case INPUT_2
           GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case INPUT_3
           GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case INPUT_4
           GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
        Case INPUT_5
           GetFilterCriteriaOrDefaultvalue = GetFilterCriteria(strInputString)
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : Method to get the filter criteria for an input
'--------------------------------------------------------------------------------------------------
Private Function GetFilterCriteria(strInputString As String) As String
On Error GoTo ErrorHandler
Const METHOD = "GetFilterCriteria"
    Select Case strInputString
        Case INPUT_1
            'PlateSystem
            GetFilterCriteria = IJPlateSystem & " AND " & " NOT [STFilterFunctions.StructFilterFunctions,ResultOfSplit] "
        Case INPUT_2
           GetFilterCriteria = IHFrame & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsFrameInstersectingPlate] "
        Case INPUT_3
           GetFilterCriteria = IHFrame & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsFrameInstersectingPlate] " & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsFromSameFrameSystemAndAxis] "
        Case INPUT_4
            'ijseam or logicalconnection or frame intersecting input_2 and input_3 at a point
            'should intersect input_1 also
            GetFilterCriteria = IJAppConnectionType & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsLCOrSeamOrFrame] " _
                                & " OR " & IJSeam & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsLCOrSeamOrFrame] " _
                                & " OR " & IHFrame & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsLCOrSeamOrFrame] "
        Case INPUT_5
           GetFilterCriteria = IHFrame & " AND " & "[MfgGenericMarkSymbol.CustomFilterHelper,IsLCOrSeamOrFrame] "
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function

Private Sub RemoveOutput(pOC As IJDOutputCollection)
Const METHOD = "RemoveOutput"
On Error GoTo ErrorHandler
    
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName(REPRESENTATION1)
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
    
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

Private Function GetGirthOffsetInfoCollection(oRefColl As IJDObjectCollection, strCollectionOfgirthOffset As String) As Collection
Const METHOD = "GetGirthOffsetInfoCollection"
On Error GoTo ErrorHandler

    
    Set GetGirthOffsetInfoCollection = New Collection

    Dim lLength     As Long
    Dim pos         As Long
    Dim begin       As Long
    Dim Key         As String
    Dim range       As Long
        
    lLength = Len(strCollectionOfgirthOffset)
    begin = 1
    pos = -1

    'While adding objects to oTempColl, first item should be Starting plane and second item should be OffsetValue
    'String is of the type Framename:Offset;Framename:Offset;.....
    While pos <> lLength

        Dim oTempColl As Collection
        Set oTempColl = New Collection
        
        pos = InStr(begin, strCollectionOfgirthOffset, ";")
        If pos <= 0 Then Exit Function
        range = pos - begin
        
        Key = Mid(strCollectionOfgirthOffset, begin, range)
        begin = pos + 1
            
        'key is Framename:Offset
        
        'get the Framename from key
        Dim strFrameName    As String
        Dim posFrame        As Long
            
        posFrame = InStr(1, Key, ":")
        strFrameName = Mid(Key, 1, posFrame - 1)
            
        Dim strOffsetvalue As String
        strOffsetvalue = Mid(Key, posFrame + 1, Len(Key))
            
        Dim oObject As Object
        For Each oObject In oRefColl
            Dim oNamedItem As IJNamedItem
            Set oNamedItem = oObject
            If oNamedItem.name = strFrameName Then
                oTempColl.Add oObject
                Exit For
            End If
        Next
            
            
        Dim oUom            As IJUomVBInterface
        Dim dbuValue        As Double
    
       
        Set oUom = New UnitsOfMeasureServicesLib.UomVBInterface
        
        dbuValue = oUom.ParseUnit(UNIT_DISTANCE, strOffsetvalue)
       
        oTempColl.Add dbuValue
        
        If oTempColl.Count = 2 Then
            GetGirthOffsetInfoCollection.Add oTempColl
        End If
    Wend
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function


Private Sub InitializeLocalizeVariables()
    Const METHOD = "InitializeLocalizeVariables"
    On Error GoTo ErrorHandler
    
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = InitializeLocalizer(App.Path & "\Resource\", App.EXEName)

    m_strINPUT_1 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_1, "Plate System")
    m_strINPUT_2 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_2, "Upper Frame")
    m_strINPUT_3 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_3, "Lower Frame")
    m_strINPUT_4 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_4, "Longitudinal Reference")
    m_strINPUT_5 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_5, "Plane Reference")
    m_strINPUT_6 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_6, "Starting Planes")
    m_strINPUT_7 = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUT_7, "Girth Offsets")
    
    m_strPlateSysPrompt = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_PLATESYSTEM_PROMPT, "Select a Plate System")
    m_strBoundary1_Prompt = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_BOUNDARY1_PROMPT, "Select Boundary 1")
    m_strBoundary2_Prompt = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_BOUNDARY2_PROMPT, "Select Boundary 2")
    m_strLONGITUDINALREFERENCE_PROMPT = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_LONGITUDINALREFERENCE_PROMPT, "Select Longitudinal Reference")
    m_strPLANEREFERENCE_PROMPT = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_PLANEREFERENCE_PROMPT, "Select Plane Reference")
    m_strOFFSETVALUE_PROMPT = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_OFFSETVALUE_PROMPT, "Input an Offset value for placing the Skid Mark")
    m_strOFFSET_DISPLAYNAME = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_OFFSETVALUE_DISPLAYNAME, "Offset")
    m_strSTARTFRAME_PROMPT = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_STARTFRAME_PROMPT, "Select a frame")
    m_strINPUTOFFSETVALUE_PROMPT = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_INPUTOFFSETVALUE_PROMPT, "Key in offset distance along girth")
    m_strGirthOffset = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_GIRTHOFFSET, "GirthOffset Values")
    m_strStartingPlane = oLocalizer.GetString(IDS_GENERICMARK_SYMBOL_STARTINGPLANE, "Starting Plane")


    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Sub

Private Function InitializeLocalizer(strPath As String, strFileName As String) As IJLocalizer
    Const METHOD = "InitializeLocalizer"
    On Error GoTo ErrorHandler
     
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = CreateObject("IMSLocalizer.Localizer")
    
    If IsInDebugMode = True Then
      oLocalizer.Initialize "M:\SharedContent\Bin\StructManufacturing\Rules\Release\Resource" & strFileName
    Else
       oLocalizer.Initialize strPath & strFileName
    End If
    
    Set InitializeLocalizer = oLocalizer
Exit Function
ErrorHandler:
   Err.Clear
End Function

Private Function IsInDebugMode() As Boolean
    On Error GoTo ErrorHandler
'     If the program is compiled, the following
'     Debug statement has been removed so it will
'     not generate an error.
    Debug.Print 1 / 0
    IsInDebugMode = False
    Exit Function
ErrorHandler:
'   We got an error so the Debug statement must
'   be working.
    IsInDebugMode = True
   Err.Clear
End Function

Private Function IJMfgGenericMarkInputs_GetValuesInRange(ByVal pDispObj1 As Object, ByVal pDispObj2 As Object) As IJElements

    Set IJMfgGenericMarkInputs_GetValuesInRange = New JObjectCollection
    
    'get all the frames between reference plane 1 and reference plane 2
    Dim oFrame1 As IHFrame
    Set oFrame1 = pDispObj1
    
    Dim oFrame2 As IHFrame
    Set oFrame2 = pDispObj2

    Dim oFrameAxis1 As IHFrameAxis
    Set oFrameAxis1 = oFrame1.FrameAxis

    Dim oRefPlaneColl As IJDRefPlaneCollection
    Set oRefPlaneColl = oFrameAxis1.RefPlaneCollection

    Dim oPos1 As Double
    Dim oPos2 As Double
    oPos1 = oFrame1.PositionOnAxis
    oPos2 = oFrame2.PositionOnAxis

    Dim oPos As Double

    If oPos1 > oPos2 Then
        oPos = oPos1
        oPos1 = oPos2
        oPos2 = oPos
    End If

    Dim j As Integer

    For j = 1 To oRefPlaneColl.Count
        Dim oFrame As IHFrame
        Set oFrame = oRefPlaneColl.Item(j)
        oPos = oFrame.PositionOnAxis
        If oPos >= oPos1 And oPos < oPos2 Then
            IJMfgGenericMarkInputs_GetValuesInRange.Add oFrame
        End If
    Next
End Function


'******************************************************************************
' Routine: IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs
'
' Abstract: Get the parts or systems to be marked from the symbol inputs
'******************************************************************************
Private Function IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs(ByVal oSymbolInputColl As JCmnShp_CollectionAlias) As IJDObjectCollection
    Const METHOD = "IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs"
    On Error GoTo ErrorHandler

    Dim oPlateSystem As IJPlateSystem
    
    ' The first input in the symbol is plate system
    Set oPlateSystem = oSymbolInputColl.Item(1)
    
    If Not oPlateSystem Is Nothing Then
        Dim oConnPlatesColl         As IJDObjectCollection
        Set oConnPlatesColl = New JObjectCollection
        oConnPlatesColl.Add oPlateSystem
        
        Set IJMfgGenericMarkInputs_GetPartsOrSystemsToMarkFromInputs = oConnPlatesColl
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
End Function
