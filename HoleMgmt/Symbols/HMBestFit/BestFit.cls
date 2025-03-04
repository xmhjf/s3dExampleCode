VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BestFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project: Hole Management - HoleMgmt\Middle\Symbols\HMBestFit
'
'File: BestFit.cls
'
'Abstract: implementation of BestFit symbol
'
'Author: pmlivingston@gscad.com
'
'Revision:
'     02/22/01 - pmlivingston@gscad.com - initial release
' 04/12/2002  CRS  Removed transformation of ComplexString because the
'                  transformtion now performed by the BestFitHelper on
'                  the generation of the ComplexString.
' 04/15/2002  sypark  Removed transformation of ComplexString for equipment hole.
' 06/09/03  CRS  Pass working plane to DefinitionRule to enable utilities
'                to compute projections for pipe turn features
' 09/09/03  CRS  Modified InvokeRepresentation to consider the special
'                case of pipe sleeve catalog fitting.
' 05/23/06  CRS  Get the Fitting type from IJHoleFitting
'
'Note: ReportAndRaiseUnanticipatedError method is located in
'      CommonApp/Client/Bin/GSCADUtilities.dll and may have to
'      be replaced
'*******************************************************************

Option Explicit

Private Const MODULE = "HMBestFit.BestFit(BestFit.cls)"

Private Const m_progID = "HMBestFit.BestFit"

Private Const INPUT_1 = "HoleTraceAE"
Private Const OUTPUT_1 = "Trace_1"
Private Const REPRESENTATION_1 = "BestFit"

Private m_oErrors As New IMSErrorLog.JServerErrors
Private Const ERR_CMPLXSTRING = -2147220469
Private Const ERR_NONCIRCUALRPROJECTION = -2147220467

Implements IJDUserSymbolServices

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
    TrackText "BestFit::IJDUserSymbolServices_EditOccurence called", 0
    
    Const Method = "IJDUserSymbolServices_EditOccurence"
    On Error GoTo ErrorHandler

    IJDUserSymbolServices_EditOccurence = False
    
    Exit Function

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
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
    TrackText "BestFit::IJDUserSymbolServices_GetDefinitionName called", 0
    
    Const Method = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler
    
    IJDUserSymbolServices_GetDefinitionName = m_progID

    Exit Function
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
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
    TrackText "BestFit::IJDUserSymbolServices_InitializeSymbolDefinition called", 0
    
    Const Method = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    'remove all previous Symbol Definition information
    ppSymbolDef.IJDInputs.RemoveAllInput
    ppSymbolDef.IJDRepresentations.RemoveAllRepresentation
    ppSymbolDef.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    '
    'define inputs
    '
    
    'get the input interface of the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = ppSymbolDef

    'create an input parameter description
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput

    'HoleTraceAE
    oInput.Name = INPUT_1
    oInput.Description = INPUT_1
    oInput.Properties = 0
    oInputs.SetInput oInput, 1
    oInput.Reset

    Set oInput = Nothing
    Set oInputs = Nothing
    
    '
    'define outputs and representation
    '
    
    'set the outputs
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = ppSymbolDef
    
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation
    
    oRep.RepresentationId = 1  'define a aspect 0 (Simple_physical)
    oRep.Name = REPRESENTATION_1
    oRep.Description = REPRESENTATION_1
    oRep.Properties = igREPRESENTATION_VBFUNCTION

    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep

    Dim oOutput As IMSSymbolEntities.IJDOutput
    Set oOutput = New IMSSymbolEntities.DOutput

    oOutput.Properties = 0
    oOutput.Name = OUTPUT_1
    oOutput.Description = OUTPUT_1
    oOutputs.SetOutput oOutput

    Set oOutput = Nothing
    Set oOutputs = Nothing
    
    oReps.SetRepresentation oRep
    
    Set oRep = Nothing
    Set oReps = Nothing
    
    '
    'define evaluations
    '
    
    Dim oRepEvals As IJDRepresentationEvaluations
    Set oRepEvals = ppSymbolDef
    
    Dim oRepEval As IJDRepresentationEvaluation
    Set oRepEval = New DRepresentationEvaluation
    
    oRepEval.Name = "BestFit"
    oRepEval.Description = "evaluation function identification for the hole trace representation"
    oRepEval.Properties = igREPRESENTATION_HIDDEN
    oRepEval.Type = igREPRESENTATION_VBFUNCTION
    oRepEval.ProgId = m_progID

    oRepEvals.AddRepresentationEvaluation oRepEval

    Set oRepEval = Nothing
    Set oRepEvals = Nothing
    
Cleanup:
    Set oInput = Nothing
    Set oInputs = Nothing
    Set oOutput = Nothing
    Set oOutputs = Nothing

    Set oRep = Nothing
    Set oReps = Nothing
    Set oRepEval = Nothing
    Set oRepEvals = Nothing
    
    Exit Sub

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
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
    TrackText "BestFit::IJDUserSymbolServices_InstanciateDefinition called", 0
    
    Const Method = "IJDUserSymbolServices_InstanciateDefinition"
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
    oSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition

Cleanup:
    Set oSymbolFactory = Nothing
    Set oSymbolDefinition = Nothing

    Exit Function
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
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
'
'Outline:
'    * gather inputs
'    * check the hole trace for state. if in critical (2) no further processing
'      needs to be done
'    * check the hole trace for symbol status. the hole trace will keep a flag
'      indicating if the symbol needs to be recomputed. if the status is set
'      to no recompute, no further processing needs to be done
'    * get the structure and outfitting ports from the hole trace
'    * create the helpers (IJHoleTraceSymbolHelper and IJHMBestFitHelper)
'    * set all inputs to the IJHoleTraceSymbolHelper to determine which of the
'      inputs have been added or modified
'    * use the IJHoleTraceSymbolHelper to filter the different inputs
'    * if this is an existing symbol, check to see if user has only added new
'      outfitting that is already inside hole trace. if so, no futher processing
'      needs to be done
'    * set the inputs to the IJHMBestFitHelper
'    * use the IJHMBestFitHelper to calculate the intersection points
'    * use the IJHMBestFitHelper to determine the vertices of the hole trace
'    * use the IJHMBestFitHelper to determine the shape of the hole trace
'    * use the IJHMBestFitHelper to generate the hole trace geometry
'    * set the output to the complex string built by the IJHMBestFitHelper
'    * set the hole trace state to good (0)
'    * clean up and return
'********************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, _
                                                       ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    TrackText "BestFit::IJDUserSymbolServices_InvokeRepresentation called", 0

    Const Method = "IJDUserSymbolServices_InvokeRepresentation"
    On Error GoTo ErrorHandler
    
    Dim oHoleTrace As IJHoleTraceAE
    Dim oStructure As Object
    Dim oOutfitting As IMSCoreCollections.IJDObjectCollection
    
'******************************************************************************************************
'gather all inputs
'******************************************************************************************************
    TrackText "gather all inputs", 0

    'get the required inputs
    Set oHoleTrace = arrayOfInputs(1)
    
'******************************************************************************************************
'check the hole trace for state. if in critical (2) no further processing needs to be done
'******************************************************************************************************
    TrackText "check the hole trace state", 0

    Dim state As HMHoleState
    state = oHoleTrace.state
    If state = HM_StateCritical Then GoTo Cleanup

'******************************************************************************************************
'check the hole trace for symbol status. the hole trace will keep a flag
' indicating if the symbol needs to be recomputed. if the status is set
' to no recompute, no further processing needs to be done
'******************************************************************************************************
    TrackText "check the symbol recompute status", 0

    Dim bRecompute As Boolean
    bRecompute = oHoleTrace.SymbolRecompute
    If Not bRecompute Then GoTo Cleanup

'******************************************************************************************************
'    * get the structure and outfitting from the hole trace
'******************************************************************************************************
    TrackText "get the structure and outfitting from the hole trace", 0

    Set oStructure = oHoleTrace.GetParentStructure
    Set oOutfitting = oHoleTrace.GetParentOutfitting
    
'******************************************************************************************************
'    * ComputeTrace ( included with simplification )
'******************************************************************************************************
    TrackText "Compute Trace", 0
    
    TrackText "Get the Working Plane", 0
    'Get the Working Plane
    Dim oWorkingPlane As IJPlane
    Set oWorkingPlane = oHoleTrace.GetWorkingPlane
    
    If oWorkingPlane Is Nothing Then
        'MsgBox "Could not determine working plane, please sketch hole" ', vbOKOnly, "Hole Management"
        GoTo Cleanup
    End If
    
    Dim bProjectionsComputed As Boolean
    bProjectionsComputed = False
    Dim oRepresentOutfitting As IMSCoreCollections.IJDObjectCollection
    Set oRepresentOutfitting = New IMSCoreCollections.JObjectCollection
    
    Dim oOutfittingsWithNoHoleAspect As IMSCoreCollections.IJDObjectCollection
    Set oOutfittingsWithNoHoleAspect = New IMSCoreCollections.JObjectCollection

    
    Dim oHMMiddleHelper As New CHMMiddleHelper
    Dim oOutfittingObj As Object
    
    For Each oOutfittingObj In oOutfitting
    Dim bSupportHoleAspect As Boolean
        bSupportHoleAspect = False
        If TypeOf oOutfittingObj Is IJPort Then
            'If object is equipment, convert equipment to port, and returns port object
            Dim oPort As IJPort
            Dim oConnectable As IJConnectable
            
            Set oPort = oOutfittingObj
            Set oConnectable = oPort.Connectable
            If TypeOf oConnectable Is IJEquipment Then
                Set oOutfittingObj = oConnectable
            End If
        End If

        On Error Resume Next
        bSupportHoleAspect = oHMMiddleHelper.DoesObjectSupportHoleAspect(oOutfittingObj)
        If bSupportHoleAspect = True Then
            ' Currently only equipment and MCT are known outfittings that support
            ' hole aspect and are supposed to use alone to create hole trace
            oRepresentOutfitting.Add oOutfittingObj
        Else
            oOutfittingsWithNoHoleAspect.Add oOutfittingObj
        End If
    Next
    
    If oOutfittingsWithNoHoleAspect.Count > 0 Then 'If there are any outfitting that do not support hole aspect process the code.
        ' The symbol may be dependent on the Fitting.  First determine if the
        ' HoleFitting has been created.
        Dim oComputedProjections As IJDObjectCollection
        Set oComputedProjections = New JObjectCollection
        
        Dim oDefinitionRule As New DefinitionRule
        Dim oHoleFitting As IJHoleFitting
        
        On Error Resume Next
        Set oHoleFitting = oHoleTrace.GetHoleFitting
        On Error GoTo ErrorHandler
    
        ' For a catalog fitting pipe sleeve special processing is required.
        If Not oHoleFitting Is Nothing Then
            ' Determine the type of fitting.
            Dim eFittingType As HMFittingType
            eFittingType = oHoleFitting.FittingType
        
            ' Attempt to compute the representative projections for the case
            ' of a pipe sleeve.  In the event that this fails, process the
            ' outfitting in the normal way.
            If eFittingType = HM_FittingCatalogPart Then
                bProjectionsComputed = oDefinitionRule.ComputePipeWithSleeveTrace _
                        (oHoleTrace, oOutfittingsWithNoHoleAspect, oComputedProjections)
            End If
        End If
    
        If Not bProjectionsComputed Then
            'Get the IJProjection which repersents the Outfitting
            TrackText "Get the IJProjection which repersents the Outfitting", 0
            bProjectionsComputed = oDefinitionRule.ComputeTrace _
                (oHoleTrace, oStructure, oOutfittingsWithNoHoleAspect, oWorkingPlane, oComputedProjections)
        End If
        oRepresentOutfitting.SetAdd oComputedProjections
        Set oDefinitionRule = Nothing
    End If
    
'******************************************************************************************************
'use the IJHMBestFitHelper to generate the hole trace geometry
'******************************************************************************************************
    TrackText "generate hole trace geometry", 0
    Dim oGeometry As IJComplexString
 

    Dim oObject As Object
    Set oObject = oHoleTrace.GetHoleTraceCmplxStr
    If Not oObject Is Nothing Then
        If TypeOf oObject Is IJComplexString Then
            Dim oGeometryFactory As GeometryFactory
            Set oGeometryFactory = New GeometryFactory
    
            Dim oElem As IJElements
            oObject.GetCurves oElem
    
            Dim oCmplxStrng As IJComplexString
            Set oCmplxStrng = oGeometryFactory.ComplexStrings3d.CreateByCurves(pOutputColl.ResourceManager, oElem)
    
            Set oGeometry = oCmplxStrng
        End If
    Else
        If oRepresentOutfitting.Count > 0 Then
            TrackText "create the BestFit helpers", 0
    
            Dim oBestFitHelper As IJHMBestFitHelper
            Set oBestFitHelper = New CHMBestFitHelper
    
            'Get the plate thickness
            TrackText "Get the plate thickness", 0
            Dim oPlateThickness As Double
            oPlateThickness = GetPlateThickness(oStructure)
    
            ' Generate the complex string that represents the hole trace.
           Set oGeometry = oBestFitHelper.TraceGeneration _
                  (pOutputColl.ResourceManager, _
                   oWorkingPlane, _
                   oPlateThickness, _
                   oRepresentOutfitting, _
                   bSupportHoleAspect)

            Set oBestFitHelper = Nothing
            Dim errNo As Long
            errNo = Err.Number
            If errNo = ERR_CMPLXSTRING Then
                On Error GoTo ErrorHandler
                Call RaiseTodoErrorMessage(oHoleTrace)
                Err.Raise errNo
            ElseIf errNo = ERR_NONCIRCUALRPROJECTION Then
                On Error GoTo ErrorHandler
                Call RaiseTodoErrorMessage(oHoleTrace, "CmnAppSemanticErrorMessages", 22)
                Err.Raise errNo
            End If
        End If
    End If
    On Error GoTo ErrorHandler
'******************************************************************************************************
'If intersection type is NoIntersectionWithTolerance, validate complex string within tolerance of
'possible opening/cut
'******************************************************************************************************
    Dim oRule As IJSRDRule
    
    oHoleTrace.GetHoleClearanceRule oRule

    If Not oRule Is Nothing Then
        Dim oRuleObject As Object
        Set oRuleObject = SP3DCreateObject(oRule.ProgId)
        If oRuleObject Is Nothing Then GoTo Cleanup
        
        Dim oHoleTraceIntersectionType As IJHMIntersection
        Dim eIntersectType As ISHoleTrace.HMHoleTraceIntersectionType
        Dim dTolerance As Double
        
        On Error Resume Next
        Set oHoleTraceIntersectionType = oRuleObject
        On Error GoTo ErrorHandler
        
        If Not oHoleTraceIntersectionType Is Nothing Then
            eIntersectType = oHoleTraceIntersectionType.IntersectionType
            If eIntersectType = HM_NoIntersectionWithTolerance Then
                dTolerance = oHoleTraceIntersectionType.Tolerance
                
                Dim oCurves As IJElements
                Dim varCurve As Variant
                Dim bToleranceFound As Boolean
                oGeometry.GetCurves oCurves
                
                bToleranceFound = False
                For Each varCurve In oCurves
                    ' Find closest distance between generated hole wire and
                    ' parent surface.  If distance if less than or equal to
                    ' tolerance, generated wire object is valid
                    Dim oIJGeomMisc As IMSModelGeomOps.IJGeometryMisc
                    Dim NullConnection As Object
                    Dim oModelBody As IJModelBody
                    Dim oAggr As Object
                    Dim oStructureModelBody As IJModelBody
                    
                    Set oIJGeomMisc = New IMSModelGeomOps.DGeomOpsMisc
                    
                    oIJGeomMisc.CreateModelGeometryFromGType NullConnection, varCurve, oAggr, oModelBody
                    Set oStructureModelBody = oStructure
                    Set oIJGeomMisc = Nothing
                    
                    Dim oModelBodyUtilities As GSCADShipGeomOps.SGOModelBodyUtilities
                    Dim oFirstModelPos As IJDPosition
                    Dim oSecondModelPos As IJDPosition
                    Dim dDistance As Double
                    
                    Set oModelBodyUtilities = New GSCADShipGeomOps.SGOModelBodyUtilities
                    oModelBodyUtilities.GetClosestPointsBetweenTwoBodies oStructureModelBody, _
                        oModelBody, oFirstModelPos, oSecondModelPos, dDistance
                        
                    If dDistance <= dTolerance Then
                        ' Closest distance between hole trace complex string and parent structure
                        ' is greater than tolerance; symbol should fail
                        bToleranceFound = True
                        Exit For
                    Else
                        oHoleTrace.state = HM_StateCritical
                        GoTo Cleanup
                    End If
                    
                    Set oFirstModelPos = Nothing
                    Set oSecondModelPos = Nothing
                    Set oStructureModelBody = Nothing
                    Set oModelBody = Nothing
                    Set oModelBodyUtilities = Nothing
                    dDistance = 0#
                Next
                
                If bToleranceFound = False Then
                    Set oGeometry = Nothing
                End If
            End If
        End If
    End If
'******************************************************************************************************
'set the output to the complex string built by the IJHMBestFitHelper
'******************************************************************************************************
    TrackText "set symbol output", 0

    If oGeometry Is Nothing Then GoTo Cleanup
 
    Dim oGeometryEx As IJComplexString
    Dim bProjected As Boolean
  
    bProjected = False
    Set oGeometryEx = oHMMiddleHelper.ProjectTraceOnPart(pOutputColl.ResourceManager, oStructure, oGeometry, oWorkingPlane, bProjected)
                     
    If (bProjected = True) Then
        Dim pObject As IJDObject
        Set pObject = oGeometry
        pObject.Remove
        Set pObject = Nothing
        Set oGeometry = oGeometryEx
     
    End If
    
     pOutputColl.AddOutput OUTPUT_1, oGeometry
      
'******************************************************************************************************
' With complex string established, the Diameter property of the HoleTraceAE can be set.
'******************************************************************************************************
    
    TrackText "set hole trace diameter", 0
    
     oHoleTrace.Diameter = oGeometry
    
'******************************************************************************************************
'set the hole trace state to good (0)
'******************************************************************************************************
    TrackText "set the hole trace state to good", 0

    oHoleTrace.state = HM_StateGood
    
'******************************************************************************************************
'clean up and return
'******************************************************************************************************
    TrackText "clean up and return", 0

Cleanup:

'    Set oVertices = Nothing
    Set oGeometry = Nothing
    Set oObject = Nothing
    Set oCmplxStrng = Nothing

    Set oHoleTrace = Nothing
    Set oStructure = Nothing
    Set oOutfitting = Nothing
    
    Set oRepresentOutfitting = Nothing
    Set oWorkingPlane = Nothing

    Set oBestFitHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: TrackText
'
'Interface: private sub
'
'Abstract:
'********************************************************************
Private Sub TrackText(strText As String, nIndent As Integer)
    On Error GoTo SkipTrack
    
    Dim oTracker As GSTracker.Tracker
    Set oTracker = New GSTracker.Tracker
    On Error GoTo KillTrack
    If nIndent < 0 Then oTracker.Outdent
    oTracker.UseFontDefaults
    oTracker.ForegroundColor = RGB(0, 255, 255)
    oTracker.WriteLn strText
    If nIndent > 0 Then oTracker.Indent

KillTrack:
    Set oTracker = Nothing

SkipTrack:

End Sub
Private Sub RaiseTodoErrorMessage(oHoleTrace As IJHoleTraceAE, Optional codelistTable As String = "HoleMgmtSemanticErrors", Optional CodeListNumber As Integer = 19)
    If Not oHoleTrace Is Nothing Then
        Dim oHolesmartOcc As IJHoleSmartOcc
        Dim oTodoHelper As IJToDoListHelper
        Set oHolesmartOcc = oHoleTrace.GetSmartOccurrence
        Set oTodoHelper = oHolesmartOcc
        oTodoHelper.SetErrorInfo codelistTable, CodeListNumber, oHolesmartOcc
        Set oHolesmartOcc = Nothing
        Set oTodoHelper = Nothing
    End If
End Sub

 
