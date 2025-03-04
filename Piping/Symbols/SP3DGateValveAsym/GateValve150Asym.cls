VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGateValve150Asym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'   Copyright (c) 2003-09, Intergraph Corporation. All rights reserved.
'
'   GateValve150 - Solid wedge Gate Valve class
'   06/29/99
'   LK
'   last revised 07/14/99
'   7/14/99 - Reusing the same b-spline curve to create different 3d objects
'   revision 20010201 NN constructed an asymmetric Drain Vent / Instrument Root Gate Valve SW X FT 800 # with graphics similar to the Class 300.
'   ProgId: SP3DGateValveAsym.CGateValve150Asym
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'   21.APR.2009     PK     TR-163529    CreateObject() has been replaced with 'New' representation
'                                        class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "GateValve150Asym" 'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String

Private InputTypes()            As InputType
Private OutputTypes()           As OutputType
Private RepresentationTypes()   As RepresentationType
Private nRepresentations        As Integer
Private nInputs                 As Integer
Private nOutputs                As Integer
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices


Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProgId = "SP3DGateValveAsym.CGateValve150Asym"

' Inputs
    nInputs = 4
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).name = "Face1toCenter"
    InputTypes(iCount).description = "Face1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.06
    
    iCount = 1 + iCount
    InputTypes(iCount).name = "Face2toCenter"
    InputTypes(iCount).description = "Face2 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.08
    
    iCount = 1 + iCount
    InputTypes(iCount).name = "OperatorHeight"
    InputTypes(iCount).description = "Height of Operator"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.614
        
    iCount = 1 + iCount
    InputTypes(iCount).name = "OperatorDiameter"
    InputTypes(iCount).description = "Diameter of Operator Wheel"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.25
        
' Outputs
    nOutputs = 12
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "BodyH"
    OutputTypes(iCount).description = "Horisontal Body of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "BodyVB"
    OutputTypes(iCount).description = "Vertical Body of GateValve; Bottom Part"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "BodyVT"
    OutputTypes(iCount).description = "Vertical Body of GateValve"
    OutputTypes(iCount).representation = SimplePhysical

    iCount = 1 + iCount
    OutputTypes(iCount).name = "BodyVFlange"
    OutputTypes(iCount).description = "Vertical Body Flange of GateValve"
    OutputTypes(iCount).representation = SimplePhysical

    iCount = 1 + iCount
    OutputTypes(iCount).name = "BonnetFlange"
    OutputTypes(iCount).description = "Bonnet Flange of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "Bonnet"
    OutputTypes(iCount).description = "Bonnet of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "Stem"
    OutputTypes(iCount).description = "Stem of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "Actuator"
    OutputTypes(iCount).description = "Actuator of GateValve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "ActuatorSpike1"
    OutputTypes(iCount).description = "Spike 1 of Actuator wheel"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    OutputTypes(iCount).name = "ActuatorSpike2"
    OutputTypes(iCount).description = "Spike 2 of Actuator wheel"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    'OutputTypes(iCount).name = "PipePort1"
    OutputTypes(iCount).name = "VNoz1"
    OutputTypes(iCount).description = "PipingPort1 of Gate valve"
    OutputTypes(iCount).representation = SimplePhysical
    
    iCount = 1 + iCount
    'OutputTypes(iCount).name = "PipePort2"
    OutputTypes(iCount).name = "VNoz2"
    OutputTypes(iCount).description = "PipingPort2 of Gate valve"
    OutputTypes(iCount).representation = SimplePhysical
    
' Representations
    nRepresentations = 1
    ReDim RepresentationTypes(1 To nRepresentations) As RepresentationType
    
    iCount = 1
    RepresentationTypes(iCount).name = "Physical"
    RepresentationTypes(iCount).description = "Physical representation"
    RepresentationTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    RepresentationTypes(iCount).representationId = SimplePhysical
    
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub Class_Terminate()
    
    Set m_SymbolFactory = Nothing

End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    
' This method is in charge of the creation of the symbol definition object
' You can keep the current design unchanged
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler
    
    Dim oSymbolDefinition  As IMSSymbolEntities.IJDSymbolDefinition
    
' Create the definition from the symbol factory
    Set m_SymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set oSymbolDefinition = m_SymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set m_SymbolFactory = Nothing
' Initialize it
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_ProgId
    oSymbolDefinition.CodeBase = CodeBase

' Give a unique name to the symbol definition
    oSymbolDefinition.name = oSymbolDefinition.ProgId
   
' returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
  IJDUserSymbolServices_GetDefinitionName = m_ProgId

End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler
   
    Dim oPart   As IMSSymbolEntities.DInput
    
' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
   Dim oSymbolCache As New CustomCache
   oSymbolCache.SetupCustomCache pSymbolDefinition
    
' ----------Parameter Input --------------
    Dim Inputs()    As IMSSymbolEntities.IJDInput
    ReDim Inputs(1 To nInputs) As IMSSymbolEntities.IJDInput
    Dim iCount      As Integer
' Create a default parameter
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    PC.Type = igValue
    For iCount = 1 To nInputs
        Set Inputs(iCount) = New IMSSymbolEntities.DInput
        Inputs(iCount).name = InputTypes(iCount).name
        Inputs(iCount).description = InputTypes(iCount).description
        Inputs(iCount).properties = InputTypes(iCount).properties
        PC.uomValue = InputTypes(iCount).uomValue
' Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount
        
' Define output
    Dim symbRepresentation As IMSSymbolEntities.IJDRepresentation
    Set symbRepresentation = New IMSSymbolEntities.DRepresentation
' ---------Set the representation to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For iCount = 1 To nRepresentations
        symbRepresentation.name = RepresentationTypes(iCount).name
        symbRepresentation.description = RepresentationTypes(iCount).description
        symbRepresentation.properties = RepresentationTypes(iCount).properties
        symbRepresentation.representationId = RepresentationTypes(iCount).representationId
        'symbRepresentation.RemoveAllOutput
Set oSymbAspect = symbRepresentation
        
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).representation And RepresentationTypes(iCount).representationId Then
                Output.name = OutputTypes(iOutput).name
                Output.description = OutputTypes(iOutput).description
                Output.properties = 0            ' OutputTypes(iOutput).Properties
                oSymbAspect.SetOutput Output
            End If
        Next iOutput
        RepsIf.SetRepresentation symbRepresentation
        Set Output = Nothing
    Next iCount
    

    ' ----------Definition of the function identification corresponding
    Dim oRepEval     As IJDRepresentationEvaluation
    Dim RepsEvalsIf     As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition
    
    For iCount = 1 To nRepresentations
        Set oRepEval = New DRepresentationEvaluation
        oRepEval.name = RepresentationTypes(iCount).name
        oRepEval.description = RepresentationTypes(iCount).description
        oRepEval.properties = igREPRESENTATION_HIDDEN
        oRepEval.Type = igREPRESENTATION_VBFUNCTION
        oRepEval.ProgId = m_ProgId
        RepsEvalsIf.AddRepresentationEvaluation oRepEval
        Set oRepEval = Nothing
    Next iCount
    Set RepsEvalsIf = Nothing
    
'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
        
  Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    Dim arrayOfOutputs()    As String
    Dim iOutputs            As Integer
    Dim iCount              As Integer
    Dim lReresentationID    As SymbolRepIds
    Dim representationName  As String
    Dim bsucceded           As Boolean
    
    Set m_OutputColl = outputcoll
    
    bsucceded = False
    For iCount = 1 To nRepresentations
        If RepresentationTypes(iCount).name = repName Then
            bsucceded = True
            lReresentationID = RepresentationTypes(iCount).representationId
            Exit For
        End If
    Next iCount
    If Not bsucceded Then
'        MsgBox "Unknown representation; Check representation data / code"
        Exit Sub
    End If
    
    iOutputs = 1
    For iCount = 1 To nOutputs
        If OutputTypes(iCount).representation = lReresentationID Then
            ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
            arrayOfOutputs(iOutputs) = OutputTypes(iCount).name
            iOutputs = iOutputs + 1
        End If
    Next iCount
    Dim objRepresentation As Object
    
    If StrComp(repName, "Physical", vbTextCompare) = 0 Then
        Set objRepresentation = New CPhysical
        objRepresentation.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objRepresentation = Nothing
    End If
            
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function

