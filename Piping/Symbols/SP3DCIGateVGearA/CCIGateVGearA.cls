VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCIGateVGearA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-09, Intergraph Corporation. All rights reserved.
'
'   CCIGateVGearA.cls
'   ProgID:        SP3DCIGateVGearA.CCIGateVGearA
'   Author:         NN
'   Creation Date:  Sunday, Dec 15 2002
'   Description:    Custom Instrument control gate valve with gear actuator
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   Dec 13, 2002    NN   Based on GSCADGateValveAngleOp.CGateValve150; created dynamic ports,
'                                   all port properties are Occurence Attributes.
'   April 20, 2003  NN   Define NPD units, Nozzle IDs as "text" inputs.
'   May 3, 2003     NN   define upper case W for parameter HandWheelDiameter
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  21.Feb.2007      RRK                     TR-113129 Changes made to include Insulation Aspect
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   21.APR.2009     PK     TR-163529    CreateObject() has been replaced with 'New' representation
'                                        class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CCIGateVGearA:"  'Used for error messages

Private m_oSymbolFactory    As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_sProgId           As String
Private m_sProjName         As String
Private m_sClassName        As String

Private m_InputTypes()      As InputType
Private m_OutputTypes()     As OutputType
Private m_AspectTypes()     As AspectType
Private m_iAspects          As Integer
Private m_iInputs           As Integer
Private m_iOutputs          As Integer
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_sProjName = "SP3DCIGateVGearA"
    m_sClassName = "CCIGateVGearA"
    m_sProgId = m_sProjName & "." & m_sClassName

    m_iInputs = 30
    ReDim m_InputTypes(1 To m_iInputs) As InputType
    
    iCount = 1
    m_InputTypes(iCount).Name = "FacetoFace"
    m_InputTypes(iCount).Description = "valve width"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0.25
    
    iCount = 2
    m_InputTypes(iCount).Name = "ValveHeight"
    m_InputTypes(iCount).Description = "valve height"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0.45
    
    iCount = 3
    m_InputTypes(iCount).Name = "HandWheelDiameter"
    m_InputTypes(iCount).Description = "handwheel diameter"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0.2
    
    iCount = 4
    m_InputTypes(iCount).Name = "HandwheelAngle"
    m_InputTypes(iCount).Description = "handwheel angle"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 1.570796
    
    iCount = 5
    m_InputTypes(iCount).Name = "Npd"
    m_InputTypes(iCount).Description = "NPD"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 100
        
    iCount = 6
    m_InputTypes(iCount).Name = "EndPreparation"
    m_InputTypes(iCount).Description = "EndPreparation"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 21
    
    iCount = 7
    m_InputTypes(iCount).Name = "ScheduleThickness"
    m_InputTypes(iCount).Description = "ScheduleThickness"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 100
    
    iCount = 8
    m_InputTypes(iCount).Name = "EndStandard"
    m_InputTypes(iCount).Description = "EndStandard"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 5
    
    iCount = 9
    m_InputTypes(iCount).Name = "PressureRating"
    m_InputTypes(iCount).Description = "PressureRating"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 35
    
    iCount = 10
    m_InputTypes(iCount).Name = "FlowDirection"
    m_InputTypes(iCount).Description = "FlowDirection"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 3
            
    iCount = 11
    m_InputTypes(iCount).Name = "PortIndex1"
    m_InputTypes(iCount).Description = "PortIndex1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 1

    iCount = 12
    m_InputTypes(iCount).Name = "Npd1"
    m_InputTypes(iCount).Description = "NPD1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
        
    iCount = 13
    m_InputTypes(iCount).Name = "EndPreparation1"
    m_InputTypes(iCount).Description = "EndPreparation1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 14
    m_InputTypes(iCount).Name = "ScheduleThickness1"
    m_InputTypes(iCount).Description = "ScheduleThickness1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 15
    m_InputTypes(iCount).Name = "EndStandard1"
    m_InputTypes(iCount).Description = "EndStandard1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 16
    m_InputTypes(iCount).Name = "PressureRating1"
    m_InputTypes(iCount).Description = "PressureRating1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 17
    m_InputTypes(iCount).Name = "FlowDirection1"
    m_InputTypes(iCount).Description = "FlowDirection1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 3
    
    iCount = 18
    m_InputTypes(iCount).Name = "PortIndex2"
    m_InputTypes(iCount).Description = "PortIndex2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 2

    iCount = 19
    m_InputTypes(iCount).Name = "Npd2"
    m_InputTypes(iCount).Description = "NPD2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
        
    iCount = 20
    m_InputTypes(iCount).Name = "EndPreparation2"
    m_InputTypes(iCount).Description = "EndPreparation2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 21
    m_InputTypes(iCount).Name = "ScheduleThickness2"
    m_InputTypes(iCount).Description = "ScheduleThickness2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 22
    m_InputTypes(iCount).Name = "EndStandard2"
    m_InputTypes(iCount).Description = "EndStandard2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 23
    m_InputTypes(iCount).Name = "PressureRating2"
    m_InputTypes(iCount).Description = "PressureRating2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0
    
    iCount = 24
    m_InputTypes(iCount).Name = "FlowDirection2"
    m_InputTypes(iCount).Description = "FlowDirection2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = 0

    iCount = 25
    m_InputTypes(iCount).Name = "Id1"
    m_InputTypes(iCount).Description = "Id1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = "VNoz1"

    iCount = 26
    m_InputTypes(iCount).Name = "Id2"
    m_InputTypes(iCount).Description = "Id2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = "VNoz2"

    iCount = 27
    m_InputTypes(iCount).Name = "NpdUnitType"
    m_InputTypes(iCount).Description = "Npd Unit Type"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = "mm"

    iCount = 28
    m_InputTypes(iCount).Name = "NpdUnitType1"
    m_InputTypes(iCount).Description = "Npd Unit Type 1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = ""

    iCount = 29
    m_InputTypes(iCount).Name = "NpdUnitType2"
    m_InputTypes(iCount).Description = "Npd Unit Type 2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).UomValue = ""
    
    iCount = 30
    m_InputTypes(iCount).Name = "InsulationThickness"
    m_InputTypes(iCount).Description = "Insulation Thickness"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    m_InputTypes(iCount).UomValue = 0.025
   
' Outputs
    m_iOutputs = 14
    ReDim m_OutputTypes(1 To m_iOutputs) As OutputType

    iCount = 1
    m_OutputTypes(iCount).Name = "ValveBody"
    m_OutputTypes(iCount).Description = "Main body"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "ValvePost"
    m_OutputTypes(iCount).Description = "Vertical post"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "Valvestem"
    m_OutputTypes(iCount).Description = "Stem"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "Support brace"
    m_OutputTypes(iCount).Description = "Brace"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "Crossbrace"
    m_OutputTypes(iCount).Description = "Cross brace"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "TransferCase"
    m_OutputTypes(iCount).Description = "Transfer case"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "SpindleCone"
    m_OutputTypes(iCount).Description = "Spindle cone"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "HandWheelSpoke1"
    m_OutputTypes(iCount).Description = "Spoke 1"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "HandWheelSpoke2"
    m_OutputTypes(iCount).Description = "Spoke 2"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "HandWheel"
    m_OutputTypes(iCount).Description = "Hand wheel"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "Knob"
    m_OutputTypes(iCount).Description = "Hand wheel knob"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "VNoz1"
    m_OutputTypes(iCount).Description = "Nozzle 1"
    m_OutputTypes(iCount).Aspect = SimplePhysical

    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "VNoz2"
    m_OutputTypes(iCount).Description = "Nozzle 2"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = iCount + 1
    m_OutputTypes(iCount).Name = "InsulatedValve"
    m_OutputTypes(iCount).Description = "Insulated Valve"
    m_OutputTypes(iCount).Aspect = Insulation
    
    
' Aspects
    m_iAspects = 2
    ReDim m_AspectTypes(1 To m_iAspects) As AspectType
    
    iCount = 1
    m_AspectTypes(iCount).Name = "Physical"
    m_AspectTypes(iCount).Description = "Physical"
    m_AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    m_AspectTypes(iCount).AspectId = SimplePhysical
     
    iCount = 2
    m_AspectTypes(iCount).Name = "Insulation"
    m_AspectTypes(iCount).Description = "Insulation"
    m_AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    m_AspectTypes(iCount).AspectId = Insulation

    
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub Class_Terminate()
    
    Set m_oSymbolFactory = Nothing

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
    Set m_oSymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set oSymbolDefinition = m_oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set m_oSymbolFactory = Nothing
' Initialize it
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_sProgId
    oSymbolDefinition.CodeBase = CodeBase

' Give a unique name to the symbol definition
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
   
' returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
    Debug.Assert False

End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
  IJDUserSymbolServices_GetDefinitionName = m_sProgId

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
    ReDim Inputs(1 To m_iInputs) As IMSSymbolEntities.IJDInput
    Dim iCount      As Integer
' Create a default parameter
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    For iCount = 1 To m_iInputs
        If VarType(m_InputTypes(iCount).UomValue) = 8 Then
            PC.Type = igString
            PC.String = m_InputTypes(iCount).UomValue
        Else
            PC.Type = igValue
            PC.UomValue = m_InputTypes(iCount).UomValue
        End If
        Set Inputs(iCount) = New IMSSymbolEntities.DInput
        Inputs(iCount).Name = m_InputTypes(iCount).Name
        Inputs(iCount).Description = m_InputTypes(iCount).Description
        Inputs(iCount).Properties = m_InputTypes(iCount).Properties
' Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount
        
' Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    Set symbAspect = New IMSSymbolEntities.DRepresentation
' ---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For iCount = 1 To m_iAspects
        symbAspect.Name = m_AspectTypes(iCount).Name
        symbAspect.Description = m_AspectTypes(iCount).Description
        symbAspect.Properties = m_AspectTypes(iCount).Properties
        symbAspect.RepresentationId = m_AspectTypes(iCount).AspectId
'        symbAspect.RemoveAllOutput
        
        Set oSymbAspect = symbAspect
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To m_iOutputs
            If m_OutputTypes(iOutput).Aspect And m_AspectTypes(iCount).AspectId Then
                Output.Name = m_OutputTypes(iOutput).Name
                Output.Description = m_OutputTypes(iOutput).Description
                Output.Properties = 0            ' m_OutputTypes(iOutput).Properties
                oSymbAspect.SetOutput Output
            End If
        Next iOutput
        RepsIf.SetRepresentation symbAspect
        Set Output = Nothing
    Next iCount
    

    ' ----------Definition of the function identification corresponding
    Dim oRepEval     As IJDRepresentationEvaluation
    Dim RepsEvalsIf     As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition
    
    For iCount = 1 To m_iAspects
        Set oRepEval = New DRepresentationEvaluation
        oRepEval.Name = m_AspectTypes(iCount).Name
        oRepEval.Description = m_AspectTypes(iCount).Description
        oRepEval.Properties = igREPRESENTATION_HIDDEN
        oRepEval.Type = igREPRESENTATION_VBFUNCTION
        oRepEval.ProgId = m_sProgId
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

Public Sub IJDUserSymbolServices_InvokeRepresentation( _
                    ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    '--------------------------
    'MsgBox METHOD
    '--------------------------
    
    Dim arrayOfOutputs()    As String
    Dim iOutputs            As Integer
    Dim iCount              As Integer
    Dim lAspectID           As SymbolRepIds
    Dim AspectName          As String
    Dim bsucceded           As Boolean
    
    Set m_OutputColl = outputcoll
    
    bsucceded = False
    For iCount = 1 To m_iAspects
        If m_AspectTypes(iCount).Name = repName Then
            bsucceded = True
            lAspectID = m_AspectTypes(iCount).AspectId
            Exit For
        End If
    Next iCount
    If Not bsucceded Then
'        MsgBox "Unknown Aspect; Check Aspect data / code"
        Exit Sub
    End If
    
    iOutputs = 1
    For iCount = 1 To m_iOutputs
        If m_OutputTypes(iCount).Aspect = lAspectID Then
            ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
            arrayOfOutputs(iOutputs) = m_OutputTypes(iCount).Name
            iOutputs = iOutputs + 1
        End If
    Next iCount
    Dim objAspect As Object
    
    If StrComp(repName, "Physical", vbTextCompare) = 0 Then
        Set objAspect = New CPhysical
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName, "Insulation", vbTextCompare) = 0 Then
        Set objAspect = New CInsulation
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If

        
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function
