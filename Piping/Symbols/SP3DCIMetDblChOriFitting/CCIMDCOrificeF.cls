VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCIMDCOrificeF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCIMDCOrificeF.cls.cls
'   ProgID:         SP3DCIMetDblChOriFitting.CCIMDCOrificeF
'   Author:         svsmylav
'   Creation Date:  Friday, Jan 10 2003
'   Description:
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMDC2,
'   PDMS Category Name=IFRC). The symbol consists of 12  outputs, nine physical,
'   two insulation outputs and one Operation output.
'   The 4 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   15.Sep.2003    SSP                             TR-49538 Craeted new class file COperation.cls to add operation aspect to the Symbol
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CCIMDCOrificeF:"  'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String
Private m_ProjName          As String
Private m_ClassName         As String

Private InputTypes()        As InputType
Private m_TextInputTypes()  As TextInputType
Private OutputTypes()       As OutputType
Private AspectTypes()       As AspectType
Private nAspects            As Integer
Private nInputs             As Integer
Private iTextCount         As Integer
Private nOutputs            As Integer
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProjName = "SP3DCIMetDblChOriFitting"
    m_ClassName = "CCIMDCOrificeF"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 51
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).Name = "InstrumentLength"
    InputTypes(iCount).Description = "Instrument Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 2
    InputTypes(iCount).Name = "InstrumentLength1"
    InputTypes(iCount).Description = "Instrument Length1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 3
    InputTypes(iCount).Name = "InstrumentHeight"
    InputTypes(iCount).Description = "Instrument Height"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.08
    
    iCount = 4
    InputTypes(iCount).Name = "InstrumentHeight1"
    InputTypes(iCount).Description = "Instrument Height1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 5
    InputTypes(iCount).Name = "InstrumentLength2"
    InputTypes(iCount).Description = "Instrument Length2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05
    
    iCount = 6
    InputTypes(iCount).Name = "InstrumentLength3"
    InputTypes(iCount).Description = "Instrument Length3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05
    
    iCount = 7
    InputTypes(iCount).Name = "InstrumentWidth"
    InputTypes(iCount).Description = "Instrument Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05
    
    iCount = 8
    InputTypes(iCount).Name = "InstrumentWidth1"
    InputTypes(iCount).Description = "Instrument Width1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05
    
    iCount = 9
    InputTypes(iCount).Name = "ArmLength"
    InputTypes(iCount).Description = "Arm Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 10
    InputTypes(iCount).Name = "NozzleOffset"
    InputTypes(iCount).Description = "Nozzle Offset"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025
    
    iCount = 11
    InputTypes(iCount).Name = "NozzleOffset1"
    InputTypes(iCount).Description = "Nozzle Offset1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025
    
    iCount = 12
    InputTypes(iCount).Name = "ArmHeight"
    InputTypes(iCount).Description = "Arm Height"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.15
    
    iCount = 13
    InputTypes(iCount).Name = "ArmHeight1"
    InputTypes(iCount).Description = "Arm Height1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.13
    
    iCount = 14
    InputTypes(iCount).Name = "InstrumentLength4"
    InputTypes(iCount).Description = "Instrument Length4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.3
    
    iCount = 15
    InputTypes(iCount).Name = "InstrumentLength5"
    InputTypes(iCount).Description = "Instrument Length5"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.15
    
    iCount = 16
    InputTypes(iCount).Name = "InsulationThickness"
    InputTypes(iCount).Description = "Insulation Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025
    
    iCount = 17
    InputTypes(iCount).Name = "Npd"
    InputTypes(iCount).Description = "NPD"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 100
        
    iCount = 18
    InputTypes(iCount).Name = "EndPreparation"
    InputTypes(iCount).Description = "EndPreparation"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 21
    
    iCount = 19
    InputTypes(iCount).Name = "ScheduleThickness"
    InputTypes(iCount).Description = "ScheduleThickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 100
    
    iCount = 20
    InputTypes(iCount).Name = "EndStandard"
    InputTypes(iCount).Description = "EndStandard"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 5
    
    iCount = 21
    InputTypes(iCount).Name = "PressureRating"
    InputTypes(iCount).Description = "PressureRating"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 35
    
    iCount = 22
    InputTypes(iCount).Name = "FlowDirection"
    InputTypes(iCount).Description = "FlowDirection"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 3
            
    iCount = 23
    InputTypes(iCount).Name = "PortIndex1"
    InputTypes(iCount).Description = "PortIndex1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1

    iCount = 24
    InputTypes(iCount).Name = "Npd1"
    InputTypes(iCount).Description = "NPD1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
        
    iCount = 25
    InputTypes(iCount).Name = "EndPreparation1"
    InputTypes(iCount).Description = "EndPreparation1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 26
    InputTypes(iCount).Name = "ScheduleThickness1"
    InputTypes(iCount).Description = "ScheduleThickness1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 27
    InputTypes(iCount).Name = "EndStandard1"
    InputTypes(iCount).Description = "EndStandard1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 28
    InputTypes(iCount).Name = "PressureRating1"
    InputTypes(iCount).Description = "PressureRating1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 29
    InputTypes(iCount).Name = "FlowDirection1"
    InputTypes(iCount).Description = "FlowDirection1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 30
    InputTypes(iCount).Name = "PortIndex2"
    InputTypes(iCount).Description = "PortIndex2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2

    iCount = 31
    InputTypes(iCount).Name = "Npd2"
    InputTypes(iCount).Description = "NPD2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
        
    iCount = 32
    InputTypes(iCount).Name = "EndPreparation2"
    InputTypes(iCount).Description = "EndPreparation2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 33
    InputTypes(iCount).Name = "ScheduleThickness2"
    InputTypes(iCount).Description = "ScheduleThickness2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 34
    InputTypes(iCount).Name = "EndStandard2"
    InputTypes(iCount).Description = "EndStandard2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 35
    InputTypes(iCount).Name = "PressureRating2"
    InputTypes(iCount).Description = "PressureRating2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 36
    InputTypes(iCount).Name = "FlowDirection2"
    InputTypes(iCount).Description = "FlowDirection2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0

   iCount = 37
    InputTypes(iCount).Name = "PortIndex3"
    InputTypes(iCount).Description = "PortIndex3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 3

    iCount = 38
    InputTypes(iCount).Name = "Npd3"
    InputTypes(iCount).Description = "NPD3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
        
    iCount = 39
    InputTypes(iCount).Name = "EndPreparation3"
    InputTypes(iCount).Description = "EndPreparation3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 40
    InputTypes(iCount).Name = "ScheduleThickness3"
    InputTypes(iCount).Description = "ScheduleThickness3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 41
    InputTypes(iCount).Name = "EndStandard3"
    InputTypes(iCount).Description = "EndStandard3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 42
    InputTypes(iCount).Name = "PressureRating3"
    InputTypes(iCount).Description = "PressureRating3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 43
    InputTypes(iCount).Name = "FlowDirection3"
    InputTypes(iCount).Description = "FlowDirection3"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 44
    InputTypes(iCount).Name = "PortIndex4"
    InputTypes(iCount).Description = "PortIndex4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 4

    iCount = 45
    InputTypes(iCount).Name = "Npd4"
    InputTypes(iCount).Description = "NPD4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
        
    iCount = 46
    InputTypes(iCount).Name = "EndPreparation4"
    InputTypes(iCount).Description = "EndPreparation4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 47
    InputTypes(iCount).Name = "ScheduleThickness4"
    InputTypes(iCount).Description = "ScheduleThickness4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 48
    InputTypes(iCount).Name = "EndStandard4"
    InputTypes(iCount).Description = "EndStandard4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 49
    InputTypes(iCount).Name = "PressureRating4"
    InputTypes(iCount).Description = "PressureRating4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 50
    InputTypes(iCount).Name = "FlowDirection4"
    InputTypes(iCount).Description = "FlowDirection4"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 51
    InputTypes(iCount).Name = "OperationHeight"
    InputTypes(iCount).Description = "Operation Height"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.5

    'Text Inputs
    iTextCount = 9
    ReDim m_TextInputTypes(1 To iTextCount) As TextInputType

    iCount = 1
    m_TextInputTypes(iCount).Name = "Id1"
    m_TextInputTypes(iCount).Description = "Id1"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "VNoz1"

    iCount = 2
    m_TextInputTypes(iCount).Name = "Id2"
    m_TextInputTypes(iCount).Description = "Id2"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "VNoz2"

    iCount = 3
    m_TextInputTypes(iCount).Name = "Id3"
    m_TextInputTypes(iCount).Description = "Id3"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "VNoz3"

    iCount = 4
    m_TextInputTypes(iCount).Name = "Id4"
    m_TextInputTypes(iCount).Description = "Id4"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "VNoz4"

    iCount = 5
    m_TextInputTypes(iCount).Name = "NpdUnitType"
    m_TextInputTypes(iCount).Description = "Npd Unit Type"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "mm"

    iCount = 6
    m_TextInputTypes(iCount).Name = "NpdUnitType1"
    m_TextInputTypes(iCount).Description = "Npd Unit Type 1"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""

    iCount = 7
    m_TextInputTypes(iCount).Name = "NpdUnitType2"
    m_TextInputTypes(iCount).Description = "Npd Unit Type 2"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""

    iCount = 8
    m_TextInputTypes(iCount).Name = "NpdUnitType3"
    m_TextInputTypes(iCount).Description = "Npd Unit Type 3"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""

    iCount = 9
    m_TextInputTypes(iCount).Name = "NpdUnitType4"
    m_TextInputTypes(iCount).Description = "Npd Unit Type 4"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""
    
        
' Outputs
    nOutputs = 12
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).Name = "VNoz1"
    OutputTypes(iCount).Description = "Nozzle 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).Name = "VNoz2"
    OutputTypes(iCount).Description = "Nozzle 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 3
    OutputTypes(iCount).Name = "VNoz3"
    OutputTypes(iCount).Description = "Nozzle 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).Name = "VNoz4"
    OutputTypes(iCount).Description = "Nozzle 4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 5
    OutputTypes(iCount).Name = "LHSPipe"
    OutputTypes(iCount).Description = "Left Hand Side Pipe"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).Name = "RHSPipe"
    OutputTypes(iCount).Description = "Right Hand Side Pipe"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 7
    OutputTypes(iCount).Name = "FittingBody"
    OutputTypes(iCount).Description = "Body Orifice fitting "
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 8
    OutputTypes(iCount).Name = "ArmCylinder1"
    OutputTypes(iCount).Description = "ArmCylinder1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 9
    OutputTypes(iCount).Name = "ArmBody2"
    OutputTypes(iCount).Description = "ArmBody2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 10
    OutputTypes(iCount).Name = "InsBody"
    OutputTypes(iCount).Description = "Insulation for Body of Orifice fitting"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 11
    OutputTypes(iCount).Name = "InsPipeLine"
    OutputTypes(iCount).Description = "Insulation for pipes"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 12
    OutputTypes(iCount).Name = "OpEnvelope"
    OutputTypes(iCount).Description = "Envelop for Operation "
    OutputTypes(iCount).Aspect = Operation
    
' Aspects
    nAspects = 3
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    iCount = 1
    AspectTypes(iCount).Name = "Physical"
    AspectTypes(iCount).Description = "Physical"
    AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = SimplePhysical
    iCount = 2
    AspectTypes(iCount).Name = "Insulation"
    AspectTypes(iCount).Description = "Insulation"
    AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = Insulation
    iCount = 3
    AspectTypes(iCount).Name = "Operation"
    AspectTypes(iCount).Description = "Operation"
    AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = Operation
    
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
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
   
' returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
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
        Inputs(iCount).Name = InputTypes(iCount).Name
        Inputs(iCount).Description = InputTypes(iCount).Description
        Inputs(iCount).Properties = InputTypes(iCount).Properties
        PC.uomValue = InputTypes(iCount).uomValue
' Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount
        
'----------Text Parameters ----------------
    'Create a default parameter
'uncomment to make test with text inputs
    Dim PC1 As IMSSymbolEntities.IJDParameterContent
    Set PC1 = New IMSSymbolEntities.DParameterContent
    PC1.Type = igString

    ReDim TextInputs(1 To iTextCount) As IMSSymbolEntities.IJDInput
    For iCount = 1 To iTextCount
        Set TextInputs(iCount) = New IMSSymbolEntities.DInput
        TextInputs(iCount).Name = m_TextInputTypes(iCount).Name
        TextInputs(iCount).Description = m_TextInputTypes(iCount).Description
        TextInputs(iCount).Properties = m_TextInputTypes(iCount).Properties
        PC1.String = m_TextInputTypes(iCount).Value
        TextInputs(iCount).DefaultParameterValue = PC1
        InputsIf.SetInput TextInputs(iCount), nInputs + iCount + 1
        Set TextInputs(iCount) = Nothing
    Next iCount
'------------------------------------------

' Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    Set symbAspect = New IMSSymbolEntities.DRepresentation
' ---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For iCount = 1 To nAspects
        symbAspect.Name = AspectTypes(iCount).Name
        symbAspect.Description = AspectTypes(iCount).Description
        symbAspect.Properties = AspectTypes(iCount).Properties
        symbAspect.RepresentationId = AspectTypes(iCount).AspectId
'        symbAspect.RemoveAllOutput
        
        Set oSymbAspect = symbAspect
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).Aspect And AspectTypes(iCount).AspectId Then
                Output.Name = OutputTypes(iOutput).Name
                Output.Description = OutputTypes(iOutput).Description
                Output.Properties = 0            ' OutputTypes(iOutput).Properties
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
    
    For iCount = 1 To nAspects
        Set oRepEval = New DRepresentationEvaluation
        oRepEval.Name = AspectTypes(iCount).Name
        oRepEval.Description = AspectTypes(iCount).Description
        oRepEval.Properties = igREPRESENTATION_HIDDEN
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
    Dim lAspectID           As SymbolRepIds
    Dim AspectName          As String
    Dim bsucceded           As Boolean
    
    Set m_OutputColl = outputcoll
    
    bsucceded = False
    For iCount = 1 To nAspects
        If AspectTypes(iCount).Name = repName Then
            bsucceded = True
            lAspectID = AspectTypes(iCount).AspectId
            Exit For
        End If
    Next iCount
    If Not bsucceded Then
'        MsgBox "Unknown Aspect; Check Aspect data / code"
        Exit Sub
    End If
    
    iOutputs = 1
    For iCount = 1 To nOutputs
        If OutputTypes(iCount).Aspect = lAspectID Then
            ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
            arrayOfOutputs(iOutputs) = OutputTypes(iCount).Name
            iOutputs = iOutputs + 1
        End If
    Next iCount
    Dim objAspect As Object
    If StrComp(repName,"Physical", vbTextCompare) = 0 Then
        Set objAspect = New CPhysical
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName,"Insulation", vbTextCompare) = 0 Then
        Set objAspect = New CInsulation
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName,"Operation", vbTextCompare) = 0 Then
        Set objAspect = New COperation
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
