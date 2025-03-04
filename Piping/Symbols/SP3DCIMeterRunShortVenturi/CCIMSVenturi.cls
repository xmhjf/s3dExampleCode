VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCIMSVenturi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-09, Intergraph Corporation. All rights reserved.
'
'   CCIMSVenturi.cls
'   ProgID:         SP3DCIMeterRunShortVenturi.CCIMSVenturi
'   Author:         SSP
'   Creation Date:  Wed,Dec 18 2002
'   Description:
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMRSV, PDMS Category Name=IFRD)
'   The symbol consists of eleven outputs, eight physical and three insulation outputs.
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   21-03-2003      SSP For working with both mm and in units, changed the default values of the ports 1..n to Null for string types and 0 for double types.
'   10-03-2003      SSP  Changed input parameters Angle,Angle1 to
'                           PortRotation1,PortRotation2 respectively.
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007      PS      CR-132021  Modified the default value of PortRotation1,PortRotation2
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   21.APR.2009     PK     TR-163529    CreateObject() has been replaced with 'New' representation
'                                        class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CCIMSVenturi:"  'Used for error messages

Private m_oSymbolFactory    As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_sProgId           As String
Private m_sProjName         As String
Private m_sClassName        As String

Private m_InputTypes()      As InputType
Private m_TextInputTypes()  As TextInputType
Private m_OutputTypes()     As OutputType
Private m_AspectTypes()     As AspectType
Private m_iAspects          As Integer
Private m_iInputs           As Integer
Private m_iTextInputs       As Integer
Private m_iOutputs          As Integer

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_sProjName = "SP3DCIMeterRunShortVenturi"
    m_sClassName = "CCIMSVenturi"
    m_sProgId = m_sProjName & "." & m_sClassName

' Inputs
    m_iInputs = 48
    ReDim m_InputTypes(1 To m_iInputs) As InputType
    
    iCount = 1
    m_InputTypes(iCount).Name = "InstrumentLength"
    m_InputTypes(iCount).Description = "Instrument Length"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.5
    
    iCount = 2
    m_InputTypes(iCount).Name = "InstrumentLength1"
    m_InputTypes(iCount).Description = "Instrument Length1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.4
    
    iCount = 3
    m_InputTypes(iCount).Name = "InstrumentLength2"
    m_InputTypes(iCount).Description = "Instrument Length2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.25
    
    iCount = 4
    m_InputTypes(iCount).Name = "InstrumentLength3"
    m_InputTypes(iCount).Description = "Instrument Length3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.25
    
    iCount = 5
    m_InputTypes(iCount).Name = "InstrumentDiameter"
    m_InputTypes(iCount).Description = "Instrument Diameter"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.25
    
    iCount = 6
    m_InputTypes(iCount).Name = "NozzleOffset"
    m_InputTypes(iCount).Description = "NozzleOffset"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.3
    
    iCount = 7
    m_InputTypes(iCount).Name = "NozzleOffset1"
    m_InputTypes(iCount).Description = "NozzleOffset1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 1
    
    iCount = 8
    m_InputTypes(iCount).Name = "NozzleOffset2"
    m_InputTypes(iCount).Description = "Nozzle Offset 2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.25
    
    iCount = 9
    m_InputTypes(iCount).Name = "NozzleOffset3"
    m_InputTypes(iCount).Description = "Nozzle Offset3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.25
    
    iCount = 10
    m_InputTypes(iCount).Name = "PortRotation1"
    m_InputTypes(iCount).Description = "Port Rotation of Nozz3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 11
    m_InputTypes(iCount).Name = "PortRotation2"
    m_InputTypes(iCount).Description = "Port Rotation of Nozz4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = Atn(1)
    
    iCount = 12
    m_InputTypes(iCount).Name = "InstrumentLength4"
    m_InputTypes(iCount).Description = "Instrument Length4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.7
    
    iCount = 13
    m_InputTypes(iCount).Name = "InstrumentLength5"
    m_InputTypes(iCount).Description = "Instrument Length5"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.6
     
    iCount = 14
    m_InputTypes(iCount).Name = "Npd"
    m_InputTypes(iCount).Description = "NPD"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 300
        
    iCount = 15
    m_InputTypes(iCount).Name = "EndPreparation"
    m_InputTypes(iCount).Description = "EndPreparation"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 21
    
    iCount = 16
    m_InputTypes(iCount).Name = "ScheduleThickness"
    m_InputTypes(iCount).Description = "ScheduleThickness"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 100
    
    iCount = 17
    m_InputTypes(iCount).Name = "EndStandard"
    m_InputTypes(iCount).Description = "EndStandard"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 5
    
    iCount = 18
    m_InputTypes(iCount).Name = "PressureRating"
    m_InputTypes(iCount).Description = "PressureRating"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 35
    
    iCount = 19
    m_InputTypes(iCount).Name = "FlowDirection"
    m_InputTypes(iCount).Description = "FlowDirection"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 3
            
    iCount = 20
    m_InputTypes(iCount).Name = "PortIndex1"
    m_InputTypes(iCount).Description = "PortIndex1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 1

    iCount = 21
    m_InputTypes(iCount).Name = "Npd1"
    m_InputTypes(iCount).Description = "NPD1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
        
    iCount = 22
    m_InputTypes(iCount).Name = "EndPreparation1"
    m_InputTypes(iCount).Description = "EndPreparation1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 23
    m_InputTypes(iCount).Name = "ScheduleThickness1"
    m_InputTypes(iCount).Description = "ScheduleThickness1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 24
    m_InputTypes(iCount).Name = "EndStandard1"
    m_InputTypes(iCount).Description = "EndStandard1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 25
    m_InputTypes(iCount).Name = "PressureRating1"
    m_InputTypes(iCount).Description = "PressureRating1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 26
    m_InputTypes(iCount).Name = "FlowDirection1"
    m_InputTypes(iCount).Description = "FlowDirection1"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 27
    m_InputTypes(iCount).Name = "PortIndex2"
    m_InputTypes(iCount).Description = "PortIndex2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 2

    iCount = 28
    m_InputTypes(iCount).Name = "Npd2"
    m_InputTypes(iCount).Description = "NPD2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
        
    iCount = 29
    m_InputTypes(iCount).Name = "EndPreparation2"
    m_InputTypes(iCount).Description = "EndPreparation2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 30
    m_InputTypes(iCount).Name = "ScheduleThickness2"
    m_InputTypes(iCount).Description = "ScheduleThickness2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 31
    m_InputTypes(iCount).Name = "EndStandard2"
    m_InputTypes(iCount).Description = "EndStandard2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 32
    m_InputTypes(iCount).Name = "PressureRating2"
    m_InputTypes(iCount).Description = "PressureRating2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 33
    m_InputTypes(iCount).Name = "FlowDirection2"
    m_InputTypes(iCount).Description = "FlowDirection2"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 34
    m_InputTypes(iCount).Name = "PortIndex3"
    m_InputTypes(iCount).Description = "PortIndex3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 3

    iCount = 35
    m_InputTypes(iCount).Name = "Npd3"
    m_InputTypes(iCount).Description = "NPD3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
        
    iCount = 36
    m_InputTypes(iCount).Name = "EndPreparation3"
    m_InputTypes(iCount).Description = "EndPreparation3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 37
    m_InputTypes(iCount).Name = "ScheduleThickness3"
    m_InputTypes(iCount).Description = "ScheduleThickness3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 38
    m_InputTypes(iCount).Name = "EndStandard3"
    m_InputTypes(iCount).Description = "EndStandard3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 39
    m_InputTypes(iCount).Name = "PressureRating3"
    m_InputTypes(iCount).Description = "PressureRating3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 40
    m_InputTypes(iCount).Name = "FlowDirection3"
    m_InputTypes(iCount).Description = "FlowDirection3"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 41
    m_InputTypes(iCount).Name = "PortIndex4"
    m_InputTypes(iCount).Description = "PortIndex4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 4

    iCount = 42
    m_InputTypes(iCount).Name = "Npd4"
    m_InputTypes(iCount).Description = "NPD4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
        
    iCount = 43
    m_InputTypes(iCount).Name = "EndPreparation4"
    m_InputTypes(iCount).Description = "EndPreparation4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 44
    m_InputTypes(iCount).Name = "ScheduleThickness4"
    m_InputTypes(iCount).Description = "ScheduleThickness4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 45
    m_InputTypes(iCount).Name = "EndStandard4"
    m_InputTypes(iCount).Description = "EndStandard4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 46
    m_InputTypes(iCount).Name = "PressureRating4"
    m_InputTypes(iCount).Description = "PressureRating4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 47
    m_InputTypes(iCount).Name = "FlowDirection4"
    m_InputTypes(iCount).Description = "FlowDirection4"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0
    
    iCount = 48
    m_InputTypes(iCount).Name = "InsulationThickness"
    m_InputTypes(iCount).Description = "Insulation Thickness"
    m_InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_InputTypes(iCount).uomValue = 0.025

    'Text Inputs
    m_iTextInputs = 9
    ReDim m_TextInputTypes(1 To m_iTextInputs) As TextInputType

    iCount = 1
    m_TextInputTypes(iCount).Name = "Id1"
    m_TextInputTypes(iCount).Description = "Id1"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "1"

    iCount = 2
    m_TextInputTypes(iCount).Name = "Id2"
    m_TextInputTypes(iCount).Description = "Id2"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "2"

    iCount = 3
    m_TextInputTypes(iCount).Name = "Id3"
    m_TextInputTypes(iCount).Description = "Id3"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "1"

    iCount = 4
    m_TextInputTypes(iCount).Name = "Id4"
    m_TextInputTypes(iCount).Description = "Id4"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "2"

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
    m_TextInputTypes(iCount).Description = "Npd Unit Type 2"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""
        
' Outputs
    m_iOutputs = 11
    ReDim m_OutputTypes(1 To m_iOutputs) As OutputType
    
    iCount = 1
    m_OutputTypes(iCount).Name = "vNoz1"
    m_OutputTypes(iCount).Description = "Nozzle 1"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    m_OutputTypes(iCount).Name = "vNoz2"
    m_OutputTypes(iCount).Description = "Nozzle 2"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 3
    m_OutputTypes(iCount).Name = "vNoz3"
    m_OutputTypes(iCount).Description = "Nozzle 3"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    m_OutputTypes(iCount).Name = "vNoz4"
    m_OutputTypes(iCount).Description = "Nozzle 4"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 5
    m_OutputTypes(iCount).Name = "CylinderBody1"
    m_OutputTypes(iCount).Description = "Cylinder Body1"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    m_OutputTypes(iCount).Name = "CylinderBody2"
    m_OutputTypes(iCount).Description = "Cylinder Body2"
    m_OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 7
    m_OutputTypes(iCount).Name = "ReducingCone"
    m_OutputTypes(iCount).Description = "Reducing Cone"
    m_OutputTypes(iCount).Aspect = SimplePhysical
        
    iCount = 8
    m_OutputTypes(iCount).Name = "DivergingCone"
    m_OutputTypes(iCount).Description = "Diverging Cone"
    m_OutputTypes(iCount).Aspect = SimplePhysical
          
    iCount = 9
    m_OutputTypes(iCount).Name = "VenturiBodyIns"
    m_OutputTypes(iCount).Description = "Venturi Body Insulation"
    m_OutputTypes(iCount).Aspect = Insulation
    
    iCount = 10
    m_OutputTypes(iCount).Name = "Nozz3BodyIns"
    m_OutputTypes(iCount).Description = "Nozzle3 Body Insulation"
    m_OutputTypes(iCount).Aspect = Insulation
    
    iCount = 11
    m_OutputTypes(iCount).Name = "Nozz4BodyIns"
    m_OutputTypes(iCount).Description = "Nozzle4 Body Insulation"
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
    Dim iCount      As Integer
        
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
    Dim TextInputs()    As IMSSymbolEntities.IJDInput
    
    'Create a default parameter
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    
    '----------Numeric Parameters -------------
    PC.Type = igValue
    
    ReDim Inputs(1 To m_iInputs) As IMSSymbolEntities.IJDInput
    For iCount = 1 To m_iInputs
        Set Inputs(iCount) = New IMSSymbolEntities.DInput
        Inputs(iCount).Name = m_InputTypes(iCount).Name
        Inputs(iCount).Description = m_InputTypes(iCount).Description
        Inputs(iCount).Properties = m_InputTypes(iCount).Properties
        PC.uomValue = m_InputTypes(iCount).uomValue
        'Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
        'Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next 'iCount
    
    '----------Text Parameters ----------------
    'Create a default parameter
    Dim PC1 As IMSSymbolEntities.IJDParameterContent
    Set PC1 = New IMSSymbolEntities.DParameterContent
    PC1.Type = igString

    ReDim TextInputs(1 To m_iTextInputs) As IMSSymbolEntities.IJDInput
    For iCount = 1 To m_iTextInputs
        Set TextInputs(iCount) = New IMSSymbolEntities.DInput
        TextInputs(iCount).Name = m_TextInputTypes(iCount).Name
        TextInputs(iCount).Description = m_TextInputTypes(iCount).Description
        TextInputs(iCount).Properties = m_TextInputTypes(iCount).Properties
        PC1.String = m_TextInputTypes(iCount).Value
        TextInputs(iCount).DefaultParameterValue = PC1
        InputsIf.SetInput TextInputs(iCount), m_iInputs + iCount + 1
        Set TextInputs(iCount) = Nothing
    Next 'iCount
    
    'Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    Set symbAspect = New IMSSymbolEntities.DRepresentation
    '---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput As Integer
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




