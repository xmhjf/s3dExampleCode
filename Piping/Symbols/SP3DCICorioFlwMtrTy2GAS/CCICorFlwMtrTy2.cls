VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCICorFlwMtrTy2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCICorFlwMtrTy2.cls.cls
'   ProgID:         SP3DCICorioFlwMtrTy2GAS.CCICorFlwMtrTy2
'   Author:         BG
'   Creation Date:  Thusday, Dec 26 2002
'   Description:
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM3,
'   PDMS Category Name=IFMZ). This symbol is Custom INtools Instrument with Nozzles are
'   fully Parametric. Physical aspect consist of 1 main cylinder body, 2 Support
'   Cylinder and a Box. Insulation aspect consist of a Cylinder and a box.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.  
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CCICorFlwMtrTy2:"  'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String
Private m_ProjName          As String
Private m_ClassName         As String

Private InputTypes()        As InputType
Private m_TextInputTypes() As TextInputType
Private OutputTypes()       As OutputType
Private AspectTypes()       As AspectType
Private nAspects            As Integer
Private nInputs             As Integer
Private nTextInputs       As Integer
Private nOutputs            As Integer
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProjName = "SP3DCICorioFlwMtrTy2GAS"
    m_ClassName = "CCICorFlwMtrTy2"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 28
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).Name = "FacetoFace"
    InputTypes(iCount).Description = "FacetoFace"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.5
    
    iCount = 2
    InputTypes(iCount).Name = "FlowDiameter"
    InputTypes(iCount).Description = "Flow Diameter"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 3
    InputTypes(iCount).Name = "InstrumentHeight"
    InputTypes(iCount).Description = "Instrument Height"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.6
    
    iCount = 4
    InputTypes(iCount).Name = "InstrumentLength"
    InputTypes(iCount).Description = "Instrument Length 1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.4
    
    iCount = 5
    InputTypes(iCount).Name = "InstrumentHeight1"
    InputTypes(iCount).Description = "Instrument Height 1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.4
    
    iCount = 6
    InputTypes(iCount).Name = "InstrumentWidth"
    InputTypes(iCount).Description = "Instrument Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 7
    InputTypes(iCount).Name = "InstrumentWidth1"
    InputTypes(iCount).Description = "Instrument Width 1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 8
    InputTypes(iCount).Name = "InsulationThickness"
    InputTypes(iCount).Description = "Insulation Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025
    
    iCount = 9
    InputTypes(iCount).Name = "Npd"
    InputTypes(iCount).Description = "NPD"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 200
        
    iCount = 10
    InputTypes(iCount).Name = "EndPreparation"
    InputTypes(iCount).Description = "EndPreparation"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 21
    
    iCount = 11
    InputTypes(iCount).Name = "ScheduleThickness"
    InputTypes(iCount).Description = "ScheduleThickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 100
    
    iCount = 12
    InputTypes(iCount).Name = "EndStandard"
    InputTypes(iCount).Description = "EndStandard"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 5
    
    iCount = 13
    InputTypes(iCount).Name = "PressureRating"
    InputTypes(iCount).Description = "PressureRating"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 35
    
    iCount = 14
    InputTypes(iCount).Name = "FlowDirection"
    InputTypes(iCount).Description = "FlowDirection"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 3
            
    iCount = 15
    InputTypes(iCount).Name = "PortIndex1"
    InputTypes(iCount).Description = "PortIndex1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1

    iCount = 16
    InputTypes(iCount).Name = "Npd1"
    InputTypes(iCount).Description = "NPD1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
        
    iCount = 17
    InputTypes(iCount).Name = "EndPreparation1"
    InputTypes(iCount).Description = "EndPreparation1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 18
    InputTypes(iCount).Name = "ScheduleThickness1"
    InputTypes(iCount).Description = "ScheduleThickness1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 19
    InputTypes(iCount).Name = "EndStandard1"
    InputTypes(iCount).Description = "EndStandard1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 20
    InputTypes(iCount).Name = "PressureRating1"
    InputTypes(iCount).Description = "PressureRating1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 21
    InputTypes(iCount).Name = "FlowDirection1"
    InputTypes(iCount).Description = "FlowDirection1"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 22
    InputTypes(iCount).Name = "PortIndex2"
    InputTypes(iCount).Description = "PortIndex2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2

    iCount = 23
    InputTypes(iCount).Name = "Npd2"
    InputTypes(iCount).Description = "NPD2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
        
    iCount = 24
    InputTypes(iCount).Name = "EndPreparation2"
    InputTypes(iCount).Description = "EndPreparation2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 25
    InputTypes(iCount).Name = "ScheduleThickness2"
    InputTypes(iCount).Description = "ScheduleThickness2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 26
    InputTypes(iCount).Name = "EndStandard2"
    InputTypes(iCount).Description = "EndStandard2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 27
    InputTypes(iCount).Name = "PressureRating2"
    InputTypes(iCount).Description = "PressureRating2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 28
    InputTypes(iCount).Name = "FlowDirection2"
    InputTypes(iCount).Description = "FlowDirection2"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0

    'Text Inputs
    nTextInputs = 5
    ReDim m_TextInputTypes(1 To nTextInputs) As TextInputType

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
    m_TextInputTypes(iCount).Name = "NpdUnitType"
    m_TextInputTypes(iCount).Description = "Npd Unit Type"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = "mm"

    iCount = 4
    m_TextInputTypes(iCount).Name = "NpdUnitType1"
    m_TextInputTypes(iCount).Description = "Npd Unit Type 1"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""

    iCount = 5
    m_TextInputTypes(iCount).Name = "NpdUnitType2"
    m_TextInputTypes(iCount).Description = "Npd Unit Type 2"
    m_TextInputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    m_TextInputTypes(iCount).Value = ""
    
' Outputs
    nOutputs = 10
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
    OutputTypes(iCount).Name = "FlowMtrBody"
    OutputTypes(iCount).Description = "Flow Meter Body"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).Name = "Cylinder1"
    OutputTypes(iCount).Description = "Cylinder 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 5
    OutputTypes(iCount).Name = "Cylinder2"
    OutputTypes(iCount).Description = "Cylinder 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).Name = "Box"
    OutputTypes(iCount).Description = "Box"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 7
    OutputTypes(iCount).Name = "BodyIns"
    OutputTypes(iCount).Description = "Body Insulation"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 8
    OutputTypes(iCount).Name = "Cylin1Ins"
    OutputTypes(iCount).Description = "Cylinder 1 Insulation"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 9
    OutputTypes(iCount).Name = "Cylin2Ins"
    OutputTypes(iCount).Description = "Cylinder 2 Insulation"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 10
    OutputTypes(iCount).Name = "BoxIns"
    OutputTypes(iCount).Description = "Box Insulation"
    OutputTypes(iCount).Aspect = Insulation
    
' Aspects
    nAspects = 2
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
''uncomment to make test with text inputs
    Dim PC1 As IMSSymbolEntities.IJDParameterContent
    Set PC1 = New IMSSymbolEntities.DParameterContent
    PC1.Type = igString
    ReDim TextInputs(1 To nTextInputs) As IMSSymbolEntities.IJDInput
    
    For iCount = 1 To nTextInputs
        Set TextInputs(iCount) = New IMSSymbolEntities.DInput
        TextInputs(iCount).Name = m_TextInputTypes(iCount).Name
        TextInputs(iCount).Description = m_TextInputTypes(iCount).Description
        TextInputs(iCount).Properties = m_TextInputTypes(iCount).Properties
        PC1.String = m_TextInputTypes(iCount).Value
        TextInputs(iCount).DefaultParameterValue = PC1
        InputsIf.SetInput TextInputs(iCount), nInputs + iCount + 1
        Set TextInputs(iCount) = Nothing
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
    
        
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function
