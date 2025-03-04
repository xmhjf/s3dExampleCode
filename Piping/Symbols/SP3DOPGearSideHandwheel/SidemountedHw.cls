VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SidemountedHw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   SidemountedHw.cls
'   ProgID:         SP3DOPGearSideHandwheel.SidemountedHw
'   Author:         RUK
'   Creation Date:  Wednesday, 2 April 2008
'
'   Description:
'   This Symbol details were taken from Appendix E-107 of the Design Document
'   This Symbol details were taken from Appendix E-107 of the Design Document
'   In this symbol below Valve operator PDBs are implemented
'   Default (1)
'   Gear, side mounted handwheel Type 1 (11)
'   Gear, side mounted handwheel Type 5 (mirror of Type 1) (15)
'   Gear, side mounted handwheel Type 11 (21)
'   Gear, side mounted handwheel Type 12 (mirror of Type 11) (24)
'   Gear, side mounted handwheel Type 11 with stem extension (25)
'   Gear, side mounted handwheel Type 12 with stem extension (mirror of Type 11 with stem extension) (26)
'   Gear, side mounted handwheel Type 11 with shaft (27)
'   Gear, side mounted handwheel Type 12 with shaft (mirror of Type 11 with shaft) (28)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Apr.2008     RUK     CR-CP-133524  Sample data for Dimensional Basis for valve operators is required
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SidemountedHw:"  'Used for error messages

Private m_SymbolFactory As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl As Object
Private m_ProgId As String
Private m_ProjName As String
Private m_ClassName As String

Private InputTypes() As InputType
Private OutputTypes() As OutputType
Private AspectTypes() As AspectType
Private nAspects As Integer
Private nInputs As Integer
Private nOutputs As Integer

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Dim iCount As Integer

    m_ProjName = "SP3DOPGearSideHandwheel"
    m_ClassName = "SidemountedHw"
    m_ProgId = m_ProjName & "." & m_ClassName

    'Inputs
    nInputs = 18
    ReDim InputTypes(1 To nInputs) As InputType

    iCount = 1
    InputTypes(iCount).name = "OperatorHeight"
    InputTypes(iCount).description = "Operator Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 2
    InputTypes(iCount).name = "OffsetFrmValCen"
    InputTypes(iCount).description = "Offset from Valve Centerline"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 3
    InputTypes(iCount).name = "GearDiameter"
    InputTypes(iCount).description = "Gear Diameter"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 4
    InputTypes(iCount).name = "GearRadius"
    InputTypes(iCount).description = "Gear Radius"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 5
    InputTypes(iCount).name = "GearDepth"
    InputTypes(iCount).description = "Gear Depth"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 6
    InputTypes(iCount).name = "HandwheelDiameter"
    InputTypes(iCount).description = "Handwheel Diameter"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 7
    InputTypes(iCount).name = "HandwheelRadius"
    InputTypes(iCount).description = "Handwheel Radius"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 8
    InputTypes(iCount).name = "HwXOffsetFrmOpStem"
    InputTypes(iCount).description = "Handwheel X-Offset from Operator Stem"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 9
    InputTypes(iCount).name = "HwZOffsetOpFrmStem"
    InputTypes(iCount).description = "Handwheel Z-Offset from Operator Stem"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 10
    InputTypes(iCount).name = "HwYOffsetFrmGearBot"
    InputTypes(iCount).description = "Handwheel Y Offset from Bottom of Gear Box"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 11
    InputTypes(iCount).name = "HwOffsetFrmValCen"
    InputTypes(iCount).description = "Handwheel Offset from Valve Centerline"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 12
    InputTypes(iCount).name = "HwOffsetFrmOpTop"
    InputTypes(iCount).description = "Handwheel Offset from Operator Top"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 13
    InputTypes(iCount).name = "ValCenToHWTop"
    InputTypes(iCount).description = "Valve Centerline to Top of Handwheel"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 14
    InputTypes(iCount).name = "ValCenToHWCen"
    InputTypes(iCount).description = "Valve Centerline to Handwheel Centerline"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 15
    InputTypes(iCount).name = "ExStemLen"
    InputTypes(iCount).description = "Extension Stem Length"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 16
    InputTypes(iCount).name = "ExStemLenFrmValCen"
    InputTypes(iCount).description = "Extension Stem Length from Valve Centerline"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 17
    InputTypes(iCount).name = "ExStemDia"
    InputTypes(iCount).description = "Extension Stem Diameter"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 18
    InputTypes(iCount).name = "ShaftHeight"
    InputTypes(iCount).description = "Shaft Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    'Outputs
    nOutputs = 8
    ReDim OutputTypes(1 To nOutputs) As OutputType

    iCount = 1
    OutputTypes(iCount).name = "GearBox"
    OutputTypes(iCount).description = "Gear Box"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 2
    OutputTypes(iCount).name = "HandwheelStem"
    OutputTypes(iCount).description = "Handwheel Stem"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 3
    OutputTypes(iCount).name = "Handwheel"
    OutputTypes(iCount).description = "Handwheel"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 4
    OutputTypes(iCount).name = "HandwheelSpoke1"
    OutputTypes(iCount).description = "Handwheel Spoke 1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 5
    OutputTypes(iCount).name = "HandwheelSpoke2"
    OutputTypes(iCount).description = "Handwheel Spoke 2"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 6
    OutputTypes(iCount).name = "HandwheelSpoke3"
    OutputTypes(iCount).description = "Handwheel Spoke 3"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 7
    OutputTypes(iCount).name = "MaintHandwheelStem"
    OutputTypes(iCount).description = "Maintenance Handwheel Stem"
    OutputTypes(iCount).Aspect = Maintenance
    
    iCount = 8
    OutputTypes(iCount).name = "MaintHandwheel"
    OutputTypes(iCount).description = "Maintenance Handwheel"
    OutputTypes(iCount).Aspect = Maintenance

    'Aspects
    nAspects = 2
    ReDim AspectTypes(1 To nAspects) As AspectType

    iCount = 1
    AspectTypes(iCount).name = "Physical"
    AspectTypes(iCount).description = "Physical"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = SimplePhysical

    iCount = 2
    AspectTypes(iCount).name = "Maintenance"
    AspectTypes(iCount).description = "Maintenance"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = Maintenance

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

    'This method is in charge of the creation of the symbol definition object
    'You can keep the current design unchanged
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler

    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition

    'Create the definition from the symbol factory
    Set m_SymbolFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set oSymbolDefinition = m_SymbolFactory.CreateEntity(Definition, ActiveConnection)
    Set m_SymbolFactory = Nothing
    'Initialize it
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    'Set definition progId and codebase
    oSymbolDefinition.ProgId = m_ProgId
    oSymbolDefinition.CodeBase = CodeBase

    'Give a unique name to the symbol definition
    oSymbolDefinition.name = oSymbolDefinition.ProgId

    'returned symbol definition
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

    Dim oPart As IMSSymbolEntities.DInput

    'Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition


    ' ----------Parameter Input --------------
    Dim Inputs() As IMSSymbolEntities.IJDInput
    ReDim Inputs(1 To nInputs) As IMSSymbolEntities.IJDInput
    Dim iCount As Integer
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
        'Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
        ' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount

    'Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    '---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs

    For iCount = 1 To nAspects
        Set symbAspect = New IMSSymbolEntities.DRepresentation
        symbAspect.name = AspectTypes(iCount).name
        symbAspect.description = AspectTypes(iCount).description
        symbAspect.properties = igCOLLECTION_VARIABLE
'        symbAspect.properties = AspectTypes(iCount).properties
        symbAspect.RepresentationId = AspectTypes(iCount).AspectId
        'symbAspect.RemoveAllOutput

        Set oSymbAspect = symbAspect
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).Aspect And AspectTypes(iCount).AspectId Then
                Output.name = OutputTypes(iOutput).name
                Output.description = OutputTypes(iOutput).description
                Output.properties = 0            ' OutputTypes(iOutput).Properties
                oSymbAspect.SetOutput Output
            End If
        Next iOutput
        RepsIf.SetRepresentation symbAspect
        Set Output = Nothing
        Set symbAspect = Nothing
    Next iCount

    ' ----------Definition of the function identification corresponding
    Dim oRepEval As IJDRepresentationEvaluation
    Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition

    For iCount = 1 To nAspects
        Set oRepEval = New DRepresentationEvaluation
        oRepEval.name = AspectTypes(iCount).name
        oRepEval.description = AspectTypes(iCount).description
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
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal OutputColl As Object, _
                                                      ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler

    Dim arrayOfOutputs() As String
    Dim iOutputs As Integer
    Dim iCount As Integer
    Dim lAspectID As SymbolRepIds
    Dim AspectName As String
    Dim bsucceded As Boolean

    Set m_OutputColl = OutputColl

    bsucceded = False
    For iCount = 1 To nAspects
        If AspectTypes(iCount).name = repName Then
            bsucceded = True
            lAspectID = AspectTypes(iCount).AspectId
            Exit For
        End If
    Next iCount
    If Not bsucceded Then
        'MsgBox "Unknown Aspect; Check Aspect data / code"
        Exit Sub
    End If

    iOutputs = 1
    For iCount = 1 To nOutputs
        If OutputTypes(iCount).Aspect = lAspectID Then
            ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
            arrayOfOutputs(iOutputs) = OutputTypes(iCount).name
            iOutputs = iOutputs + 1
        End If
    Next iCount
    
    Dim objAspect As Object
    If StrComp(repName, "Physical", vbTextCompare) = 0 Then
        Set objAspect = New CPhysical
        objAspect.run OutputColl, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If

    If StrComp(repName, "Maintenance", vbTextCompare) = 0 Then
        Set objAspect = New CMaintenance
        objAspect.run OutputColl, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If

    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    'The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
End Function
