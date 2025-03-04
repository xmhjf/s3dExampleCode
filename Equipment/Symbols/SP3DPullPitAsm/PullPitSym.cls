VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PullPitSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   PullPitSym.cls
'   ProgID:        SP3DPullPitAsm.PullPitSym
'   Author:        RH
'   Creation Date: 01-May-08
'   Description:
'       This is Electrical Pull pit Assmebly.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   13.June.2008     VRK     CR-134560:Provide pull-pit/manhole equipment symbol for use with duct banks
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "PullPitSym:"  'Used for error messages

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

    m_ProjName = "SP3DPullPitAsm"
    m_ClassName = "PullPitSym"
    m_ProgId = m_ProjName & "." & m_ClassName

    ' Inputs
    nInputs = 14
    ReDim InputTypes(1 To nInputs) As InputType

    iCount = 1
    InputTypes(iCount).name = "PullBoxHeight"
    InputTypes(iCount).Description = "A-Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1.5

    iCount = 2
    InputTypes(iCount).name = "PullBoxWidth"
    InputTypes(iCount).Description = "B-Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1.2

    iCount = 3
    InputTypes(iCount).name = "PullBoxLength"
    InputTypes(iCount).Description = "C-Length"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1.2

    iCount = 4
    InputTypes(iCount).name = "PullBoxThickness"
    InputTypes(iCount).Description = "D-Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05

    iCount = 5
    InputTypes(iCount).name = "PullBoxHole1Height"
    InputTypes(iCount).Description = "E-First Hole Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.56

    iCount = 6
    InputTypes(iCount).name = "PullBoxHole1Width"
    InputTypes(iCount).Description = "F-First Hole Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.87

    iCount = 7
    InputTypes(iCount).name = "PullBoxHole1Offset"
    InputTypes(iCount).Description = "G-First Hole Vertical Offset"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.3
    
    iCount = 8
    InputTypes(iCount).name = "PullBoxHole1ThruStart"
    InputTypes(iCount).Description = "First Hole Penetrates Start"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1 'True
    
    iCount = 9
    InputTypes(iCount).name = "PullBoxHole1ThruEnd"
    InputTypes(iCount).Description = "First Hole Penetrates End"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1 'True
    
    iCount = 10
    InputTypes(iCount).name = "PullBoxHole2Height"
    InputTypes(iCount).Description = "H-Second Hole Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.56

    iCount = 11
    InputTypes(iCount).name = "PullBoxHole2Width"
    InputTypes(iCount).Description = "I-Second Hole Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.87

    iCount = 12
    InputTypes(iCount).name = "PullBoxHole2Offset"
    InputTypes(iCount).Description = "J-Second Hole Vertical Offset"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.3
    
    iCount = 13
    InputTypes(iCount).name = "PullBoxHole2ThruStart"
    InputTypes(iCount).Description = "Second Hole Penetrates Start"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1 'True
    
    iCount = 14
    InputTypes(iCount).name = "PullBoxHole2ThruEnd"
    InputTypes(iCount).Description = "Second Hole Penetrates End"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1 'True
    
    'Outputs
    nOutputs = 1
    ReDim OutputTypes(1 To nOutputs) As OutputType

    iCount = 1
    OutputTypes(iCount).name = "DefaultSurface"
    OutputTypes(iCount).Description = "DefaultSurface"
    OutputTypes(iCount).Aspect = SimplePhysical

    ' Aspects
    nAspects = 1
    ReDim AspectTypes(1 To nAspects) As AspectType

    iCount = 1
    AspectTypes(iCount).name = "Physical"
    AspectTypes(iCount).Description = "Physical"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = SimplePhysical

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

    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition

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
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ProgId
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler

    Dim oPart As IMSSymbolEntities.DInput

    ' Remove all previous Symbol Definition information
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
        Inputs(iCount).Description = InputTypes(iCount).Description
        Inputs(iCount).properties = InputTypes(iCount).properties
        PC.uomValue = InputTypes(iCount).uomValue
        ' Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
        ' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount

    ' Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    ' ---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs

    For iCount = 1 To nAspects
        Set symbAspect = New IMSSymbolEntities.DRepresentation
        symbAspect.name = AspectTypes(iCount).name
        symbAspect.Description = AspectTypes(iCount).Description
        symbAspect.properties = AspectTypes(iCount).properties
        symbAspect.RepresentationId = AspectTypes(iCount).AspectId
        '        symbAspect.RemoveAllOutput

        Set oSymbAspect = symbAspect
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).Aspect And AspectTypes(iCount).AspectId Then
                Output.name = OutputTypes(iOutput).name
                Output.Description = OutputTypes(iOutput).Description
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
        oRepEval.Description = AspectTypes(iCount).Description
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
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler

    Dim arrayOfOutputs() As String
    Dim iOutputs As Integer
    Dim iCount As Integer
    Dim lAspectID As SymbolRepIds
    Dim AspectName As String
    Dim bsucceded As Boolean

    Set m_OutputColl = outputcoll

    bsucceded = False
    For iCount = 1 To nAspects
        If AspectTypes(iCount).name = repName Then
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
            arrayOfOutputs(iOutputs) = OutputTypes(iCount).name
            iOutputs = iOutputs + 1
        End If
    Next iCount
    Dim objAspect As Object
    If StrComp(repName, "Physical", vbTextCompare) = 0 Then
        Set objAspect = New CPhysical
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
