VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEqualTee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CEqualTee.cls
'   ProgID:         SP3DTee.CEqualTee
'   Author:         NN
'   Creation Date:  Sunday, Nov 5 2000
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy    who                     change description
'   -----------    ---                     ------------------
'   17-11-2000     NN                    Insulation Aspect, insulated flanges
'   09.Jul.2003    SymbolTeam(India)     Copyright Information, Header  is added/Updated.
'   19.Mar.2004    SymbolTeam(India)     TR 56826 Removed Msgbox
'   11.Sep.2006    KKC                   CR 104167 Request to update tee symbol to optionally support 2 dimensions
'   14-SEP-06      KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols.
'   21-Dec-06      svsmylav              TR-112718: Face2toCenter and Face3toCenter are optional inputs added in V7 -
'                                          these are moved to the end of first two V6.1 existing inputs (otherwise mirror operation failed).
'   19-Feb-07      svsmylav              CR-113435: Optional input 'Angle' is added to support Trimmable Tee.
'   18-Mar-08      KKC                   CR-127565 and CR-127567. Created the Sanitary Tee symbol with various gemotric options
'   21.APR.2008    PK                    DI-132538:CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   12-May-08      RRK                   CR-141196:Enhanced the symbol to support long sweep tees with part data basis values of 862, 863 and 864
'   05.July.08     RUK                   CR-CP-146405  Enhance Tee symbol for vortex tee geometry per JIS G 5527
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CEqualTee:"  'Used for error messages

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

    m_ProjName = "SP3DTee"
    m_ClassName = "CEqualTee"
    m_ProgId = m_ProjName & "." & m_ClassName

    ' Inputs
    nInputs = 29
    ReDim InputTypes(1 To nInputs) As InputType

    iCount = 1
    InputTypes(iCount).name = "FacetoCenter"
    InputTypes(iCount).description = "Face to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.10477

    iCount = 2
    InputTypes(iCount).name = "InsulationThickness"
    InputTypes(iCount).description = "Insulation Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025

    iCount = 3
    InputTypes(iCount).name = "Face2toCenter"
    InputTypes(iCount).description = "Face1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 4
    InputTypes(iCount).name = "Face3toCenter"
    InputTypes(iCount).description = "Face1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 5
    InputTypes(iCount).name = "Angle"
    InputTypes(iCount).description = "Angle of Branch measured from Header"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 2 * Atn(1)    'PI/2 radian

    iCount = 6
    InputTypes(iCount).name = "Face1toFace4"
    InputTypes(iCount).description = "Face 1 to Face 4"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.25

    iCount = 7
    InputTypes(iCount).name = "Face1toSeat2"
    InputTypes(iCount).description = "Face 1 to Seat 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.25

    iCount = 8
    InputTypes(iCount).name = "Seat3toCenter"
    InputTypes(iCount).description = "Seat 3 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.25

    iCount = 9
    InputTypes(iCount).name = "PlugOuttoCenter"
    InputTypes(iCount).description = "Plug Outlet to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.25

    iCount = 10
    InputTypes(iCount).name = "TangentLength1"
    InputTypes(iCount).description = "Face 4 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.25

    iCount = 11
    InputTypes(iCount).name = "TangentLength3"
    InputTypes(iCount).description = "Branch Direction"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 12
    InputTypes(iCount).name = "FacetoFace"
    InputTypes(iCount).description = "Face to Face"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 13
    InputTypes(iCount).name = "Face1toCenter"
    InputTypes(iCount).description = "Face 1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 14
    InputTypes(iCount).name = "Seat2toCenter"
    InputTypes(iCount).description = "Seat 2 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 15
    InputTypes(iCount).name = "OutOffsetPort1"
    InputTypes(iCount).description = "Face 1 to Tap"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 16
    InputTypes(iCount).name = "Seat1toSeat2"
    InputTypes(iCount).description = "Seat 1 to Seat 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 17
    InputTypes(iCount).name = "Seat1toCenter"
    InputTypes(iCount).description = "Seat 1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 18
    InputTypes(iCount).name = "BasetoCenter"
    InputTypes(iCount).description = "Base to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 19
    InputTypes(iCount).name = "BaseDiam"
    InputTypes(iCount).description = "Base Length"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 20
    InputTypes(iCount).name = "OffBranch1toBranch2"
    InputTypes(iCount).description = "Length between two taps"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 21
    InputTypes(iCount).name = "BranchRadius"
    InputTypes(iCount).description = "Branch Radius"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 22
    InputTypes(iCount).name = "BaseWidth"
    InputTypes(iCount).description = "Base Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 23
    InputTypes(iCount).name = "BaseThickness"
    InputTypes(iCount).description = "Base Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 24
    InputTypes(iCount).name = "RibThickness"
    InputTypes(iCount).description = "Rib Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 25
    InputTypes(iCount).name = "OutFacetoCenter"
    InputTypes(iCount).description = "Outlet Face to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 26
    InputTypes(iCount).name = "OutSeattoCenter"
    InputTypes(iCount).description = "Outlet Seat to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 27
    InputTypes(iCount).name = "Diameter"
    InputTypes(iCount).description = "Body Diameter"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 28
    InputTypes(iCount).name = "Length"
    InputTypes(iCount).description = "Length"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 29
    InputTypes(iCount).name = "SeattoSeat"
    InputTypes(iCount).description = "Seat to Seat"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    ' Outputs
    nOutputs = 8
    ReDim OutputTypes(1 To nOutputs) As OutputType

    iCount = 1
    OutputTypes(iCount).name = "InsulatedBody"
    OutputTypes(iCount).description = "Insulated Body"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 2
    OutputTypes(iCount).name = "InsulatedPort1"
    OutputTypes(iCount).description = "Insulated Port1"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 3
    OutputTypes(iCount).name = "InsulatedPort2"
    OutputTypes(iCount).description = "Insulated Port2"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 4
    OutputTypes(iCount).name = "InsulatedBranch"
    OutputTypes(iCount).description = "Insulated Branch"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 5
    OutputTypes(iCount).name = "InsulatedPort3"
    OutputTypes(iCount).description = "Insulated Port3"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 6
    OutputTypes(iCount).name = "PNoz1"
    OutputTypes(iCount).description = "Nozzle 1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 7
    OutputTypes(iCount).name = "PNoz2"
    OutputTypes(iCount).description = "Nozzle 2"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 8
    OutputTypes(iCount).name = "PNoz3"
    OutputTypes(iCount).description = "Nozzle 3"
    OutputTypes(iCount).Aspect = SimplePhysical

    ' Aspects
    nAspects = 2
    ReDim AspectTypes(1 To nAspects) As AspectType

    iCount = 1
    AspectTypes(iCount).name = "Physical"
    AspectTypes(iCount).description = "Physical"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = SimplePhysical
    iCount = 2
    AspectTypes(iCount).name = "Insulation"
    AspectTypes(iCount).description = "Insulation"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
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
    Debug.Assert False

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
        symbAspect.description = AspectTypes(iCount).description
        '        symbAspect.properties = AspectTypes(iCount).properties
        symbAspect.properties = igCOLLECTION_VARIABLE
        symbAspect.RepresentationId = AspectTypes(iCount).AspectId
        '        symbAspect.RemoveAllOutput

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
    Debug.Assert False

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
