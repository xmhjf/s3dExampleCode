VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C45Elbow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   C45Elbow.cls
'   ProgID:         SP3D45Elbow.C45Elbow
'   Author:         NN
'   Creation Date:  Sunday, Nov 26 2000
'   Description:
'   45 degree trimmable elbow.

'   The symbol consists of six outputs, three physical and three insulation outputs.
'   The symbol is enhanced to meet the requirements of following sources:
'           1)ASTM A 74, Standard Specification for Cast Iron Soil Pipe and Fittings
'           2)ASTM A 861, Standard Specification for High-Silicon Iron Pipe and Fittings
'           3)Tyler Soil Pipe Catalog
'           4)ABI Foundry Catalog
'           5)AMERICAN Ductile Iron Fastite®, Flex-Ring®, and Lok-Ring® Fittings
'           6)AWWA C110/A21 Catalog
'   The symbol implements the following Part Data Bases,
' Symmetrical Elbow-
' a) Default or Face-to-center dimension basis-10
' b) Seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth).
' c) Elbow, Symmetrical, specified by Seat to Center and Bend Radius-163
' d) Elbow, Symmetrical, specified by Face to Center and Bend Radius-164
' e) Elbow, Symmetrical, specified Seat to Center, Tangent length and Bend Radius-161
' f) Elbow, Symmetrical, specified Face to Center, Tangent length and Bend Radius-162

' Asymmetrical Elbow-
' a) Asymmetrical Face-to-Center dimension basis-13
' b) Asymmetrical seat-to-center dimension (Face-to-center dimension excludes the socket depth or thread depth)-36
' c) Elbow, Asymmetrical, specified by Seat to Center 1, Seat to Center 2 and Bend Radius-176
' d) Elbow, Asymmetrical, specified by Face to Center 1, Face to Center 2 and Bend Radius-177
' e) Elbow, Asymmetrical, specified Seat to Center 1, Seat to Center 2, Tangent length 1, Tangent Length 2 and Bend Radius-179
' f) Elbow, Asymmetrical, specified by Face-to-Center 1, Face-to-Center 2, Tangent 1, Tangent 2, and Bend Radius-180
' g) Elbow based on bend radius and leg offset-55
' h) Elbow, asymmetrical, specified by tangent 1, face-to-center 2

'   NOTE :
'   1) The user should specify the parameters required for a specific case(Even though
'   they are defined as Optional).
'   2) In cases a,b of symmetrical and asymmetrical elbows, the Elbow Radius is computed based on Face to Center and bend angle
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     ---                  ------------------
'   09.Jul.2003     SymbolTeam(India)    Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)    TR 56826  Removed Msgbox
'   08.SEP.2006     KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.Nov.2007     VRG                  CR-127554 Enhanced the existing symbol code to support 1/8 bend, 1/8 bend-extended, 1/8 bend-long cases using PDB values
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   28.Jul.2008     KKC                 CR-145699 Enhance 45 degree and 90 degree elbow symbol to support JIS case
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "C45Elbow:"  'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String
Private m_ProjName          As String
Private m_ClassName         As String

Private InputTypes()        As InputType
Private OutputTypes()       As OutputType
Private AspectTypes()       As AspectType
Private nAspects            As Integer
Private nInputs             As Integer
Private nOutputs            As Integer
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProjName = "SP3D45Elbow"
    m_ClassName = "C45Elbow"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 12
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).name = "FacetoCenter"
    InputTypes(iCount).description = "Face to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0#
        
    iCount = 2
    InputTypes(iCount).name = "InsulationThickness"
    InputTypes(iCount).description = "Insulation Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025
    
    iCount = 3
    InputTypes(iCount).name = "ElbowRadius"
    InputTypes(iCount).description = "Elbow Radius"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0#
        
    iCount = 4
    InputTypes(iCount).name = "SeattoCenter"
    InputTypes(iCount).description = "Seat to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 5
    InputTypes(iCount).name = "Seat1toCenter"
    InputTypes(iCount).description = "Seat 1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 6
    InputTypes(iCount).name = "Seat2toCenter"
    InputTypes(iCount).description = "Seat 1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 7
    InputTypes(iCount).name = "Face1toCenter"
    InputTypes(iCount).description = "Face 1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 8
    InputTypes(iCount).name = "Face2toCenter"
    InputTypes(iCount).description = "Face 2 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 9
    InputTypes(iCount).name = "TangentLength"
    InputTypes(iCount).description = "Tangent Length"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 10
    InputTypes(iCount).name = "TangentLength1"
    InputTypes(iCount).description = "Tangent Length 1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
            
    iCount = 11
    InputTypes(iCount).name = "TangentLength2"
    InputTypes(iCount).description = "Tangent Length 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 12
    InputTypes(iCount).name = "Angle"
    InputTypes(iCount).description = "Trim Angle"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = Atn(1) '45 degrees
    
        
' Outputs
    nOutputs = 6
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "Body"
    OutputTypes(iCount).description = "Body"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).name = "InsulatedBody"
    OutputTypes(iCount).description = "Insulated Body"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 3
    OutputTypes(iCount).name = "InsulatedPort1"
    OutputTypes(iCount).description = "Insulated Port1"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 4
    OutputTypes(iCount).name = "InsulatedPort2"
    OutputTypes(iCount).description = "Insulated Port2"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 5
    OutputTypes(iCount).name = "PNoz1"
    OutputTypes(iCount).description = "Nozzle 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).name = "PNoz2"
    OutputTypes(iCount).description = "Nozzle 2"
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
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    
' ---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For iCount = 1 To nAspects
        Set symbAspect = New IMSSymbolEntities.DRepresentation
        symbAspect.name = AspectTypes(iCount).name
        symbAspect.description = AspectTypes(iCount).description
        symbAspect.properties = igCOLLECTION_VARIABLE 'AspectTypes(iCount).properties
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
        Set symbAspect = Nothing
        Set Output = Nothing
    Next iCount
    

    ' ----------Definition of the function identification corresponding
    Dim oRepEval     As IJDRepresentationEvaluation
    Dim RepsEvalsIf     As IMSSymbolEntities.IJDRepresentationEvaluations
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
    
    Dim arrayOfOutputs()    As String
    Dim iOutputs            As Integer
    Dim iCount              As Integer
    Dim lAspectID           As SymbolRepIds
    Dim AspectName          As String
    Dim bsucceded           As Boolean
    
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
