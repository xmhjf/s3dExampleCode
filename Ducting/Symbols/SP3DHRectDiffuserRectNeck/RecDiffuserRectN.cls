VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecDiffuserRectN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   RecDiffuserRectN.cls
'   ProgID:         SP3DHRecDiffuserRectN.RecDiffuserRectN
'   Author:         kkk
'   Creation Date:  Tuesday, Feb 14 2006
'   Description:
'       This symbol is taken from Carnes HVAC Catalog.
'   Partdatabasis cases implemented
'   Rectangular diffuser, with rectangular neck, Type A-161
'   Rectangular diffuser, with rectangular neck, Type B -162
'   Rectangular diffuser, with rectangular neck, Type C -163
'   Rectangular diffuser, with rectangular neck, Type D -164

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670     Replace names with initials in all revision history sheets and symbols
'  21.APR.2008     PK   DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'  22.SEP.2008     PK   CR-149529    Modified the symbol code for the partdatabasis cases 161,162,163,164
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "RecDiffuserRectN:"  'Used for error messages

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
    
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProjName = "SP3DHRecDiffuserRectN"
    m_ClassName = "RecDiffuserRectN"
    m_ProgId = m_ProjName & "." & m_ClassName

    'Inputs
    nInputs = 6
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).name = "Width"
    InputTypes(iCount).description = " Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 2
    InputTypes(iCount).name = "Depth"
    InputTypes(iCount).description = " Depth"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 3
    InputTypes(iCount).name = "BWidth"
    InputTypes(iCount).description = " Diffuser Outer Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 4
    InputTypes(iCount).name = "BDepth"
    InputTypes(iCount).description = " Diffuser Outer Depth"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 5
    InputTypes(iCount).name = "DiffuserThickness"
    InputTypes(iCount).description = " Diffuser Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 6
    InputTypes(iCount).name = "DiffuserLength1"
    InputTypes(iCount).description = "Diffuser Length1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
        
    'Outputs
    nOutputs = 2
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "Neck"
    OutputTypes(iCount).description = "Rectangular Diffuser Neck"
    OutputTypes(iCount).Aspect = SimplePhysical
      
    iCount = 2
    OutputTypes(iCount).name = "DiffuserPort"
    OutputTypes(iCount).description = "Diffuser Port"
    OutputTypes(iCount).Aspect = SimplePhysical

    'Aspects
    nAspects = 1
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    iCount = 1
    AspectTypes(iCount).name = "Physical"
    AspectTypes(iCount).description = "Physical"
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
    
    'This method is in charge of the creation of the symbol definition object
    'You can keep the current design unchanged
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler
    
    Dim oSymbolDefinition  As IMSSymbolEntities.IJDSymbolDefinition
    
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
    Debug.Assert False

End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
  IJDUserSymbolServices_GetDefinitionName = m_ProgId

End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler
   
    Dim oPart   As IMSSymbolEntities.DInput
    
    'Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition
       
    '----------Parameter Input --------------
    Dim Inputs()    As IMSSymbolEntities.IJDInput
    ReDim Inputs(1 To nInputs) As IMSSymbolEntities.IJDInput
    Dim iCount      As Integer
    'Create a default parameter
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
        '---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount
        
    'Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    '---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For iCount = 1 To nAspects
        Set symbAspect = New IMSSymbolEntities.DRepresentation
        symbAspect.name = AspectTypes(iCount).name
        symbAspect.description = AspectTypes(iCount).description
        symbAspect.properties = igCOLLECTION_VARIABLE
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
        Set symbAspect = New IMSSymbolEntities.DRepresentation
    Next iCount
    

    '----------Definition of the function identification corresponding
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
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
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
