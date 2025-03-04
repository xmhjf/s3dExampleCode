VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C5TCDCraneSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:50 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   C5TCDCraneSym.cls
'   ProgID:         SP3D5TCarryDeckCraneAsm.C5TCDCraneSym
'   Author:         CYW
'   Creation Date:  Thursday, Mar 27 2003
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   04.Apr.2012    Haneef  TR-CP-213871  Multiple outputs with the same name 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "C5TCDCraneSym:"  'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_outputColl        As Object
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
    
    m_ProjName = "SP3D5TCarryDeckCraneAsm"
    m_ClassName = "C5TCDCraneSym"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
            
' Outputs
    nOutputs = 30
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).Name = "Pyra1"
    OutputTypes(iCount).Description = "Pyra1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).Name = "Pyra2"
    OutputTypes(iCount).Description = "Pyra2"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 3
    OutputTypes(iCount).Name = "Pyra3"
    OutputTypes(iCount).Description = "Pyra3"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 4
    OutputTypes(iCount).Name = "Pyra4"
    OutputTypes(iCount).Description = "Pyra4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 5
    OutputTypes(iCount).Name = "Pyra5"
    OutputTypes(iCount).Description = "Pyra5"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 6
    OutputTypes(iCount).Name = "Box1"
    OutputTypes(iCount).Description = "Box1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 7
    OutputTypes(iCount).Name = "Pyra6"
    OutputTypes(iCount).Description = "Pyra6"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 8
    OutputTypes(iCount).Name = "Pyra7"
    OutputTypes(iCount).Description = "Pyra7"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 9
    OutputTypes(iCount).Name = "Pyra8"
    OutputTypes(iCount).Description = "Pyra8"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 10
    OutputTypes(iCount).Name = "Pyra9"
    OutputTypes(iCount).Description = "Pyra9"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 11
    OutputTypes(iCount).Name = "Cyli1"
    OutputTypes(iCount).Description = "Cyli1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 12
    OutputTypes(iCount).Name = "Cyli2"
    OutputTypes(iCount).Description = "Cyli2"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 13
    OutputTypes(iCount).Name = "Cyli3"
    OutputTypes(iCount).Description = "Cyli3"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 14
    OutputTypes(iCount).Name = "Cyli4"
    OutputTypes(iCount).Description = "Cyli4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 15
    OutputTypes(iCount).Name = "Cyli5"
    OutputTypes(iCount).Description = "Cyli5"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 16
    OutputTypes(iCount).Name = "Dish1"
    OutputTypes(iCount).Description = "Dish1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 17
    OutputTypes(iCount).Name = "Dish2"
    OutputTypes(iCount).Description = "Dish2"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 18
    OutputTypes(iCount).Name = "Dish3"
    OutputTypes(iCount).Description = "Dish3"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 19
    OutputTypes(iCount).Name = "Dish4"
    OutputTypes(iCount).Description = "Dish4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 20
    OutputTypes(iCount).Name = "Dish5"
    OutputTypes(iCount).Description = "Dish5"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 21
    OutputTypes(iCount).Name = "Dish6"
    OutputTypes(iCount).Description = "Dish6"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 22
    OutputTypes(iCount).Name = "Dish7"
    OutputTypes(iCount).Description = "Dish7"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 23
    OutputTypes(iCount).Name = "Dish8"
    OutputTypes(iCount).Description = "Dish8"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 24
    OutputTypes(iCount).Name = "Box2"
    OutputTypes(iCount).Description = "Box2"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 25
    OutputTypes(iCount).Name = "Box3"
    OutputTypes(iCount).Description = "Box3"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 26
    OutputTypes(iCount).Name = "Box4"
    OutputTypes(iCount).Description = "Box4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 27
    OutputTypes(iCount).Name = "Pyra10"
    OutputTypes(iCount).Description = "Pyra10"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 28
    OutputTypes(iCount).Name = "Pyra11"
    OutputTypes(iCount).Description = "Pyra11"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 29
    OutputTypes(iCount).Name = "Pyra12"
    OutputTypes(iCount).Description = "Pyra12"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 30
    OutputTypes(iCount).Name = "Cyli6"
    OutputTypes(iCount).Description = "Cyli6"
    OutputTypes(iCount).Aspect = SimplePhysical

    
    
' Aspects
    nAspects = 1
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    iCount = 1
    AspectTypes(iCount).Name = "SimplePhysical"
    AspectTypes(iCount).Description = "SimplePhysical"
    AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
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
    'Dim Inputs()    As IMSSymbolEntities.IJDInput
    'ReDim Inputs(1 To nInputs) As IMSSymbolEntities.IJDInput
    Dim iCount      As Integer
' Create a default parameter
    'Dim PC As IMSSymbolEntities.DParameterContent
    'Set PC = New IMSSymbolEntities.DParameterContent
    'PC.Type = igValue
    'For iCount = 1 To nInputs
        'Set Inputs(iCount) = New IMSSymbolEntities.DInput
        'Inputs(iCount).Name = InputTypes(iCount).Name
        'Inputs(iCount).Description = InputTypes(iCount).Description
        'Inputs(iCount).Properties = InputTypes(iCount).Properties
        'PC.uomValue = InputTypes(iCount).uomValue
' Associate a default parameter to the input
        'Inputs(iCount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        'InputsIf.SetInput Inputs(iCount), iCount + 1
        'Set Inputs(iCount) = Nothing
    'Next iCount
        
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
    
    Set m_outputColl = outputcoll
    
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
    If StrComp(repName, "SimplePhysical", vbTextCompare) = 0 Then
        Set objAspect = New CSimplePhysical
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
