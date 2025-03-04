VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFTruckSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:05:19 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CForkTruck.cls
'   ProgID:         SP3DForkTruck.CForkTruck
'   Author:         nka8226
'   Creation Date:  Friday, Nov 22 2002
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)       TR 56826 Removed Msgbox
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CFTruckSym:"  'Used for error messages

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
    
    m_ProjName = "SP3DForkTruckAsm"
    m_ClassName = "CFTruckSym"
    m_ProgId = m_ProjName & "." & m_ClassName

        
' Outputs
    nOutputs = 31
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "leftFrontTireWheel"
    OutputTypes(iCount).description = "Left Front Tire Wheel"
    OutputTypes(iCount).Aspects = SimplePhysical
        
    iCount = 2
    OutputTypes(iCount).name = "rightFrontTireWheel"
    OutputTypes(iCount).description = "Right Front Tire Wheel"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 3
    OutputTypes(iCount).name = "leftRearTireWheel"
    OutputTypes(iCount).description = "Left Rear Tire Wheel"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).name = "rightRearTireWheel"
    OutputTypes(iCount).description = "Right Rear Tire Wheel"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 5
    OutputTypes(iCount).name = "mainBodyPart"
    OutputTypes(iCount).description = "Main Body Part"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).name = "rearBodyPart"
    OutputTypes(iCount).description = "Rear Body Part"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 7
    OutputTypes(iCount).name = "steeringColumnHousing"
    OutputTypes(iCount).description = "Steering Column Housing"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 8
    OutputTypes(iCount).name = "steeringColumnShaft"
    OutputTypes(iCount).description = "Steering Column Shaft"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 9
    OutputTypes(iCount).name = "steeringWheel"
    OutputTypes(iCount).description = "Steering Wheel"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 10
    OutputTypes(iCount).name = "seatBottom"
    OutputTypes(iCount).description = "Seat Bottom"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 11
    OutputTypes(iCount).name = "seatBack"
    OutputTypes(iCount).description = "Seat Back"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 12
    OutputTypes(iCount).name = "leftFork"
    OutputTypes(iCount).description = "Left Fork"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 13
    OutputTypes(iCount).name = "rightFork"
    OutputTypes(iCount).description = "Right Fork"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 14
    OutputTypes(iCount).name = "forkUpperCrossOver"
    OutputTypes(iCount).description = "Fork Upper Cross Over"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 15
    OutputTypes(iCount).name = "forkLeftUpright"
    OutputTypes(iCount).description = "Fork Left Upright"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 16
    OutputTypes(iCount).name = "forkRightUpright"
    OutputTypes(iCount).description = "Fork Right Upright"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 17
    OutputTypes(iCount).name = "forkLeftBumper"
    OutputTypes(iCount).description = "Fork Left Bumper"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 18
    OutputTypes(iCount).name = "forkRightBumper"
    OutputTypes(iCount).description = "Fork Right Bumper"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 19
    OutputTypes(iCount).name = "forkGrill"
    OutputTypes(iCount).description = "Fork Grill"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 20
    OutputTypes(iCount).name = "leftRearPost"
    OutputTypes(iCount).description = "Left Rear Post"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 21
    OutputTypes(iCount).name = "rightRearPost"
    OutputTypes(iCount).description = "Right Rear Post"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 22
    OutputTypes(iCount).name = "leftFrontPost"
    OutputTypes(iCount).description = "Left Front Post"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 23
    OutputTypes(iCount).name = "rightFrontPost"
    OutputTypes(iCount).description = "Right Front Post"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 24
    OutputTypes(iCount).name = "leftTopFramePiece"
    OutputTypes(iCount).description = "Left Top Frame Piece"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 25
    OutputTypes(iCount).name = "rightTopFramePiece"
    OutputTypes(iCount).description = "Right Top Frame Piece"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 26
    OutputTypes(iCount).name = "frontTopFramePiece"
    OutputTypes(iCount).description = "Front Top Frame Piece"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 27
    OutputTypes(iCount).name = "rearTopFramePiece"
    OutputTypes(iCount).description = "Rear Top Frame Piece"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 28
    OutputTypes(iCount).name = "mufflerLargePart"
    OutputTypes(iCount).description = "Muffler Large Part"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 29
    OutputTypes(iCount).name = "mufflerSmallPart"
    OutputTypes(iCount).description = "Muffler Small Part"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 30
    OutputTypes(iCount).name = "PositionPoint"
    OutputTypes(iCount).description = "Position Point"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 31
    OutputTypes(iCount).name = "DefaultSurface"
    OutputTypes(iCount).description = "Default Surface"
    OutputTypes(iCount).Aspects = SimplePhysical


' Aspects
    nAspects = 1
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    iCount = 1
    AspectTypes(iCount).name = "SimplePhysical"
    AspectTypes(iCount).description = "SimplePhysical"
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
        'Inputs(iCount).name = InputTypes(iCount).name
        'Inputs(iCount).description = InputTypes(iCount).description
        'Inputs(iCount).properties = InputTypes(iCount).properties
        'PC.uomValue = InputTypes(iCount).uomValue
' Associate a default parameter to the input
'        'Inputs(iCount).DefaultParameterValue = InputTypes(iCount).PC
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
        symbAspect.name = AspectTypes(iCount).name
        symbAspect.description = AspectTypes(iCount).description
        symbAspect.properties = AspectTypes(iCount).properties
        symbAspect.RepresentationId = AspectTypes(iCount).AspectId
'        symbAspect.RemoveAllOutput
        
        Set oSymbAspect = symbAspect
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).Aspects And AspectTypes(iCount).AspectId Then
                Output.name = OutputTypes(iOutput).name
                Output.description = OutputTypes(iOutput).description
                Output.properties = 0            ' OutputTypes(iOutput).Properties
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
        If OutputTypes(iCount).Aspects = lAspectID Then
            ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
            arrayOfOutputs(iOutputs) = OutputTypes(iCount).name
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
