VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C6FWorkerSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:42 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CSixFoorWorker.cls
'   ProgID:         SP3D6FootWorkerAsm.C6F6WorkerSym
'   Author:         nka8226
'   Creation Date:  Friday, Nov 22 2002
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)       TR 56826 Removed Msgbox
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "C6FWorkerSym:"  'Used for error messages

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
    
    m_ProjName = "SP3D6FootWorkerAsm"
    m_ClassName = "C6FWorkerSym"
    m_ProgId = m_ProjName & "." & m_ClassName

        
' Outputs
    nOutputs = 33
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "head"
    OutputTypes(iCount).description = "Head"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 2
    OutputTypes(iCount).name = "leftEye"
    OutputTypes(iCount).description = "Left Eye"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 3
    OutputTypes(iCount).name = "hat"
    OutputTypes(iCount).description = "Hat"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 4
    OutputTypes(iCount).name = "hatBrim"
    OutputTypes(iCount).description = "Hat Brim"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 5
    OutputTypes(iCount).name = "neck"
    OutputTypes(iCount).description = "Neck"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 6
    OutputTypes(iCount).name = "rightEye"
    OutputTypes(iCount).description = "Right Eye"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 7
    OutputTypes(iCount).name = "shoulders"
    OutputTypes(iCount).description = "Shoulders"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 8
    OutputTypes(iCount).name = "leftShoulderJoint"
    OutputTypes(iCount).description = "Left Shoulder Joint"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 9
    OutputTypes(iCount).name = "rightShoulderJoint"
    OutputTypes(iCount).description = "Right Shoulder Joint"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 10
    OutputTypes(iCount).name = "leftUpperArm"
    OutputTypes(iCount).description = "Left Upper Arm"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 11
    OutputTypes(iCount).name = "rightUpperArm"
    OutputTypes(iCount).description = "Right Upper Arm"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 12
    OutputTypes(iCount).name = "leftElbowJoint"
    OutputTypes(iCount).description = "Left Elbow Joint"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 13
    OutputTypes(iCount).name = "rightElbowJoint"
    OutputTypes(iCount).description = "Right Elbow Joint"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 14
    OutputTypes(iCount).name = "leftForearm"
    OutputTypes(iCount).description = "Left Forearm"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 15
    OutputTypes(iCount).name = "rightForearm"
    OutputTypes(iCount).description = "Right Forearm"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 16
    OutputTypes(iCount).name = "leftHand"
    OutputTypes(iCount).description = "Left Hand"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 17
    OutputTypes(iCount).name = "rightHand"
    OutputTypes(iCount).description = "Right Hand"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 18
    OutputTypes(iCount).name = "torsoBottom"
    OutputTypes(iCount).description = "Torso Bottom"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 19
    OutputTypes(iCount).name = "lowerMiddle"
    OutputTypes(iCount).description = "Lower Middle"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 20
    OutputTypes(iCount).name = "upperMiddle"
    OutputTypes(iCount).description = "Upper Middle"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 21
    OutputTypes(iCount).name = "torsoTop"
    OutputTypes(iCount).description = "Torso Top"
    OutputTypes(iCount).Aspects = SimplePhysical
    
    iCount = 22
    OutputTypes(iCount).name = "leftThigh"
    OutputTypes(iCount).description = "Left Thigh"
    OutputTypes(iCount).Aspects = SimplePhysical
        
    iCount = 23
    OutputTypes(iCount).name = "rightThigh"
    OutputTypes(iCount).description = "Right Thigh"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 24
    OutputTypes(iCount).name = "leftKnee"
    OutputTypes(iCount).description = "Left Knee"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 25
    OutputTypes(iCount).name = "rightKnee"
    OutputTypes(iCount).description = "Right Knee"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 26
    OutputTypes(iCount).name = "lowerLeftLeg"
    OutputTypes(iCount).description = "Lower Left Leg"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 27
    OutputTypes(iCount).name = "lowerRightLeg"
    OutputTypes(iCount).description = "Lower Right Leg"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 28
    OutputTypes(iCount).name = "leftAnkle"
    OutputTypes(iCount).description = "Left Ankle"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 29
    OutputTypes(iCount).name = "leftAnkle"
    OutputTypes(iCount).description = "Left Ankle"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 30
    OutputTypes(iCount).name = "leftFoot"
    OutputTypes(iCount).description = "Left Foot"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 31
    OutputTypes(iCount).name = "rightFoot"
    OutputTypes(iCount).description = "Right Foot"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 32
    OutputTypes(iCount).name = "controlPoint"
    OutputTypes(iCount).description = "Control Point"
    OutputTypes(iCount).Aspects = SimplePhysical

    iCount = 33
    OutputTypes(iCount).name = "defaultSurface"
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
