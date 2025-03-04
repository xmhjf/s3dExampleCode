VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LLWithBackingPlSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:27 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   ProgID:         SP3DLiftLugBPCompAsm.LLWithBackingPlSym
'   Author:         svsmylav
'   Creation Date:  Tuesday, Feb 25 2003
'   Description:
'    Support Lug Detailed Symbol
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "LLWithBackingPlSym:"  'Used for error messages

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
    
    m_ProjName = "SP3DLiftLugBPCompAsm"
    m_ClassName = "LLWithBackingPlSym"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 11
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).Name = "VesselDiameter"
    InputTypes(iCount).Description = "Vessel Shell Outside Diameter"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1.524
    
    iCount = 2
    InputTypes(iCount).Name = "VesselLugOffset"
    InputTypes(iCount).Description = "Vessel Support Lug ControlPoint Offset"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.01
    
    iCount = 3
    InputTypes(iCount).Name = "LugBasePlateWidth"
    InputTypes(iCount).Description = "Vessel Support Lug Base Plate Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.15
    
    iCount = 4
    InputTypes(iCount).Name = "LugBasePlateLength"
    InputTypes(iCount).Description = "Vessel Support Lug Base Plate Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.15
    
    iCount = 5
    InputTypes(iCount).Name = "LugBasePlateThickness"
    InputTypes(iCount).Description = "Vessel Support Lug Base Plate Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.01
    
    iCount = 6
    InputTypes(iCount).Name = "LugGussetHeight"
    InputTypes(iCount).Description = "Vessel Support Lug Gusset Height"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 7
    InputTypes(iCount).Name = "LugGussetWidth"
    InputTypes(iCount).Description = "Vessel Support Lug Gusset Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.14
    
    iCount = 8
    InputTypes(iCount).Name = "LugGussetThickness"
    InputTypes(iCount).Description = "Vessel Support Lug Gusset Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.01
    
    iCount = 9
    InputTypes(iCount).Name = "LugGussetSpacing"
    InputTypes(iCount).Description = "Vessel Support Lug Gusset Spacing"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.12
    
    iCount = 10
    InputTypes(iCount).Name = "LugBoltSlotEccentricity"
    InputTypes(iCount).Description = "Vessel Support Lug Bolt Slot Eccentricity"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.115
    
    iCount = 11
    InputTypes(iCount).Name = "LugBoltDiameter"
    InputTypes(iCount).Description = "Vessel Support Lug Bolt Diameter"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.02
    
' Outputs
    nOutputs = 34
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).Name = "Lug1_PlaneOuput1"
    OutputTypes(iCount).Description = "Lug1_PlaneOuput 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).Name = "Lug1_PlaneOuput2"
    OutputTypes(iCount).Description = "Lug1_PlaneOuput 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 3
    OutputTypes(iCount).Name = "Lug1_PlaneOuput3"
    OutputTypes(iCount).Description = "Lug1_PlaneOuput 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).Name = "Lug1_PlaneOuput4"
    OutputTypes(iCount).Description = "Lug1_PlaneOuput 4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 5
    OutputTypes(iCount).Name = "Box1"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 6
    OutputTypes(iCount).Name = "Box2"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 7
    OutputTypes(iCount).Name = "Pyra1"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 8
    OutputTypes(iCount).Name = "Pyra2"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 9
    OutputTypes(iCount).Name = "Lug2_PlaneOuput1"
    OutputTypes(iCount).Description = "Lug2_PlaneOuput 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 10
    OutputTypes(iCount).Name = "Lug2_PlaneOuput2"
    OutputTypes(iCount).Description = "Lug2_PlaneOuput 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 11
    OutputTypes(iCount).Name = "Lug2_PlaneOuput3"
    OutputTypes(iCount).Description = "Lug2_PlaneOuput 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 12
    OutputTypes(iCount).Name = "Lug2_PlaneOuput4"
    OutputTypes(iCount).Description = "Lug2_PlaneOuput 4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 13
    OutputTypes(iCount).Name = "Box3"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 14
    OutputTypes(iCount).Name = "Box4"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 15
    OutputTypes(iCount).Name = "Pyra3"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 16
    OutputTypes(iCount).Name = "Pyra4"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 17
    OutputTypes(iCount).Name = "Lug3_PlaneOuput1"
    OutputTypes(iCount).Description = "Lug3_PlaneOuput 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 18
    OutputTypes(iCount).Name = "Lug3_PlaneOuput2"
    OutputTypes(iCount).Description = "Lug3_PlaneOuput 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 19
    OutputTypes(iCount).Name = "Lug3_PlaneOuput3"
    OutputTypes(iCount).Description = "Lug3_PlaneOuput 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 20
    OutputTypes(iCount).Name = "Lug3_PlaneOuput4"
    OutputTypes(iCount).Description = "Lug3_PlaneOuput 4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 21
    OutputTypes(iCount).Name = "Box5"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 22
    OutputTypes(iCount).Name = "Box6"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 23
    OutputTypes(iCount).Name = "Pyra5"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 24
    OutputTypes(iCount).Name = "Pyra6"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 25
    OutputTypes(iCount).Name = "Lug4_PlaneOuput1"
    OutputTypes(iCount).Description = "Lug4_PlaneOuput 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 26
    OutputTypes(iCount).Name = "Lug4_PlaneOuput2"
    OutputTypes(iCount).Description = "Lug4_PlaneOuput 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 27
    OutputTypes(iCount).Name = "Lug4_PlaneOuput3"
    OutputTypes(iCount).Description = "Lug4_PlaneOuput 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 28
    OutputTypes(iCount).Name = "Lug4_PlaneOuput4"
    OutputTypes(iCount).Description = "Lug4_PlaneOuput 4"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 29
    OutputTypes(iCount).Name = "Box7"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 30
    OutputTypes(iCount).Name = "Box8"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 31
    OutputTypes(iCount).Name = "Pyra7"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 32
    OutputTypes(iCount).Name = "Pyra8"
    OutputTypes(iCount).Description = "Support Lug"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 33
    OutputTypes(iCount).Name = "Point"
    OutputTypes(iCount).Description = "Drum Point"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 34
    OutputTypes(iCount).Name = "Plane"
    OutputTypes(iCount).Description = "Default Plane"
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
        Inputs(iCount).Name = InputTypes(iCount).Name
        Inputs(iCount).Description = InputTypes(iCount).Description
        Inputs(iCount).Properties = InputTypes(iCount).Properties
        PC.uomValue = InputTypes(iCount).uomValue
 'Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
 '---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
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
        'symbAspect.Properties = AspectTypes(iCount).Properties
        symbAspect.Properties = igCOLLECTION_VARIABLE
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
