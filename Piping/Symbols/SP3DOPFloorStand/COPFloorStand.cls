VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COPFloorStand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2009, Intergraph Corporation. All rights reserved.
'
'   COPFloorStand.cls.cls
'   ProgID:         SP3DOPFloorStand.COPFloorStand
'   Author:       ACM
'   Creation Date:  Thursday, Dec 19 2002
'   Description:
'   Operator Floor stand symbol is prepared based on SP3D Concept design document
''  "Valve Accessories"(valveaccessoriescd.doc).
''  This symbol consists of 4 inputs and 4 outputs all are physical.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)      TR 56826 Removed Msgbox
'   18.Jun.2004     MS    (RI 32889)  Changed input parameter name from
'                                           HandWheelDia to HandWheelDiameter(as an existing interface can be used instead of creating a new interface for HandWheelDia)
'   08.22.2004      MS     CR 57807 Added Maintenace Envelope to the operator
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   21.APR.2009         PK     TR-163529    CreateObject() has been replaced with 'New' representation
'                                        class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "COPFloorStand:"  'Used for error messages

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
    
    Dim icount As Integer
    
    m_ProjName = "SP3DOPFloorStand"
    m_ClassName = "COPFloorStand"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 4
    ReDim InputTypes(1 To nInputs) As InputType
    
    icount = 1
    InputTypes(icount).name = "OperatorHeight"
    InputTypes(icount).description = "Height of the Operator form valve centre to HW top"
    InputTypes(icount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(icount).uomValue = 0.6
    
    icount = 2
    InputTypes(icount).name = "OperatorHeight1"
    InputTypes(icount).description = "Height of operator from bottom of base plate to HW top"
    InputTypes(icount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(icount).uomValue = 0.3
    
    icount = 3
    InputTypes(icount).name = "HandWheelDiameter"
    InputTypes(icount).description = "Diameter of Hand Wheel"
    InputTypes(icount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(icount).uomValue = 0.05
    
    icount = 4
    InputTypes(icount).name = "BasePlateDia"
    InputTypes(icount).description = "Diameter of base Plate"
    InputTypes(icount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(icount).uomValue = 0.08
    
        
' Outputs
    nOutputs = 5
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    icount = 1
    OutputTypes(icount).name = "HandWheelTorus"
    OutputTypes(icount).description = "Hand Wheel Body Torus"
    OutputTypes(icount).Aspect = SimplePhysical
    
    icount = 2
    OutputTypes(icount).name = "SpokeCyl1"
    OutputTypes(icount).description = "Spoke Cylinder 1"
    OutputTypes(icount).Aspect = SimplePhysical
    
    icount = 3
    OutputTypes(icount).name = "SpokeCyl2"
    OutputTypes(icount).description = "Spoke Cylinder 2"
    OutputTypes(icount).Aspect = SimplePhysical
    
    icount = 4
    OutputTypes(icount).name = "BodyRevolution"
    OutputTypes(icount).description = "Body by Revolution of complex str"
    OutputTypes(icount).Aspect = SimplePhysical
    
    icount = 5
    OutputTypes(icount).name = "MaintCyl"
    OutputTypes(icount).description = "Maintenance Cylinder"
    OutputTypes(icount).Aspect = Maintenance
    
' Aspects
    nAspects = 2
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    icount = 1
    AspectTypes(icount).name = "Physical"
    AspectTypes(icount).description = "Physical"
    AspectTypes(icount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(icount).AspectId = SimplePhysical
    
    icount = 2
    AspectTypes(icount).name = "Maintenance"
    AspectTypes(icount).description = "Maintenance"
    AspectTypes(icount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(icount).AspectId = Maintenance
    
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
    Dim Inputs()    As IMSSymbolEntities.IJDInput
    ReDim Inputs(1 To nInputs) As IMSSymbolEntities.IJDInput
    Dim icount      As Integer
' Create a default parameter
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    PC.Type = igValue
    For icount = 1 To nInputs
        Set Inputs(icount) = New IMSSymbolEntities.DInput
        Inputs(icount).name = InputTypes(icount).name
        Inputs(icount).description = InputTypes(icount).description
        Inputs(icount).properties = InputTypes(icount).properties
        PC.uomValue = InputTypes(icount).uomValue
' Associate a default parameter to the input
        Inputs(icount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(icount), icount + 1
        Set Inputs(icount) = Nothing
    Next icount
        
' Define output
    Dim symbAspect As IMSSymbolEntities.IJDRepresentation
    Set symbAspect = New IMSSymbolEntities.DRepresentation
' ---------Set the Aspect to definition
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations
    Set RepsIf = pSymbolDefinition
    Dim Output As IMSSymbolEntities.IJDOutput
    Dim iOutput      As Integer
    Dim oSymbAspect As IMSSymbolEntities.IJDOutputs
    
    For icount = 1 To nAspects
        symbAspect.name = AspectTypes(icount).name
        symbAspect.description = AspectTypes(icount).description
        symbAspect.properties = AspectTypes(icount).properties
        symbAspect.RepresentationId = AspectTypes(icount).AspectId
'        symbAspect.RemoveAllOutput
        
        Set oSymbAspect = symbAspect
        Set Output = New IMSSymbolEntities.DOutput
        For iOutput = 1 To nOutputs
            If OutputTypes(iOutput).Aspect And AspectTypes(icount).AspectId Then
                Output.name = OutputTypes(iOutput).name
                Output.description = OutputTypes(iOutput).description
                Output.properties = 0            ' OutputTypes(iOutput).Properties
                oSymbAspect.SetOutput Output
            End If
        Next iOutput
        RepsIf.SetRepresentation symbAspect
        Set Output = Nothing
    Next icount
    

    ' ----------Definition of the function identification corresponding
    Dim oRepEval     As IJDRepresentationEvaluation
    Dim RepsEvalsIf     As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition
    
    For icount = 1 To nAspects
        Set oRepEval = New DRepresentationEvaluation
        oRepEval.name = AspectTypes(icount).name
        oRepEval.description = AspectTypes(icount).description
        oRepEval.properties = igREPRESENTATION_HIDDEN
        oRepEval.Type = igREPRESENTATION_VBFUNCTION
        oRepEval.ProgId = m_ProgId
        RepsEvalsIf.AddRepresentationEvaluation oRepEval
        Set oRepEval = Nothing
    Next icount
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
    Dim icount              As Integer
    Dim lAspectID           As SymbolRepIds
    Dim AspectName          As String
    Dim bsucceded           As Boolean
    
    Set m_OutputColl = outputcoll
    
    bsucceded = False
    For icount = 1 To nAspects
        If AspectTypes(icount).name = repName Then
            bsucceded = True
            lAspectID = AspectTypes(icount).AspectId
            Exit For
        End If
    Next icount
    If Not bsucceded Then
'        MsgBox "Unknown Aspect; Check Aspect data / code"
        Exit Sub
    End If
    
    iOutputs = 1
    For icount = 1 To nOutputs
        If OutputTypes(icount).Aspect = lAspectID Then
            ReDim Preserve arrayOfOutputs(1 To iOutputs) As String
            arrayOfOutputs(iOutputs) = OutputTypes(icount).name
            iOutputs = iOutputs + 1
        End If
    Next icount
    Dim objAspect As Object
    
    If StrComp(repName, "Physical", vbTextCompare) = 0 Then
        Set objAspect = New CPhysical
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName, "Maintenance", vbTextCompare) = 0 Then
        Set objAspect = New CMaintenance
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
