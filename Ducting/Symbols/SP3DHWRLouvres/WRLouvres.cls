VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WRLouvres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright(c)2005-08 Intergraph GmbH
'   All Rights Reserved
'
'   WRLouvres.cls
'   ProgID:         SP3DHWRLouvres.WRLouvres
'   Author:         GW
'   Creation Date:  Montag, Sep 14 2005
'   Description:
'   The following Part data basis cases are addressed for the parameters specified:
'   Case A (Part data Basis value -105): Fixed Blade Weather Louver, Type A
'   Case B (Part data Basis value -110): Fixed Blade Weather Louver, Type B
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   Modification 02-19-03 to check for Symbol creation using SP3DSymbolHelper.SymbolServices
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   08.AUG.2007     MA   CR-124221 Implemented part data basis for values 105 and 110.
'   21.APR.2008     PK   DI-132538 CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'   29.AUG.2008     VKC  CR-149323 Added 3 input parameters Frame Width ,Frame Depth and Blade Angle.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "WRLouvres:"  'Used for error messages

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
Private PI                  As Double
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProjName = "SP3DHWRLouvres"
    m_ClassName = "WRLouvres"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 11
    ReDim InputTypes(1 To nInputs) As InputType

    iCount = 1
    InputTypes(iCount).name = "Width"
    InputTypes(iCount).description = "Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1

    iCount = 2
    InputTypes(iCount).name = "Depth"
    InputTypes(iCount).description = "Depth"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2

    iCount = 3
    InputTypes(iCount).name = "LouvreLength"
    InputTypes(iCount).description = "Louvre Length"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0

    iCount = 4
    InputTypes(iCount).name = "LouvreDistance"
    InputTypes(iCount).description = "Louvre Distance"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 5
    InputTypes(iCount).name = "LouvreWidth"
    InputTypes(iCount).description = "Louvre Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 6
    InputTypes(iCount).name = "FlangeThickness"
    InputTypes(iCount).description = "Flange Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0.005

    iCount = 7
    InputTypes(iCount).name = "Offset1"
    InputTypes(iCount).description = "Doubler Plate Offset 1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 8
    InputTypes(iCount).name = "Offset2"
    InputTypes(iCount).description = "Doubler Plate Offset 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
     iCount = 9
    InputTypes(iCount).name = "FrameWidth"
    InputTypes(iCount).description = "Frame Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
     iCount = 10
    InputTypes(iCount).name = "FrameDepth"
    InputTypes(iCount).description = "Frame Depth"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
     iCount = 11
    InputTypes(iCount).name = "Angle"
    InputTypes(iCount).description = "Blade Angle"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = (5 * PI) / 24
    
' Outputs
    ReDim OutputTypes(1 To lngNMaxOutputs + 1) As OutputType
    nOutputs = lngNMaxOutputs + 1
    OutputTypes(1).name = "HNoz1"
    OutputTypes(1).description = "HNoz1"
    OutputTypes(1).Aspect = SimplePhysical
    
    Dim i As Long
    For i = 1 To lngNMaxOutputs
    OutputTypes(i + 1).name = "Feature" & i
    OutputTypes(i + 1).description = "Feature" & i
    OutputTypes(i + 1).Aspect = SimplePhysical
    Next i
    

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
        symbAspect.properties = igCOLLECTION_VARIABLE
'        symbAspect.properties = AspectTypes(iCount).properties
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
