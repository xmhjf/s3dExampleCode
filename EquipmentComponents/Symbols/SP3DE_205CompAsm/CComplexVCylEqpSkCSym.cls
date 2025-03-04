VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CE_205CSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:59 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 -07, Intergraph Corporation. All rights reserved.
'
'   CComplexVCylEqpSkCSym.cls
'   ProgID:         SP3DE_205CompAsm.CE_205CSym
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This is Complex Vertical Cylindrical Equipment Skirt Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E205 Symbol in Page no 286.
'   Symbol is created using the following Outputs:
'   i)  4 standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: Vessel uses 'PlaceRevolution'
'       c) Two ReferenceGeometry aspect outputs: a Default Surface and a Control point
'   ii) Variable Outputs:
'        a) Support
'        b) Surface for the support and
'        c) Intermediate dome for shell section 3
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'   14.Mar.2007     svsmylav          TR-83678 InitializeSymbolDefinition method is modified to eliminate unnecessary output definitions.
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "CE_205CSym:"  'Used for error messages

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

    m_ProjName = "SP3DE_205CompAsm"
    m_ClassName = "CE_205CSym"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 16
    ReDim InputTypes(1 To nInputs) As InputType

    iCount = 1
    InputTypes(iCount).name = "VesselHeight1"
    InputTypes(iCount).description = "Vessel Height 1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2

    iCount = 2
    InputTypes(iCount).name = "VesselDiameter1"
    InputTypes(iCount).description = "Vessel Diameter 1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 3

    iCount = 3
    InputTypes(iCount).name = "SupportHeight1"
    InputTypes(iCount).description = "Support Height 1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 4

    iCount = 4
    InputTypes(iCount).name = "VesselHeight2"
    InputTypes(iCount).description = "Vessel Height 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 5

    iCount = 5
    InputTypes(iCount).name = "VesselDiameter2"
    InputTypes(iCount).description = "Vessel Diameter 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 6

    iCount = 6
    InputTypes(iCount).name = "SupportHeight2"
    InputTypes(iCount).description = "Support Height 2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 7

    iCount = 7
    InputTypes(iCount).name = "VesselHeight3"
    InputTypes(iCount).description = "Vessel Height 3"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 8

    iCount = 8
    InputTypes(iCount).name = "VesselDiameter3"
    InputTypes(iCount).description = "Vessel Diameter 3"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 9

    iCount = 9
    InputTypes(iCount).name = "SupportHeight3"
    InputTypes(iCount).description = "Support Height 3"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 3.2

    iCount = 10
    InputTypes(iCount).name = "VesselHeight4"
    InputTypes(iCount).description = "Vessel Height 4"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 5.5

    iCount = 11
    InputTypes(iCount).name = "VesselDiameter4"
    InputTypes(iCount).description = "Vessel Diameter 4"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 5.4

    iCount = 12
    InputTypes(iCount).name = "SkirtHeight"
    InputTypes(iCount).description = "Skirt Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 8

    iCount = 13
    InputTypes(iCount).name = "SkirtTopDiameter"
    InputTypes(iCount).description = "Skirt Top Diameter"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 5.4

    iCount = 14
    InputTypes(iCount).name = "SkirtBottomDiameter"
    InputTypes(iCount).description = "Skirt Bottom Diameter"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 7

    iCount = 15
    InputTypes(iCount).name = "VesselReferencePoint"
    InputTypes(iCount).description = "Vessel Reference Point"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 16.6

    iCount = 16
    InputTypes(iCount).name = "InsulationThickness"
    InputTypes(iCount).description = "Insulation Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025


' Outputs
    nOutputs = 4
    ReDim OutputTypes(1 To nOutputs) As OutputType

    iCount = 1
    OutputTypes(iCount).name = "InsulatedVessel"
    OutputTypes(iCount).description = "InsulatedVessel"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 2
    OutputTypes(iCount).name = "Vessel"
    OutputTypes(iCount).description = "Vessel"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 3
    OutputTypes(iCount).name = "DefaultSurface"
    OutputTypes(iCount).description = "Default Surface of Complex Vertical Cylindrical Equipment Skirt Component"
    OutputTypes(iCount).Aspect = ReferenceGeometry

    iCount = 4
    OutputTypes(iCount).name = "ComplexVCylEqpSkCompControlPoint"
    OutputTypes(iCount).description = "Control Point of Complex Vertical Cylindrical Equipment Skirt Component"
    OutputTypes(iCount).Aspect = ReferenceGeometry

' Aspects
    nAspects = 3
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
    iCount = 3
    AspectTypes(iCount).name = "ReferenceGeometry"
    AspectTypes(iCount).description = "ReferenceGeometry"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = ReferenceGeometry

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
    If StrComp(repName,"Physical", vbTextCompare) = 0 Then
        Set objAspect = New CPhysical
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName,"Insulation", vbTextCompare) = 0 Then
        Set objAspect = New CInsulation
        objAspect.run outputcoll, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName,"ReferenceGeometry", vbTextCompare) = 0 Then
        Set objAspect = New CReferenceGeometry
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
