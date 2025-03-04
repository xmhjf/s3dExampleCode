VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiportDia5Way"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-05, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   MultiportDia5Way.cls
'   Author:         RUK
'   Creation Date:  Thursday Oct 18 2007
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 5 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'   Change History:
'   dd.mmm.yyyy     who               change description
'   -----------     ---               ------------------
'   22.Oct.2007     RUK               Created
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "MultiportDia5Way:"  'Used for error messages
'
Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String
Private m_ProjName          As String
Private m_ClassName         As String
'
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
    
    m_ProjName = "SP3DMultiportDia5Way"
    m_ClassName = "MultiportDia5Way"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 18
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).name = "ValveBodyWidth"
    InputTypes(iCount).description = "ValveBodyHeight"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1
    
    iCount = 2
    InputTypes(iCount).name = "Face1toCenter"
    InputTypes(iCount).description = "Face1 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 3
    InputTypes(iCount).name = "Face2toCenter"
    InputTypes(iCount).description = "Face2 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 4
    InputTypes(iCount).name = "Face3toCenter"
    InputTypes(iCount).description = "Face3 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 5
    InputTypes(iCount).name = "Face4toCenter"
    InputTypes(iCount).description = "Face4 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
   
    iCount = 6
    InputTypes(iCount).name = "Face5toCenter"
    InputTypes(iCount).description = "Face5 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
   
    iCount = 7
    InputTypes(iCount).name = "Face6toCenter"
    InputTypes(iCount).description = "Face6 to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 8
    InputTypes(iCount).name = "InletPortGeometry"
    InputTypes(iCount).description = "Inlet port geometry"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1

    iCount = 9
    InputTypes(iCount).name = "OutletPort1Geometry"
    InputTypes(iCount).description = "Outlet port1 geometry"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1

    iCount = 10
    InputTypes(iCount).name = "OutletPort2Geometry"
    InputTypes(iCount).description = "Outlet port2 geometry"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1

    iCount = 11
    InputTypes(iCount).name = "OutletPort3Geometry"
    InputTypes(iCount).description = "Outlet Port3 Geometry"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1
    
    iCount = 12
    InputTypes(iCount).name = "OutletPort4Geometry"
    InputTypes(iCount).description = "Outlet Port 4 Geometry"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1
    
    iCount = 13
    InputTypes(iCount).name = "OutletPort5Geometry"
    InputTypes(iCount).description = "Outlet Port 5 Geometry"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 1
    
    iCount = 14
    InputTypes(iCount).name = "ElbowEndFacetoCenter"
    InputTypes(iCount).description = "Elbow End Face to Center"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 15
    InputTypes(iCount).name = "OffsetBetOutlets"
    InputTypes(iCount).description = "Offset between Center and Outlet"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0

    iCount = 16
    InputTypes(iCount).name = "Width"
    InputTypes(iCount).description = "Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 17
    InputTypes(iCount).name = "Offset"
    InputTypes(iCount).description = "Offset"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER Or igDESCRIPTION_OPTIONAL
    InputTypes(iCount).uomValue = 0
    
    iCount = 18
    InputTypes(iCount).name = "InsulationThickness"
    InputTypes(iCount).description = "Insulation Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.025

' Outputs
    nOutputs = 31
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "ValveBody"
    OutputTypes(iCount).description = "Valve Body"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).name = "Inlet"
    OutputTypes(iCount).description = "Inlet"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 3
    OutputTypes(iCount).name = "Outlet1"
    OutputTypes(iCount).description = "Outlet 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).name = "Outlet2"
    OutputTypes(iCount).description = "Outlet 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 5
    OutputTypes(iCount).name = "Outlet3"
    OutputTypes(iCount).description = "Outlet 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).name = "Outlet4"
    OutputTypes(iCount).description = "Outlet 4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 7
    OutputTypes(iCount).name = "Outlet5"
    OutputTypes(iCount).description = "Outlet 5"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 8
    OutputTypes(iCount).name = "InletNozzle"
    OutputTypes(iCount).description = "Inlet Nozzle"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 9
    OutputTypes(iCount).name = "OutletNozzle1"
    OutputTypes(iCount).description = "Outlet Nozzle 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 10
    OutputTypes(iCount).name = "OutletNozzle2"
    OutputTypes(iCount).description = "Outlet Nozzle 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 11
    OutputTypes(iCount).name = "OutletNozzle3"
    OutputTypes(iCount).description = "Outlet Nozzle 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 12
    OutputTypes(iCount).name = "OutletNozzle4"
    OutputTypes(iCount).description = "Outlet Nozzle 4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 13
    OutputTypes(iCount).name = "OutletNozzle5"
    OutputTypes(iCount).description = "Outlet Nozzle 5"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 14
    OutputTypes(iCount).name = "ValveOperator1"
    OutputTypes(iCount).description = "Valve Operator 1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 15
    OutputTypes(iCount).name = "ValveOperator2"
    OutputTypes(iCount).description = "Valve Operator 2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 16
    OutputTypes(iCount).name = "ValveOperator3"
    OutputTypes(iCount).description = "Valve Operator 3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 17
    OutputTypes(iCount).name = "ValveOperator4"
    OutputTypes(iCount).description = "Valve Operator 4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 18
    OutputTypes(iCount).name = "ValveOperator5"
    OutputTypes(iCount).description = "Valve Operator 5"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 19
    OutputTypes(iCount).name = "InsValveBody"
    OutputTypes(iCount).description = "Insulated Valve Body"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 20
    OutputTypes(iCount).name = "InsInletPort_"
    OutputTypes(iCount).description = "Insulated Inlet"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 21
    OutputTypes(iCount).name = "InsOutletPort1"
    OutputTypes(iCount).description = "Insulated Outlet 1"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 22
    OutputTypes(iCount).name = "InsOutletPort2"
    OutputTypes(iCount).description = "Insulated Outlet 2"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 23
    OutputTypes(iCount).name = "InsOutletPort3"
    OutputTypes(iCount).description = "Insulated Outlet 3"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 24
    OutputTypes(iCount).name = "InsOutletPort4"
    OutputTypes(iCount).description = "Insulated Outlet 4"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 25
    OutputTypes(iCount).name = "InsOutletPort5"
    OutputTypes(iCount).description = "Insulated Outlet 5"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 26
    OutputTypes(iCount).name = "InsNozzle1"
    OutputTypes(iCount).description = "Insulated Inlet Nozzle"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 27
    OutputTypes(iCount).name = "InsNozzle2"
    OutputTypes(iCount).description = "Insulated Outlet Nozzle 1"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 28
    OutputTypes(iCount).name = "InsNozzle3"
    OutputTypes(iCount).description = "Insulated Outlet Nozzle 2"
    OutputTypes(iCount).Aspect = Insulation

    iCount = 29
    OutputTypes(iCount).name = "InsNozzle4"
    OutputTypes(iCount).description = "Insulated Outlet Nozzle 4"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 30
    OutputTypes(iCount).name = "InsNozzle5"
    OutputTypes(iCount).description = "Insulated Outlet Nozzle 5"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 31
    OutputTypes(iCount).name = "InsNozzle6"
    OutputTypes(iCount).description = "Insulated Outlet Nozzle 6"
    OutputTypes(iCount).Aspect = Insulation
    
    'Aspects
    nAspects = 2
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    iCount = 1
    AspectTypes(iCount).name = "SimplePhysical"
    AspectTypes(iCount).description = "SimplePhysical"
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
        symbAspect.properties = AspectTypes(iCount).properties
'        symbAspect.properties = igCOLLECTION_VARIABLE
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
                    ByVal OutputColl As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    Dim arrayOfOutputs()    As String
    Dim iOutputs            As Integer
    Dim iCount              As Integer
    Dim lAspectID           As SymbolRepIds
    Dim AspectName          As String
    Dim bsucceded           As Boolean
    
    Set m_OutputColl = OutputColl
    
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
        objAspect.run OutputColl, arrayOfInputs, arrayOfOutputs
        Set objAspect = Nothing
    End If
    
    If StrComp(repName, "Insulation", vbTextCompare) = 0 Then
        Set objAspect = New CInsulation
        objAspect.run OutputColl, arrayOfInputs, arrayOfOutputs
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































'''Private Sub Class_Initialize()
'''
'''    Const METHOD = "Class_Initialize:"
'''    On Error GoTo ErrorHandler
'''
'''    Set m_oSymbolHelper = New SymbolServices
'''    m_oSymbolHelper.ProjectName = "SP3DMultiportDia5Way"
'''    m_oSymbolHelper.ClassName = "MultiportDia5Way"
'''
'''' Inputs
'''    m_oSymbolHelper.NumInputs = 10
'''    m_oSymbolHelper.AddInputDef 1, "ValveBodyWidth", "Valve Body Width", igValue
'''    m_oSymbolHelper.AddInputDef 2, "ValveBodyDepth", "Valve Body Depth", igValue
'''    m_oSymbolHelper.AddInputDef 3, "ValveBodyHeight", "Valve Body Height", igValue
'''    m_oSymbolHelper.AddInputDef 4, "Face1toCenter", "Face 1 to Center", igValue
'''    m_oSymbolHelper.AddInputDef 5, "Face2toCenter", "Face 2 to Center", igValue
'''    m_oSymbolHelper.AddInputDef 6, "Face3toCenter", "Face 3 to Center", igValue
'''    m_oSymbolHelper.AddInputDef 7, "Face4toCenter", "Face 4 to Center", igValue
'''    m_oSymbolHelper.AddInputDef 8, "Face5toCenter", "Face 5 to Center", igValue
'''    m_oSymbolHelper.AddInputDef 9, "Offset", "Offset", igValue
'''    m_oSymbolHelper.AddInputDef 10, "InsulationThickness", "Insulation Thickness", 0.025
'''
'''' Outputs
'''    m_oSymbolHelper.NumOutputs = 32
'''    m_oSymbolHelper.AddOutputDef 1, "ValveBody", "Valve Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 2, "Nozzle2CylinderBody", "Nozzle 2  Cylinder Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 3, "Nozzle2CurvedBody", "Nozzle 2  Curved Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 4, "Nozzle3CylinderBody", "Nozzle 3  Cylinder Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 5, "Nozzle3CurvedBody", "Nozzle 3  Curved Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 6, "Nozzle4CylinderBody", "Nozzle 4  Cylinder Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 7, "Nozzle4CurvedBody", "Nozzle 4  Curved Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 8, "Nozzle5CylinderBody", "Nozzle 5  Cylinder Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 9, "Nozzle5CurvedBody", "Nozzle 5  Curved Body", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 10, "Nozzle1", "Nozzle 1 with length", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 11, "Nozzle2", "Nozzle 2 with length", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 12, "Nozzle3", "Nozzle 3 with length", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 13, "Nozzle4", "Nozzle 4 with length", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 14, "Nozzle5", "Nozzle 5 with length", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 15, "ValveOperator1", "Port 1 Valve Operator", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 16, "ValveOperator2", "Port 2 Valve Operator", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 17, "ValveOperator3", "Port 3 Valve Operator", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 18, "ValveOperator4", "Port 4 Valve Operator", SimplePhysical
'''    m_oSymbolHelper.AddOutputDef 19, "InsValveBody", "Ins Valve Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 20, "InsNozzle2CylinderBody", "Ins Nozzle 2  Cylinder Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 21, "InsNozzle2CurvedBody", "Ins Nozzle 2  Curved Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 22, "InsNozzle3CylinderBody", "Ins Nozzle 3  Cylinder Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 23, "InsNozzle3CurvedBody", "Ins Nozzle 3  Curved Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 24, "InsNozzle4CylinderBody", "Ins Nozzle 4  Cylinder Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 25, "InsNozzle4CurvedBody", "Ins Nozzle 4  Curved Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 26, "InsNozzle5CylinderBody", "Ins Nozzle 5  Cylinder Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 27, "InsNozzle5CurvedBody", "Ins Nozzle 5  Curved Body", Insulation
'''    m_oSymbolHelper.AddOutputDef 28, "InsNozzle1", "Ins Nozzle 1", Insulation
'''    m_oSymbolHelper.AddOutputDef 29, "InsNozzle2", "Ins Nozzle 2", Insulation
'''    m_oSymbolHelper.AddOutputDef 30, "InsNozzle3", "Ins Nozzle 3", Insulation
'''    m_oSymbolHelper.AddOutputDef 31, "InsNozzle4", "Ins Nozzle 4", Insulation
'''    m_oSymbolHelper.AddOutputDef 32, "InsNozzle5", "Ins Nozzle 5", Insulation
'''
'''' Aspects
'''    m_oSymbolHelper.NumAspects = 2
'''    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
'''    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
'''
'''    Exit Sub
'''
'''ErrorHandler:
'''  Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
'''       Err.HelpFile, Err.HelpContext
'''
'''End Sub
'''
'''Private Sub Class_Terminate()
'''
'''    Set m_oSymbolHelper = Nothing
'''
'''End Sub
'''
'''' note that the rest of the code is generated by the Symbol Wizard and should not need to be modified
'''Public Function IJDUserSymbolServices_InstanciateDefinition( _
'''                ByVal CodeBase As String, _
'''                ByVal defParameters As Variant, _
'''                ByVal ActiveConnection As Object) As Object
'''    ' call symbol services default implementation of this method
'''    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolHelper.InstanciateDefinition(CodeBase, defParameters, ActiveConnection)
'''End Function
'''Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
'''
'''  'IJDUserSymbolServices_GetDefinitionName = m_ProgId
'''    IJDUserSymbolServices_GetDefinitionName = m_oSymbolHelper.ProjectName + "." + m_oSymbolHelper.ClassName
'''
'''End Function
'''
'''Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
'''
'''
'''    ' call symbol services default implementation of this method passing in the array of input types,
'''    ' AspectTypes and OutputTypes which are defined in the class initialize method of the symbol
'''    m_oSymbolHelper.InitializeSymbolDefinition pSymbolDefinition
'''
'''End Sub
'''
'''Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
'''                    ByVal repName As String, _
'''                    ByVal outputcoll As Object, _
'''                    ByRef arrayOfInputs())
'''
'''    m_oSymbolHelper.InvokeRepresentation sblOcc, repName, outputcoll, arrayOfInputs
'''
'''End Sub
'''
'''Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
'''
'''    ' The definition uses the generic EditOccurrence command
'''    IJDUserSymbolServices_EditOccurence = False
'''
'''End Function
