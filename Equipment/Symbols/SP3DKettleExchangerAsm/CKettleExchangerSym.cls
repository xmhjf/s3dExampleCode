VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKXSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:03 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'
'   CKettleExchangerSym.cls
'   ProgID:         SP3DKettleXAsm.CKXSym
'   Author:         BG
'   Creation Date:  Thursday, Apr 18 2002
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   Jan 27 2003     SSP                  Create Control Point Output  at the origin.
'   Apr 28 2003     BG/svsmylav             TR# 39291 Removed space character in input parameters definition.
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)       TR 56826 Removed Msgbox
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related inputs and outputs.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   14.Mar.2007     svsmylav          TR-83678 InitializeSymbolDefinition method is modified to eliminate unnecessary output definitions.
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CKXSym:"  'Used for error messages

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
    
    m_ProjName = "SP3DKettleXAsm"
    m_ClassName = "CKXSym"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 22
    ReDim InputTypes(1 To nInputs) As InputType
    
    iCount = 1
    InputTypes(iCount).name = "ExchangerLength"
    InputTypes(iCount).description = "Exchanger Length P1"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2
    
    iCount = 2
    InputTypes(iCount).name = "ExchangerNeckLength"
    InputTypes(iCount).description = "Exchanger Neck Length P2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.65
    
    iCount = 3
    InputTypes(iCount).name = "ExchangerNeckDiameter"
    InputTypes(iCount).description = "Exchanger Neck Diameter P3"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.3
    
    iCount = 4
    InputTypes(iCount).name = "ExchangerTaperLength"
    InputTypes(iCount).description = "Exchanger Taper Length P4"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.4
    
    iCount = 5
    InputTypes(iCount).name = "ExchangerDiameter"
    InputTypes(iCount).description = "Exchanger Diameter P5"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.5
    
    iCount = 6
    InputTypes(iCount).name = "BundleFlangeTk"
    InputTypes(iCount).description = "Bundle Flange Thickness P6"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05
    
    iCount = 7
    InputTypes(iCount).name = "BundleFlangeDia"
    InputTypes(iCount).description = "Bundle Flange Diameter P7"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.55
    
    iCount = 8
    InputTypes(iCount).name = "ExchangerFlangeTk1"
    InputTypes(iCount).description = "Exchanger Flange Thickness P8"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.065
    
    iCount = 9
    InputTypes(iCount).name = "BundlePullingLength"
    InputTypes(iCount).description = "Bundle Pulling Length P9"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2
    
    iCount = 10
    InputTypes(iCount).name = "BotSupportCenFromPP"
    InputTypes(iCount).description = "Bot Support Center FromPP P10"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1.4
    
    iCount = 11
    InputTypes(iCount).name = "BottomSupportCentoCen"
    InputTypes(iCount).description = "Bottom SupportCentertoCenterP11"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.7
    
    iCount = 12
    InputTypes(iCount).name = "Support1Thickness"
    InputTypes(iCount).description = "Support 1 Thickness P12"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.25
    
    iCount = 13
    InputTypes(iCount).name = "Support2Thickness"
    InputTypes(iCount).description = "Support 2 Thickness P13"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 14
    InputTypes(iCount).name = "BottomSupportHeight"
    InputTypes(iCount).description = "BottomSupportHeight P14"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.5
    
    iCount = 15
    InputTypes(iCount).name = "SupportLength"
    InputTypes(iCount).description = "SupportLength P15"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.4
    
    iCount = 16
    InputTypes(iCount).name = "FrontEndFlangeDia"
    InputTypes(iCount).description = "FrontEndFlangeDia P30"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.65
    
    iCount = 17
    InputTypes(iCount).name = "FrontEndFlangeTk1"
    InputTypes(iCount).description = "FrontEndFlangeTk1 P31"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.065
    
    iCount = 18
    InputTypes(iCount).name = "FrontEndLength1"
    InputTypes(iCount).description = "FrontEndLength1  P32"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.5
    
    iCount = 19
    InputTypes(iCount).name = "FrontEndLength2"
    InputTypes(iCount).description = "FrontEndLength2  P33"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 20
    InputTypes(iCount).name = "FrontEndFlangeTk2"
    InputTypes(iCount).description = "FrontEndFlangeTk2 P34"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.05
    
    iCount = 21
    InputTypes(iCount).name = "FrontEndFlangeTk3"
    InputTypes(iCount).description = "FrontEndFlangeTk3 P35"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.045
    
    iCount = 22
    InputTypes(iCount).name = "InsulationThickness"
    InputTypes(iCount).description = "Insulation Thickness"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.06
        
' Outputs
    nOutputs = 10
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).name = "ExchangerBody"
    OutputTypes(iCount).description = "ExchangerBody"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).name = "ExTaperBody"
    OutputTypes(iCount).description = "ExchangerTaperBody"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 3
    OutputTypes(iCount).name = "ExneckBody"
    OutputTypes(iCount).description = "Exchanger Neck Portion"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).name = "ExFrontEnd"
    OutputTypes(iCount).description = "Exchanger Front End"
    OutputTypes(iCount).Aspect = SimplePhysical
        
    iCount = 5
    OutputTypes(iCount).name = "ExBodyInsul"
    OutputTypes(iCount).description = "ExchangerBody Insulation"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 6
    OutputTypes(iCount).name = "ExTapBodyInsul"
    OutputTypes(iCount).description = "ExchangerTaperBody ins"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 7
    OutputTypes(iCount).name = "ExneckBodyInsul"
    OutputTypes(iCount).description = "Exchanger Neck Portion ins"
    OutputTypes(iCount).Aspect = Insulation
    
    iCount = 8
    OutputTypes(iCount).name = "BundlePullCylin"
    OutputTypes(iCount).description = "Bundle Pulling Cylinder"
    OutputTypes(iCount).Aspect = Maintenance
    
    iCount = 9
    OutputTypes(iCount).name = "ExFrontEndIns"
    OutputTypes(iCount).description = "Ex front end Ins"
    OutputTypes(iCount).Aspect = Insulation
            
    iCount = 10
    OutputTypes(iCount).name = "KettleExchangerControlPoint"
    OutputTypes(iCount).description = "Control Point of Kettle Exchange"
    OutputTypes(iCount).Aspect = SimplePhysical
        
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
    AspectTypes(iCount).name = "Maintenance"
    AspectTypes(iCount).description = "Maintenance"
    AspectTypes(iCount).properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = Maintenance
    
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
    
    If StrComp(repName,"Maintenance", vbTextCompare) = 0 Then
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
