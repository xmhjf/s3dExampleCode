VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Doors_1_Sym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:10 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'
'   CG7C_2011_17Sym.cls
'   ProgID:         SP3DG7C_2011_17Asm.CG7C_2011_17Sym
'   Author:         Cynthia Wagner
'   Creation Date:  Thursday, Feb 27 2003
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   May 18, 2003    Nabil Nicolas           Define functions classes in Modules folder
'                                           Define SimplePhysical aspect before Maintenance aspect.
'                                           Declare dblXLength, dblYLength, dblZLength as double instead of as string
'                                           Replace argument oriOrientation by Nothing, Set oriOrientation = Nothing
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   19.Mar.2004     SymbolTeam(India)       TR 56826 Removed Msgbox
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   05.Apr.2006     Regis Dechance          Adaptation for Doors, Windows, Louvers
'   14.Mar.2008     SS                      DI#134831 - Changed the code from CreateObject() to NEW
'                                           as the symbol is no longer registered.
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Doors_1_Sym:"  'Used for error messages

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
    
    m_ProjName = "SP3DDoorsAsm"
    m_ClassName = "Doors_1_Sym"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    
    nInputs = 26
    ReDim InputTypes(1 To nInputs) As InputType
    iCount = 1
    InputTypes(iCount).Name = "DoorType"
    InputTypes(iCount).Description = "Type of Door"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 2
    InputTypes(iCount).Name = "OpeningRatio"
    InputTypes(iCount).Description = "Opening rate of Door"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0#
    
    iCount = 3
    InputTypes(iCount).Name = "Kinematics"
    InputTypes(iCount).Description = "Kinematics"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 4
    InputTypes(iCount).Name = "Push"
    InputTypes(iCount).Description = "Push"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1
    
    iCount = 5
    InputTypes(iCount).Name = "Height"
    InputTypes(iCount).Description = "Height of Door"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2.2
    
    iCount = 6
    InputTypes(iCount).Name = "Width"
    InputTypes(iCount).Description = "Width of Door"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.8
    
    iCount = 7
    InputTypes(iCount).Name = "CTLength"
    InputTypes(iCount).Description = "Top Casement Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.8
    
    iCount = 8
    InputTypes(iCount).Name = "CTWidth"
    InputTypes(iCount).Description = "Top Casement Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 9
    InputTypes(iCount).Name = "CTThickness"
    InputTypes(iCount).Description = "Top Casement Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 10
    InputTypes(iCount).Name = "CTEdge"
    InputTypes(iCount).Description = "Top Casement Edge"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 11
    InputTypes(iCount).Name = "CBLength"
    InputTypes(iCount).Description = "Bottom Casement Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.8
    
    iCount = 12
    InputTypes(iCount).Name = "CBWidth"
    InputTypes(iCount).Description = "Bottom Casement Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 13
    InputTypes(iCount).Name = "CBThickness"
    InputTypes(iCount).Description = "Bottom Casement Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 14
    InputTypes(iCount).Name = "CBEdge"
    InputTypes(iCount).Description = "Bottom Casement Edge"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 15
    InputTypes(iCount).Name = "CLLength"
    InputTypes(iCount).Description = "Left Casement Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.8
    
    iCount = 16
    InputTypes(iCount).Name = "CLWidth"
    InputTypes(iCount).Description = "Left Casement Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 17
    InputTypes(iCount).Name = "CLThickness"
    InputTypes(iCount).Description = "Left Casement Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 18
    InputTypes(iCount).Name = "CLEdge"
    InputTypes(iCount).Description = "Left Casement Edge"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 19
    InputTypes(iCount).Name = "CRLength"
    InputTypes(iCount).Description = "Right Casement Length"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.8
    
    iCount = 20
    InputTypes(iCount).Name = "CRWidth"
    InputTypes(iCount).Description = "Right Casement Width"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.2
    
    iCount = 21
    InputTypes(iCount).Name = "CRThickness"
    InputTypes(iCount).Description = "Right Casement Thickness"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.1
    
    iCount = 22
    InputTypes(iCount).Name = "CREdge"
    InputTypes(iCount).Description = "Right Casement Edge"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 23
    InputTypes(iCount).Name = "PannelThickness"
    InputTypes(iCount).Description = "Thickness of the Door pannel"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.01
    
    iCount = 24
    InputTypes(iCount).Name = "Xposition"
    InputTypes(iCount).Description = "Door X Position Door"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 25
    InputTypes(iCount).Name = "Yposition"
    InputTypes(iCount).Description = "Door Y Position Doorl"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
    
    iCount = 26
    InputTypes(iCount).Name = "Zposition"
    InputTypes(iCount).Description = "Door Z Position Door"
    InputTypes(iCount).Properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0
                
                
' Outputs
    nOutputs = 13
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
    iCount = 1
    OutputTypes(iCount).Name = "LeftCasementFace_"
    OutputTypes(iCount).Description = "Face Of Left Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
        
    iCount = 2
    OutputTypes(iCount).Name = "LeftCasementStrip"
    OutputTypes(iCount).Description = "Left Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 3
    OutputTypes(iCount).Name = "RightCasementFace_"
    OutputTypes(iCount).Description = "Face Of Right Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).Name = "RightCasementStrip"
    OutputTypes(iCount).Description = "Right Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 5
    OutputTypes(iCount).Name = "TopCasementFace_"
    OutputTypes(iCount).Description = "Face Of Top Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).Name = "TopCasementStrip"
    OutputTypes(iCount).Description = "Top Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 7
    OutputTypes(iCount).Name = "BottomCasementFace_"
    OutputTypes(iCount).Description = "Bottom Of Top Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 8
    OutputTypes(iCount).Name = "BottomCasementStrip"
    OutputTypes(iCount).Description = "Bottom Casement Strip"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 9
    OutputTypes(iCount).Name = "PannelSolidFace_"
    OutputTypes(iCount).Description = "Face Of Pannel Solid"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 10
    OutputTypes(iCount).Name = "PannelSolid_"
    OutputTypes(iCount).Description = "Pannel Solid"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 11
    OutputTypes(iCount).Name = "DefaultSurface"
    OutputTypes(iCount).Description = "Default Mating Surface"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 12
    OutputTypes(iCount).Name = "CutoutCurve"
    OutputTypes(iCount).Description = "Cutout Complex lineString"
    OutputTypes(iCount).Aspect = Operation
    
    iCount = 13
    OutputTypes(iCount).Name = "OperationalEnvelope1"
    OutputTypes(iCount).Description = "Operational Envelope of the Door or Window"
    OutputTypes(iCount).Aspect = Operation

    
' Aspects
    nAspects = 2
    ReDim AspectTypes(1 To nAspects) As AspectType
    
    iCount = 1
    AspectTypes(iCount).Name = "SimplePhysical"
    AspectTypes(iCount).Description = "SimplePhysical"
    AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = SimplePhysical
    
    iCount = 2
    AspectTypes(iCount).Name = "CutoutContour"
    AspectTypes(iCount).Description = "Cutout Contour"
    AspectTypes(iCount).Properties = igREPRESENTATION_ISVBFUNCTION
    AspectTypes(iCount).AspectId = Operation
    
    
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub Class_Terminate()
    
    Set m_SymbolFactory = Nothing

End Sub

Public Sub CMCache(pInput As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)
 Const METHOD = "CMCache"
 On Error GoTo ErrorHandler
    
    If bArgToCache Then
        
        Dim pParameterContent As IJDParameterContent
        Set pParameterContent = New DParameterContent
        Let pParameterContent.Type = igString
        Let pParameterContent.String = "CachedInput4Support"
    
        Set pOutput = pParameterContent
    Else
        Err.Raise ReportError(Err, MODULE, METHOD, "CMCache reverse mode not implemented").Number
    End If
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
    Set oSymbolDefinition = m_SymbolFactory.CreateEntity(definition, ActiveConnection)
    Set m_SymbolFactory = Nothing
    
' Set definition progId and codebase
    oSymbolDefinition.ProgId = m_ProgId
    oSymbolDefinition.CodeBase = CodeBase
    oSymbolDefinition.Name = oSymbolDefinition.ProgId
   
' Initialize it
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    
    ' to avoid a new bulkload'
    'oSymbolDefinition.MetaDataOption = igSYMBOL_STATIC_METADATA
  oSymbolDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  oSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA

   
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
        Inputs(iCount).Name = InputTypes(iCount).Name
        Inputs(iCount).Description = InputTypes(iCount).Description
        Inputs(iCount).Properties = InputTypes(iCount).Properties
        PC.uomValue = InputTypes(iCount).uomValue
' Associate a default parameter to the input
        Inputs(iCount).DefaultParameterValue = PC
' ---------Set Input into the definition-
        InputsIf.SetInput Inputs(iCount), iCount + 1
        Set Inputs(iCount) = Nothing
    Next iCount
        
        
  'Cached input for support, connected by aggregator
  'put at the end because of a bug in symbols
  'Set default value as string to be a parameter (and have the igSYMBOL_CACHE_OPTION_SHARED option)
  Dim RefInput    As IMSSymbolEntities.IJDInput
  Set RefInput = New IMSSymbolEntities.DInput
  RefInput.Name = "Support"
  RefInput.Description = "Cached input for support"
  RefInput.Properties = igINPUT_IS_A_PARAMETER + igDESCRIPTION_OPTIONAL + igINPUT_IS_CACHED
  PC.Type = igString
  PC.String = "CachedInput4Support"
  RefInput.DefaultParameterValue = PC
  ' ---------Set Input into the definition-
  InputsIf.SetInput RefInput, iCount + 1
  Dim mthCookie As Long
  mthCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCache", imsCOOKIE_ID_USS_LIB)
  RefInput.IJDInputStdCustomMethod.SetCMCache imsCOOKIE_ID_USS_LIB, mthCookie
  Set RefInput = Nothing
    
        
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
                    ByVal outputColl As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    Dim arrayOfOutputs()    As String
    Dim iOutputs            As Integer
    Dim iCount              As Integer
    Dim lAspectID           As SymbolRepIds
    Dim AspectName          As String
    Dim bsucceded           As Boolean
    
    Set m_outputColl = outputColl
    
    bsucceded = False
    For iCount = 1 To nAspects
        If AspectTypes(iCount).Name = repName Then
            bsucceded = True
            lAspectID = AspectTypes(iCount).AspectId
            Exit For
        End If
    Next iCount
    If Not bsucceded Then
        Err.Raise ReportError(Err, MODULE, METHOD, "Unknown Aspect; Check Aspect data / code").Number
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
    If (repName = "SimplePhysical") Then
        Set objAspect = New CSimplePhysical
    ElseIf (repName = "CutoutContour") Then
        Set objAspect = New CCutoutContour
    End If
    objAspect.Run outputColl, arrayOfInputs, arrayOfOutputs
   
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function
