VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlatformSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:04 AM"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPlatformSym.cls
'   ProgID:         SP3DCESVVessel2PlatfAsm.CPlatformSym
'   Author:         MS
'   Creation Date:  Wednesday, Feb 5 2003
'   Description:
'   1. The symbol is based on the Vertical Cylindrical Equipment with Skirt(E210 of PDS guide)
'   2.  The symbol consists of the basic equipment and the Platform together. It contains Two Platforms
'   one after the conical skirt and other near the top third of the vessel. The two platform geometry(like platformangle) would be the same.
'   3. The Equipment has 3 fixed number of nozzles (one on the top and two on the side of the vessel) and
'   all the nozzles would be fully parametric. Two nozzles on the vessel side wall can be relocated based
'   on the parameters NozzleOrientation2,NozzlePosition2,NozzleOrientation3 and NozzlePosition3
'   4.  Control Point is added to the Equipment
'   5. The 3 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'    9 Feb 2003  Nabil Nicolas      Redefine the input parameters for Nozzles to optimize the Interfaces.
'                                   Redefine the Progid (Project & classes names).
'   29 Jan 2003  BG                 TR# 39362 Modified the control point diameter to zero to avoid interference.
'   09.Jul.2003  SymbolTeam(India)  Copyright Information, Header  is added/Updated.
'   19.Mar.2004  SymbolTeam(India)  TR 56826 Removed Msgbox
'   29.Nov.2004  V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005  MS                 CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006   KKC                DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007   PS                 CR-132021  Modified the default value of PlatformAngle
'   21.APR.2008     PK     DI-132538    CreateObject() has been replaced with 'New' representation class in the InvokeRepresentation procedure
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CPlatformSym:"  'Used for error messages

Private m_SymbolFactory     As IMSSymbolEntities.DSymbolEntitiesFactory
Private m_OutputColl        As Object
Private m_ProgId            As String
Private m_ProjName          As String
Private m_ClassName         As String

Private InputTypes()        As InputType
Private m_TextInputTypes()  As TextInputType
Private OutputTypes()       As OutputType
Private AspectTypes()       As AspectType
Private nAspects            As Integer
Private nInputs             As Integer
Private iTextCount         As Integer
Private nOutputs            As Integer
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount As Integer
    
    m_ProjName = "SP3DCESVVessel2PlatfAsm"
    m_ClassName = "CPlatformSym"
    m_ProgId = m_ProjName & "." & m_ClassName

' Inputs
    nInputs = 10
    ReDim InputTypes(1 To nInputs) As InputType

    iCount = 1
    InputTypes(iCount).name = "VesselHeight"
    InputTypes(iCount).description = "Shell Height P2"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 16
    
    iCount = 2
    InputTypes(iCount).name = "VesselDiameter"
    InputTypes(iCount).description = "Shell Diameter P3"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2.1
    
    iCount = 3
    InputTypes(iCount).name = "SkirtTopDiameter"
    InputTypes(iCount).description = "Support Top Diameter P5"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2.1
    
    iCount = 4
    InputTypes(iCount).name = "SkirtBottomDiameter"
    InputTypes(iCount).description = "Support Bottom Diameter P6"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 3.5
    
    iCount = 5
    InputTypes(iCount).name = "SkirtHeight"
    InputTypes(iCount).description = "Skirt Height P7"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2.8
    
    iCount = 6
    InputTypes(iCount).name = "VesselStartHeight"
    InputTypes(iCount).description = "Vessel Start Height P4"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2.8
 
      'Platform Input parameters. The two platform are based on the same below four input parameters.
    iCount = 7
    InputTypes(iCount).name = "PlatformWidth"
    InputTypes(iCount).description = "Platform Width"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 2
    
    iCount = 8
    InputTypes(iCount).name = "PlatformHeight"
    InputTypes(iCount).description = "Platform Height"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 0.4
        
    iCount = 9
    InputTypes(iCount).name = "PlatformRadius"
    InputTypes(iCount).description = "Platform Radius"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 1.056
    
    iCount = 10
    InputTypes(iCount).name = "PlatformAngle"
    InputTypes(iCount).description = "Platform Angle"
    InputTypes(iCount).properties = igINPUT_IS_A_PARAMETER
    InputTypes(iCount).uomValue = 6 * Atn(1)
    
   
        
' Outputs
    nOutputs = 24
    ReDim OutputTypes(1 To nOutputs) As OutputType
    
  
    iCount = 1
    OutputTypes(iCount).name = "Vessel"
    OutputTypes(iCount).description = "Vessel Shell"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 2
    OutputTypes(iCount).name = "Support"
    OutputTypes(iCount).description = "Truncated Cone Support "
    OutputTypes(iCount).Aspect = SimplePhysical
       
    iCount = 3
    OutputTypes(iCount).name = "DefaultSurface"
    OutputTypes(iCount).description = "Default Surface"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 4
    OutputTypes(iCount).name = "Plat1Plane1"
    OutputTypes(iCount).description = "Plat1Plane1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 5
    OutputTypes(iCount).name = "Plat1Plane2"
    OutputTypes(iCount).description = "Plat1Plane2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 6
    OutputTypes(iCount).name = "Plat1Revolution1"
    OutputTypes(iCount).description = "Plat1Revolution1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 7
    OutputTypes(iCount).name = "Plat1Revolution2"
    OutputTypes(iCount).description = "Plat1Revolution2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 8
    OutputTypes(iCount).name = "Plat1TopSurface"
    OutputTypes(iCount).description = "Plat1TopSurface"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 9
    OutputTypes(iCount).name = "Plat1BottomSurface"
    OutputTypes(iCount).description = "Plat1BottomSurface"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 10
    OutputTypes(iCount).name = "Plat1Line1"
    OutputTypes(iCount).description = "Plat1Line1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 11
    OutputTypes(iCount).name = "Plat1Line2"
    OutputTypes(iCount).description = "Plat1Line2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 12
    OutputTypes(iCount).name = "Plat1Line3"
    OutputTypes(iCount).description = "Plat1Line3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 13
    OutputTypes(iCount).name = "Plat1Line4"
    OutputTypes(iCount).description = "Plat1Line4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 14
    OutputTypes(iCount).name = "Plat2Plane1"
    OutputTypes(iCount).description = "Plat1Plane1"
    OutputTypes(iCount).Aspect = SimplePhysical

    iCount = 15
    OutputTypes(iCount).name = "Plat2Plane2"
    OutputTypes(iCount).description = "Plat1Plane2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 16
    OutputTypes(iCount).name = "Plat2Revolution1"
    OutputTypes(iCount).description = "Plat1Revolution1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 17
    OutputTypes(iCount).name = "Plat2Revolution2"
    OutputTypes(iCount).description = "Plat1Revolution2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 18
    OutputTypes(iCount).name = "Plat2TopSurface"
    OutputTypes(iCount).description = "Plat1TopSurface"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 19
    OutputTypes(iCount).name = "Plat2BottomSurface"
    OutputTypes(iCount).description = "Plat1BottomSurface"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 20
    OutputTypes(iCount).name = "Plat2Line1"
    OutputTypes(iCount).description = "Plat1Line1"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 21
    OutputTypes(iCount).name = "Plat2Line2"
    OutputTypes(iCount).description = "Plat1Line2"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 22
    OutputTypes(iCount).name = "Plat2Line3"
    OutputTypes(iCount).description = "Plat1Line3"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 23
    OutputTypes(iCount).name = "Plat2Line4"
    OutputTypes(iCount).description = "Plat1Line4"
    OutputTypes(iCount).Aspect = SimplePhysical
    
    iCount = 24
    OutputTypes(iCount).name = "CESimVerVesselControlPoint"
    OutputTypes(iCount).description = "Control Point of  Simple Vertical Vessel with Platform"
    OutputTypes(iCount).Aspect = SimplePhysical
        
    
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
            If OutputTypes(iOutput).Aspect And AspectTypes(iCount).AspectId Then
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
    Debug.Assert False

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
