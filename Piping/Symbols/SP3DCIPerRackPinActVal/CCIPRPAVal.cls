VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCIPRPAVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCIPRPAVal.cls
'   ProgID:         SP3DCIPerRackPinActVal.CCIPRPAVal
'   Author:         MS
'   Creation Date:  Friday, May 23 2003
'   Description:
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRI2).
'   The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CCIPRPAVal:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices


Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DCIPerRackPinActVal"
    m_oSymbolHelper.ClassName = "CCIPRPAVal"
    
' Inputs
    m_oSymbolHelper.NumInputs = 31
    m_oSymbolHelper.AddInputDef 1, "Face1toCenter", "Face1 to Center", 0.25
    m_oSymbolHelper.AddInputDef 2, "FacetoFace", "Face to Face", 0.5
    m_oSymbolHelper.AddInputDef 3, "ActuatorHeight", "Actuator Height", 0.75
    m_oSymbolHelper.AddInputDef 4, "ActuatorCylDiameter", "Actuator Cylinder Diameter", 0.75
    m_oSymbolHelper.AddInputDef 5, "ActuatorDiameter", "Actuator Diameter", 0.1
    m_oSymbolHelper.AddInputDef 6, "ActuatorCylLength", "Actuator Cylinder Length", 0.25
    m_oSymbolHelper.AddInputDef 7, "Npd", "NPD", 100
    m_oSymbolHelper.AddInputDef 8, "EndPreparation", "End Preparation", 21
    m_oSymbolHelper.AddInputDef 9, "ScheduleThickness", "Schedule Thickness", 100
    m_oSymbolHelper.AddInputDef 10, "EndStandard", "End Standard", 5
    m_oSymbolHelper.AddInputDef 11, "PressureRating", "Pressure Rating", 35
    m_oSymbolHelper.AddInputDef 12, "FlowDirection", "Flow Direction", 3
    m_oSymbolHelper.AddInputDef 13, "PortIndex1", "Port Index1", 1
    m_oSymbolHelper.AddInputDef 14, "Npd1", "NPD1", 0
    m_oSymbolHelper.AddInputDef 15, "EndPreparation1", "End Preparation1", 0
    m_oSymbolHelper.AddInputDef 16, "ScheduleThickness1", "Schedule Thickness1", 0
    m_oSymbolHelper.AddInputDef 17, "EndStandard1", "End Standard1", 0
    m_oSymbolHelper.AddInputDef 18, "PressureRating1", "Pressure Rating1", 0
    m_oSymbolHelper.AddInputDef 19, "FlowDirection1", "Flow Direction1", 0
    m_oSymbolHelper.AddInputDef 20, "PortIndex2", "Port Index2", 2
    m_oSymbolHelper.AddInputDef 21, "Npd2", "NPD2", 0
    m_oSymbolHelper.AddInputDef 22, "EndPreparation2", "End Preparation2", 0
    m_oSymbolHelper.AddInputDef 23, "ScheduleThickness2", "Schedule Thickness2", 0
    m_oSymbolHelper.AddInputDef 24, "EndStandard2", "End Standard2", 0
    m_oSymbolHelper.AddInputDef 25, "PressureRating2", "Pressure Rating2", 0
    m_oSymbolHelper.AddInputDef 26, "FlowDirection2", "Flow Direction2", 0
    m_oSymbolHelper.AddInputDef 27, "Id1", "Id1", "1"
    m_oSymbolHelper.AddInputDef 28, "Id2", "Id2", "2"
    m_oSymbolHelper.AddInputDef 29, "NpdUnitType", "Npd Unit Type", "mm"
    m_oSymbolHelper.AddInputDef 30, "NpdUnitType1", "Npd Unit Type1", ""
    m_oSymbolHelper.AddInputDef 31, "NpdUnitType2", "Npd Unit Type2", ""

' Outputs
    m_oSymbolHelper.NumOutputs = 6
    m_oSymbolHelper.AddOutputDef 1, "PNoz1", "Nozzle 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "PNoz2", "Nozzle 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "ConvergingCone", "Converging Cone", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "DivergingCone", "Diverging Cone", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "ActuatorStem", "Actuator Stem", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "ActuatorBody", "Actuator Body", SimplePhysical
       
    ' Aspects
    m_oSymbolHelper.NumAspects = 1
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    
    Exit Sub
    
ErrorHandler:
 Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    Set m_oSymbolHelper = Nothing
End Sub

' note that the rest of the code is generated by the Symbol Wizard and should not need to be modified
Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    ' call symbol services default implementation of this method
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
On Error GoTo Errx
    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolHelper.InstanciateDefinition(CodeBase, defParameters, ActiveConnection)
    
    Exit Function

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
Const METHOD = "IJDUserSymbolServices_GetDefinitionName:"
  On Error GoTo Errx
      IJDUserSymbolServices_GetDefinitionName = m_oSymbolHelper.ProjectName + "." + m_oSymbolHelper.ClassName
    
    Exit Function

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
On Error GoTo Errx
    m_oSymbolHelper.InitializeSymbolDefinition pSymbolDefinition
Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
      
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
On Error GoTo Errx
    m_oSymbolHelper.InvokeRepresentation sblOcc, repName, outputcoll, arrayOfInputs
        
Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
Const METHOD = "IJDUserSymbolServices_EditOccurence:"
On Error GoTo Errx
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
        Exit Function

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext

End Function




