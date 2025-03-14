VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSPTSteamTrapWOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   File:   CSPTSteamTrapWOS.cls
'   ProgID:         SP3DCSPTSteamTrapWOS.CSPTSteamTrapWOS
'   Author: Sundar(svsmylav)
'   Creation Date:  Friday, Oct 8 2004
'
'   Description:
'       ASPECTS
'           Simple physical, SimplePhysical, 0
'           Detailed physical, DetailedPhysical, 4
'           Insulation, Insulation, 5
'           Operation, Operation, 6
'           Maintenance, Maintenance, 7
'           Reference geometry, ReferenceGeometry, 8
'       END
'     This is PDS on-the-fly S47A1  Float Thermostatic Steam Trap without Strainer Symbol.
'
'  Change History:
'**********************************************************************************

Option Explicit

Private Const MODULE = "CSPTSteamTrapWOS"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DCSPTSteamTrapWOS"
    m_oSymbolHelper.ClassName = "CSPTSteamTrapWOS"

'**********************************************************************************
'**********************************************************************************
'   BEGIN
'**********************************************************************************
'**********************************************************************************
    m_oSymbolHelper.NumInputs = 30
    m_oSymbolHelper.AddInputDef 1, "Height", "Height", 0.6
    m_oSymbolHelper.AddInputDef 2, "FacetoFace", "Faceto Face", 0.8
    m_oSymbolHelper.AddInputDef 3, "Length", "Length", 0.6
    m_oSymbolHelper.AddInputDef 4, "Offset", "Offset", 0.35
    m_oSymbolHelper.AddInputDef 5, "Width", "Width", 0.2
    m_oSymbolHelper.AddInputDef 6, "Npd", "NPD", 100
    m_oSymbolHelper.AddInputDef 7, "EndPreparation", "EndPreparation", 21
    m_oSymbolHelper.AddInputDef 8, "ScheduleThickness", "ScheduleThickness", 100
    m_oSymbolHelper.AddInputDef 9, "EndStandard", "EndStandard", 5
    m_oSymbolHelper.AddInputDef 10, "PressureRating", "PressureRating", 35
    m_oSymbolHelper.AddInputDef 11, "FlowDirection", "FlowDirection", 3
    m_oSymbolHelper.AddInputDef 12, "PortIndex1", "PortIndex1", 1
    m_oSymbolHelper.AddInputDef 13, "Npd1", "NPD1", 0
    m_oSymbolHelper.AddInputDef 14, "EndPreparation1", "EndPreparation1", 0
    m_oSymbolHelper.AddInputDef 15, "ScheduleThickness1", "ScheduleThickness1", 0
    m_oSymbolHelper.AddInputDef 16, "EndStandard1", "EndStandard1", 0
    m_oSymbolHelper.AddInputDef 17, "PressureRating1", "PressureRating1", 0
    m_oSymbolHelper.AddInputDef 18, "FlowDirection1", "FlowDirection1", 0
    m_oSymbolHelper.AddInputDef 19, "PortIndex2", "PortIndex2", 2
    m_oSymbolHelper.AddInputDef 20, "Npd2", "NPD2", 0
    m_oSymbolHelper.AddInputDef 21, "EndPreparation2", "EndPreparation2", 0
    m_oSymbolHelper.AddInputDef 22, "ScheduleThickness2", "ScheduleThickness2", 0
    m_oSymbolHelper.AddInputDef 23, "EndStandard2", "EndStandard2", 0
    m_oSymbolHelper.AddInputDef 24, "PressureRating2", "PressureRating2", 0
    m_oSymbolHelper.AddInputDef 25, "FlowDirection2", "FlowDirection2", 0
    m_oSymbolHelper.AddInputDef 26, "Id1", "Id1", "VNoz1"
    m_oSymbolHelper.AddInputDef 27, "Id2", "Id2", "VNoz2"
    m_oSymbolHelper.AddInputDef 28, "NpdUnitType", "Npd Unit Type", "mm"
    m_oSymbolHelper.AddInputDef 29, "NpdUnitType1", "Npd Unit Type 1", ""
    m_oSymbolHelper.AddInputDef 30, "NpdUnitType2", "Npd Unit Type 2", ""

    m_oSymbolHelper.NumOutputs = 3
    m_oSymbolHelper.AddOutputDef 1, "VNoz1", "Nozzle 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "VNoz2", "Nozzle 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "oBox", "Box", SimplePhysical
    
'**********************************************************************************
'**********************************************************************************
'   END
'**********************************************************************************
'**********************************************************************************
    
' Aspects
    m_oSymbolHelper.NumAspects = 1
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "SimplePhysicalAspect Description", 1
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    
    Set m_oSymbolHelper = Nothing

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'NOTE: The code below this line was generated by the Symbol Wizard and should not be modified
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    
' This method is in charge of the creation of the symbol definition object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler
    
    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolHelper.InstanciateDefinition(CodeBase, defParameters, ActiveConnection)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName:"
    On Error GoTo ErrorHandler
    
    IJDUserSymbolServices_GetDefinitionName = m_oSymbolHelper.ProjectName + "." + m_oSymbolHelper.ClassName
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler
   
    m_oSymbolHelper.InitializeSymbolDefinition pSymbolDefinition
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    m_oSymbolHelper.InvokeRepresentation sblOcc, repName, outputcoll, arrayOfInputs
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    Const METHOD = "IJDUserSymbolServices_EditOccurence:"
    On Error GoTo ErrorHandler
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Function

