VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSValve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSValve.cls
'   Author:         dkl
'   Creation Date:  Tuesday, Jan 23 2007
'   Description:
'   This symbol is based on a functionality of a SP3DCSValve with in-built Gear Operator.
'   The in-built operator is based on SP3DOP_291.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   23.Jan.2007     dkl    CR-113194. Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSValve:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DCSValveOpGear"
    m_oSymbolHelper.ClassName = "CSValve"
    
' Inputs
    m_oSymbolHelper.NumInputs = 35
    m_oSymbolHelper.AddInputDef 1, "Face1toCenter", "Face 1 to Center", 0.25
    m_oSymbolHelper.AddInputDef 2, "Face2toCenter", "Face 2 to Center", 0.25
    m_oSymbolHelper.AddInputDef 3, "OperatorHeight", "Total Height of Operator", 0.6
    m_oSymbolHelper.AddInputDef 4, "OperatorDiameter", "Diameter of Operator", 0.4
    m_oSymbolHelper.AddInputDef 5, "ActuatorLength", "Operator to Top of Hand Wheel", 0.18
    m_oSymbolHelper.AddInputDef 6, "OperatorDepth", "Thickness of Operator", 0.15
    m_oSymbolHelper.AddInputDef 7, "ActuatorDiameter", "Diameter of the Hand Wheel", 0.15
    m_oSymbolHelper.AddInputDef 8, "ActuatorOffset", "Distance between Operator and Handwheel", 0.1
    m_oSymbolHelper.AddInputDef 9, "StemExtension", "Top of Operator to Top of Stem", 0.075
    m_oSymbolHelper.AddInputDef 10, "InsulationThickness", "Insulation Thickness", 0.25
    m_oSymbolHelper.AddInputDef 11, "Npd", "NPD", 100
    m_oSymbolHelper.AddInputDef 12, "NpdUnitType", "Npd Unit Type", "mm"
    m_oSymbolHelper.AddInputDef 13, "EndPreparation", "End Preparation", 21
    m_oSymbolHelper.AddInputDef 14, "ScheduleThickness", "Schedule Thickness", 100
    m_oSymbolHelper.AddInputDef 15, "EndStandard", "End Standard", 5
    m_oSymbolHelper.AddInputDef 16, "PressureRating", "Pressure Rating", 35
    m_oSymbolHelper.AddInputDef 17, "FlowDirection", "Flow Direction", 3
    m_oSymbolHelper.AddInputDef 18, "Id1", "Id1", "VNoz1"
    m_oSymbolHelper.AddInputDef 19, "PortIndex1", "Port Index 1", 1
    m_oSymbolHelper.AddInputDef 20, "Npd1", "Npd 1", 0
    m_oSymbolHelper.AddInputDef 21, "NpdUnitType1", "Npd Unit Type 1", ""
    m_oSymbolHelper.AddInputDef 22, "EndPreparation1", "End Preparation 1", 0
    m_oSymbolHelper.AddInputDef 23, "ScheduleThickness1", "Schedule Thickness 1", 0
    m_oSymbolHelper.AddInputDef 24, "EndStandard1", "End Standard 1", 0
    m_oSymbolHelper.AddInputDef 25, "PressureRating1", "Pressure Rating 1", 0
    m_oSymbolHelper.AddInputDef 26, "FlowDirection1", "Flow Direction 1", 0
    m_oSymbolHelper.AddInputDef 27, "Id2", "Id2", "VNoz2"
    m_oSymbolHelper.AddInputDef 28, "PortIndex2", "Port Index 2", 2
    m_oSymbolHelper.AddInputDef 29, "Npd2", "Npd 2", 0
    m_oSymbolHelper.AddInputDef 30, "NpdUnitType2", "Npd Unit Type 2", ""
    m_oSymbolHelper.AddInputDef 31, "EndPreparation2", "End Preparation 2", 0
    m_oSymbolHelper.AddInputDef 32, "ScheduleThickness2", "Schedule Thickness 2", 0
    m_oSymbolHelper.AddInputDef 33, "EndStandard2", "End Standard 2", 0
    m_oSymbolHelper.AddInputDef 34, "PressureRating2", "Pressure Rating 2", 0
    m_oSymbolHelper.AddInputDef 35, "FlowDirection2", "Flow Direction 2", 0
          
' Outputs
    m_oSymbolHelper.NumOutputs = 13
    m_oSymbolHelper.AddOutputDef 1, "VNoz1", "Nozzle 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "VNoz2", "Nozzle 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "LeftBodyCone", "Left Body Cone", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "RightBodyCone", "Right Body Cone", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Stem", "Vertical Stem", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Operator", "Horizontal Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "AxisCylinder", "Verical Axis Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "HandWheelCylinder", "Top Hand Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "BodyCylIns", "Insulation Cylinderical Body", Insulation
    m_oSymbolHelper.AddOutputDef 10, "MaintStem", "Maintenance Vertical Stem", Maintenance
    m_oSymbolHelper.AddOutputDef 11, "MaintOperator", "Maintenance Horizontal Operator", Maintenance
    m_oSymbolHelper.AddOutputDef 12, "MaintAxisCylinder", "aintenance Verical Axis Cylinder", Maintenance
    m_oSymbolHelper.AddOutputDef 13, "MaintHandWheelCylinder", "Maintenance Top Hand Cylinder", Maintenance
        
' Aspects
    m_oSymbolHelper.NumAspects = 3
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "SimplePhysical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "Maintenance", "Maintenance", Maintenance
     
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    
    Set m_oSymbolHelper = Nothing

End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    
' This method is in charge of the creation of the symbol definition object
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
