VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BB5PumpSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:59 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   BB5PumpSym.cls
'   ProgID:         SP3DHoriPumpBB5Asm.BB5PumpSym
'   Author:         RUK
'   Creation Date:  Tuesday, April 10 2007
'
'   Description:
'   This is Horizontal Pump Of Type BB5 Assembly.
'   This Assembly Uses four Aspects
'   No.Of Aspects           Name of Aspect                  No.of Outputs
'       1                       Physical                            25
'       2                       Maintenance                         1
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   05.10.2007      RUK     CP 128473 Chaged the names of "MotorDiameter", "PumpTopSupportWidth"
'                           to "Motor Height", "Motor Width.
'   05.10.2007      RUK     CP 128584 Added new attributes JunBoxHeight, JunBoxWidth, JunBoxLength
'                               and used to create the Junction Box of Motor.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "BB5PumpSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHoriPumpBB5Asm"
    m_oSymbolHelper.ClassName = "BB5PumpSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 19
    
    m_oSymbolHelper.AddInputDef 1, "MotorDiameter", "D-Motor Diameter", 0.508
    m_oSymbolHelper.AddInputDef 2, "CenterToNozzFace", "NH-Height Of Nozzle Face From Center", 0.3302
    m_oSymbolHelper.AddInputDef 3, "DischargeToSuction", "X-Discharge Axis to Suction Axis", 0.508
    m_oSymbolHelper.AddInputDef 4, "ImpShaftFrontEnd", "Y-Center to Impeller Shaft Front End", 0.762
    m_oSymbolHelper.AddInputDef 5, "ImpShaftRearEnd", "W-Center to Impeller Shaft Rear End", 1.01
    m_oSymbolHelper.AddInputDef 6, "DistBetShaftEnds", "DBSE-Distace Between Shaft Ends", 0.127
    m_oSymbolHelper.AddInputDef 7, "MotorShaftToEnd", "CMAX-Motor Shaft Front End to Motor Rear End", 0.762
    m_oSymbolHelper.AddInputDef 8, "PumpSupTopWidth", "PMAX-Pump Support Top Width", 0.635
    m_oSymbolHelper.AddInputDef 9, "DismantlingLength", "EX-Dismantling Length", 1.143
    m_oSymbolHelper.AddInputDef 10, "BasePlateBottom", "HD-Center to Base Plate Bottom", 0.508
    m_oSymbolHelper.AddInputDef 11, "CenToFouPortStart", "HR-Center to Base Plate Start", 0.508
    m_oSymbolHelper.AddInputDef 12, "BasePlateWidth", "HA-Base Plate Width", 0.762
    m_oSymbolHelper.AddInputDef 13, "BasePlateLength", "HB-Base Plate Length", 2.921
    m_oSymbolHelper.AddInputDef 14, "BasePlateHeight", "BD-Base Plate Height", 0.127
    m_oSymbolHelper.AddInputDef 15, "JunBoxHeight", "JH-Junction Box Height", 0.01
    m_oSymbolHelper.AddInputDef 16, "JunBoxWidth", "JW-Junction Box Width", 0.01
    m_oSymbolHelper.AddInputDef 17, "JunBoxLength", "JL-Junction Box Length", 0.01
    m_oSymbolHelper.AddInputDef 18, "XBoltHole", "HE-Hole Bolt X Location", 0.35
    m_oSymbolHelper.AddInputDef 19, "YBoltHole", "HF-Hole Bolt Y Location", 0.96225
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 25
    
    m_oSymbolHelper.AddOutputDef 1, "ShaftFrontEnd", "Shaft Front End", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "FrontEndCyl", "Front End Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Flange1", "Flange 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ImpellerPart1", "Impeller Part 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "ImpellerPart2", "Impeller Part 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "ImpellerPart3", "Impeller Part 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Flange2", "Flange 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "RearEndCyl", "Rear End Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "ShaftRearEnd", "Shaft Rear End", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "CouplingFlange1", "Coupling Flange 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Coupling", "Coupling", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "CouplingFlange2", "Coupling Flange 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "MotorFlange", "Motor Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "Motor", "Motor", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "MotorSupport", "Motor Support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "MotorBox", "Motor Box", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "Shaft", "Shaft", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "PumpSupport_", "Pump Support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "BasePlate_", "Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Edges_", "Edges For Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Points_", "Points On Each Face Of Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "EqpFoundationPort", "Equipment Foundation Port", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "EqpControlPoint", "Contol Point at the Origin", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "MaintenanceCylinder", "Maintenance Cylinder", Maintenance
    
    'Aspects
    m_oSymbolHelper.NumAspects = 2
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Maintenance", "Maintenance", Maintenance
    
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Class_Terminate()
    Set m_oSymbolHelper = Nothing
End Sub

'Note that the rest of the code is generated by the Symbol Wizard and should not need to be modified
Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    'Call symbol services default implementation of this method
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
    
    'The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
    
    Exit Function
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function
