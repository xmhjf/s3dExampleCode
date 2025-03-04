VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BB2PumpSym"
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
'   BB2PumpSym.cls
'   ProgID:         SP3DHoriPumpBB2Asm.BB2PumpSym
'   Author:         RUK
'   Creation Date:  Tuesday, April 3 2007
'
'   Description:
'   This is Horizontal Pump Of Type BB2 Assembly.
'   This Assembly Uses One Aspect
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             25
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "BB2PumpSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHoriPumpBB2Asm"
    m_oSymbolHelper.ClassName = "BB2PumpSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 15
    
    m_oSymbolHelper.AddInputDef 1, "ImpellerDiameter", "D-Motor Diameter", 0.5
    m_oSymbolHelper.AddInputDef 2, "DischargeNozLoc", "Z-Center Line to Discharge Center Line", 0.2
    m_oSymbolHelper.AddInputDef 3, "SuctionNozLoc", "S-Center Line to Suction Center Line", 0.175
    m_oSymbolHelper.AddInputDef 4, "CenterToNozzFace", "X-Height Of Nozzle Face From Center", 0.35
    m_oSymbolHelper.AddInputDef 5, "ShaftFrontEnd", "Y-Center to Impller Shaft Front End", 0.45
    m_oSymbolHelper.AddInputDef 6, "CouplingStart", "W-Center to Impeller Shaft Rear End", 0.55
    m_oSymbolHelper.AddInputDef 7, "CoupEndToMotorEnd", "CMAX-Motor Shaft Front End to Motor Rear End", 0.8
    m_oSymbolHelper.AddInputDef 8, "CouplingLength", "DBSE-Distace Between Shaft Ends", 0.125
    m_oSymbolHelper.AddInputDef 9, "SupportBottomLoc", "HC-Center to Base Plate Bottom", 0.45
    m_oSymbolHelper.AddInputDef 10, "FounPortStartPoint", "HR-Center to Foundation Port Start", 0.5
    m_oSymbolHelper.AddInputDef 11, "SupportHeight", "HG-Base Plate Height", 0.1
    m_oSymbolHelper.AddInputDef 12, "SupportWidth", "HA-Base Plate Width", 0.75
    m_oSymbolHelper.AddInputDef 13, "SupportLength", "HB-Base Plate Length", 1.975
    m_oSymbolHelper.AddInputDef 14, "XBoltHole", "HE-Hole Bolt X Location", 0.35
    m_oSymbolHelper.AddInputDef 15, "YBoltHole", "HF-Hole Bolt Y Location", 0.96225
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 25
    
    m_oSymbolHelper.AddOutputDef 1, "ShaftFrontEnd", "Shaft Front End", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Flange1", "Flange 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "PumpCylinder1", "Pump Cylinder 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "PumpCylinder2", "Pump Cylinder 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Flange2", "Flange 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "PumpCylinder3", "Pump Cylinder 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Flange3", "Flange 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Flange4", "Flange 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "ShaftRearEnd", "Shaft Rear End", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "CouplingFlange1", "Coupling Flange 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Coupling", "Coupling", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "CouplingFlange2", "Coupling Flange 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "MotorFlange", "Motor Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "MotorFrontHead", "Motor Front Head", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "MotorBody", "Motor Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "MotorRearHead", "Motor Rear Head", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "Shaft", "Shaft", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "MotorSupport", "Motor Support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "PumpSupport_", "Pump Support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "BasePlate_", "Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Edges_", "Edges For Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Points_", "Points On Each Face Of Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "EqpFoundationPort", "Equipment Foundation Port", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "EqpControlPoint", "Contol Point at the Origin", SimplePhysical

    'Aspects
    m_oSymbolHelper.NumAspects = 1
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    
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


