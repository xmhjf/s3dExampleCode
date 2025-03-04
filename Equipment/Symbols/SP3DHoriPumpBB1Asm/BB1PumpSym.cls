VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BB1PumpSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   BB1PumpSym.cls
'   ProgID:         SP3DHoriPumpBB1Asm.BB1PumpSym
'   Author:         VRK
'   Creation Date:  Tuesday, April 3 2007
'
'   Description:
'   This is Horizontal Pump Of Type BB1 Assembly.
'   This Assembly Uses One Aspect
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             30
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "BB1PumpSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHoriPumpBB1Asm"
    m_oSymbolHelper.ClassName = "BB1PumpSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 16
    
    m_oSymbolHelper.AddInputDef 1, "PumpFacetoDischargeCL", "Y-Center to Impeller Shaft Front End", 0.584
    m_oSymbolHelper.AddInputDef 2, "PumpCouplingFacetoDischargeCL", "W-Center to Impeller Shaft Rear End", 0.635
    m_oSymbolHelper.AddInputDef 3, "CouplingtoMotorEnd", "CMax-Motor Shaft Front End to Motor Rear End", 0.889
    m_oSymbolHelper.AddInputDef 4, "PumpCLtoDischargeFace", "T-Pump Centerline to Discharge Face", 0.508
    m_oSymbolHelper.AddInputDef 5, "PumpCLtoSuctionFace", "U-Pump Centerline to Suction Face", 0.508
    m_oSymbolHelper.AddInputDef 6, "PumpCLtoDischargeCL", "Z-PumpCenterline to Discharge Centerline", 0.2
    m_oSymbolHelper.AddInputDef 7, "PumpCLtoSuctionCL", "S-PumpCenterline to Suction Centerline", 0.2
    m_oSymbolHelper.AddInputDef 8, "ImpellerDiameter", "D-Impeller Diameter", 0.6
    m_oSymbolHelper.AddInputDef 9, "DistBetweenShaftEnds", "HT-Distance Between Shaft Ends", 0.152
    m_oSymbolHelper.AddInputDef 10, "BaseplateLength", "HB-Baseplate Length", 2.489
    m_oSymbolHelper.AddInputDef 11, "BaseplateWidth", "HA-Baseplate Width", 0.762
    m_oSymbolHelper.AddInputDef 12, "BaseplateHeight", "Baseplate Height", 0.2
    m_oSymbolHelper.AddInputDef 13, "BaseplateBottomtoPumpCL", "HD-Baseplate Bottom to Pump Centerline", 0.635
    m_oSymbolHelper.AddInputDef 14, "FoundationPortStartPoint", "HR-Center to Foundation Port Start", 0.254
    m_oSymbolHelper.AddInputDef 15, "XboltHole", "HE-Bolt X Hole Location", 0.254
    m_oSymbolHelper.AddInputDef 16, "YboltHole", "HF-Bolt Y Hole Location", 0.8636

    'Outputs
    m_oSymbolHelper.NumOutputs = 30
    
    m_oSymbolHelper.AddOutputDef 1, "MotorCylinderBody", "Motor Cylinder Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "MotorShaftBody1", "Motor Shaft Body1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "MotorShaftBody2", "Motor Shaft Body2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "PumpShaftBody", "Pump Shaft Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "CouplingCylinder", "Coupling Cylinder Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "ImpellerCylinder", "Impeller Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "PumpCylinder1", "Pump Cylinder1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "PumpCylinder2", "Pump Cylinder2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "PumpCylinder3", "Pump Cylinder3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "PumpCylinder4", "Pump Cylinder4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "PumpCylinder5", "Pump Cylinder5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "PumpCylinder6", "Pump Cylinder6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "PumpCylinder7", "Pump Cylinder7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "PumpCylinder8", "Pump Cylinder8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "CouplingCylinder", "Coupling Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "PumpSupport1_", "Pump Support1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "PumpSupport2_", "Pump Support2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "MotorSupport1_", "Motor Support1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "MotorSupport2_", "Motor Support2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "SupportPlate1", "Support Plate1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "SupportPlate2", "Support Plate2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "SupportPlateDome1", "SupportPlateDome1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "SupportPlateDome2", "SupportPlateDome2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "BasePlate_", "BasePlate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "Edges_", "Edges For Baseplate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "Points_", "Points For Baseplate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "Body1", "Body1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "EqpFoundationPort", "Foundation Port under support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 30, "PumpOrigin", "Control Point at the Origin", SimplePhysical

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

           







