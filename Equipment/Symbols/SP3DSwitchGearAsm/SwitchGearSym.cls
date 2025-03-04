VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SwitchGearSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:37 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   SwitchGearSym.cls
'   Author: Veena
'   Creation Date:  Friday, Sep 29 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit

Private Const MODULE = "SwitchGearSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DSwitchGearAsm"
    m_oSymbolHelper.ClassName = "SwitchGearSym"
    
    'Inputs
    m_oSymbolHelper.NumInputs = 5
    
    m_oSymbolHelper.AddInputDef 1, "SwitchGearHeight", "Height of the SwitchGear", 1#
    m_oSymbolHelper.AddInputDef 2, "SwitchGearWidth", "Width of the SwitchGear", 0.8
    m_oSymbolHelper.AddInputDef 3, "SwitchGearLength", "Length of the SwitchGear", 0.8
    m_oSymbolHelper.AddInputDef 4, "XboltHole", "XboltHole", 0.025
    m_oSymbolHelper.AddInputDef 5, "YboltHole", "YboltHole", 0.5


    'Outputs
    m_oSymbolHelper.NumOutputs = 34
    
    m_oSymbolHelper.AddOutputDef 1, "SwitchGearBody", "SwitchGear Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "ConduitPort1", "ConduitPort 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "ConduitPort2", "ConduitPort 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ConduitPort3", "ConduitPort 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "ConduitPort4", "ConduitPort 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "CablePort1", "CablePort 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "CablePort2", "Cable Port 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "CablePort3", "Cable Port 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "CablePort4", "Cable Port 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "CableTrayPort1", "CableTrayPort 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "CablePort5", "CablePort 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "STFndPort1", "Foundation Port 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "CableCircle1", "CableCircle1", SimplePhysical

    m_oSymbolHelper.AddOutputDef 14, "CenterPos1", "FaceCenter Position1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "CenterPos2", "FaceCenter Position2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "CenterPos3", "FaceCenter Position3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "CenterPos4", "FaceCenter Position4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "CenterPos5", "FaceCenter Position5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "CenterPos6", "FaceCenter Position6", SimplePhysical
    
    m_oSymbolHelper.AddOutputDef 20, "Edge1", "Edge 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Edge2", "Edge 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Edge3", "Edge 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Edge4", "Edge 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "Edge5", "Edge 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "Edge6", "Edge 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "Edge7", "Edge 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "Edge8", "Edge 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "Edge9", "Edge 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "Edge10", "Edge 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 30, "Edge11", "Edge 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 31, "Edge12", "Edge 12", SimplePhysical
    
    m_oSymbolHelper.AddOutputDef 32, "SwitchGearBodyOp", "SwitchGear Body Operation", Operation
    m_oSymbolHelper.AddOutputDef 33, "SwitchGearDoorSwingOp1", "SwitchGear Door Swing Operation CCWise", Operation
    m_oSymbolHelper.AddOutputDef 34, "SwitchGearDoorSwingOp2", "SwitchGear Door Swing Operation CWise", Operation

    'Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical Description", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Operation", "Operation Description", Operation

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




