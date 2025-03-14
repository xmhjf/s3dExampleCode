VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MFldLightSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008-10, Intergraph Corporation. All rights reserved.
'
'   MFldLightSym.cls
'   ProgID:        SP3DHighMastFloodLightAsm.MFldLightSym
'   Author:        VRK
'   Creation Date:  FriDay,June 6 2008
'
'   Description:
'   This is High Mast Flood Light Assembly.
'
'   This Assembly Uses One Aspect
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             11
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.June.2008    VRK     CR-141836:Provide new electrical equipment symbols for lighting fixtures
'   03.Feb.2010     PRB     TR-CP-157988  TO Do List entries are generated whern placing lighting fixtures
'                            (Modified nozzle code to create with placeholder)
'   25.Mar.2010     RUK     CR-CP-177179  Enhance lighting symbols to provide control point for dimensioning purposes
'                           CR-CP-113328  "can be modified" flag should be True for all best practice equipment symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "MFldLightSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHighMastFloodLightAsm"
    m_oSymbolHelper.ClassName = "MFldLightSym"

' Inputs
    m_oSymbolHelper.NumInputs = 16
    
    m_oSymbolHelper.AddInputDef 1, "PoleHeight", "A-Pole Height", 2.1
    m_oSymbolHelper.AddInputDef 2, "PoleDiameter1", "D1-Pole Diameter 1", 0.07
    m_oSymbolHelper.AddInputDef 3, "PoleDiameter2", "D2-Pole Diameter 2", 0.05
    m_oSymbolHelper.AddInputDef 4, "MountPlateLength", "PL-Mount Plate Length", 0.2
    m_oSymbolHelper.AddInputDef 5, "MountPlateWidth", "PW-Mount Plate Width", 0.25
    m_oSymbolHelper.AddInputDef 6, "MountPlateThickness", "PT-Mount Plate Thickness", 0.025
    m_oSymbolHelper.AddInputDef 7, "FixtureHeight", "P1-Fixture Height", 0.4
    m_oSymbolHelper.AddInputDef 8, "FixtureWidth", "P2-Fixture Width", 0.075
    m_oSymbolHelper.AddInputDef 9, "FixtureLength", "P3-Fixture Length", 0.4
    m_oSymbolHelper.AddInputDef 10, "BoxHeight", "P4-Box Height", 0.4
    m_oSymbolHelper.AddInputDef 11, "BoxWidth", "P5-Box Width", 0.2
    m_oSymbolHelper.AddInputDef 12, "BoxLength", "P6-Box Length", 0.4
    m_oSymbolHelper.AddInputDef 13, "Wattage", "Wattage", 100
    m_oSymbolHelper.AddInputDef 14, "CPx", "P5-Box Width", 0#
    m_oSymbolHelper.AddInputDef 15, "CPy", "P6-Box Length", 0#
    m_oSymbolHelper.AddInputDef 16, "CPz", "Wattage", 0#
' Outputs
    m_oSymbolHelper.NumOutputs = 12
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "MountPlate", "Mount Plate Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "PoleBody1", "Pole Body 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "PoleBody2 ", "Pole Body 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "PoleBody3", "Pole Body 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "PoleBody4", "Pole Body 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Light1", "Light Fixture 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Light2", "Light Fixture 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Light3", "Light Fixture 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "ConduitNozz", "Conduit Nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "CableNozz", "Cable Nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "EqpControlPoint", "Contol Point at the Origin", ReferenceGeometry

' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "RefGeometry", "ReferenceGeometry", ReferenceGeometry
    
    Exit Sub

ErrorHandler:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

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


