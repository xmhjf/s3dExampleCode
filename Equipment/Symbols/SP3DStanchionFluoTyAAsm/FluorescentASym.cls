VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FluorescentASym"
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
'   FluorescentASym.cls
'   ProgID:        SP3DStanchionFluoTyAAsm.FluorescentASym
'   Author:        VRK
'   Creation Date: FriDay,May 16 2008
'   Description:
'   This is Fluorescent Handrail Mounted Lighting Fixture Assembly
'   This Assembly Uses One Aspect
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             10
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   16.May.2008     VRK     CR-141836:Provide new electrical equipment symbols for lighting fixtures
'   03.Feb.2010     VKC     TR-CP-157988  TO Do List entries are generated whern placing lighting fixtures
'                            (Modified nozzle code to create with placeholder)
'   25.Mar.2010     RUK     CR-CP-177179  Enhance lighting symbols to provide control point for dimensioning purposes
'                           CR-CP-113328  "can be modified" flag should be True for all best practice equipment symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "FluorescentASym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DStanchionFluoTyAAsm"
    m_oSymbolHelper.ClassName = "FluorescentASym"

' Inputs
    m_oSymbolHelper.NumInputs = 10
    m_oSymbolHelper.AddInputDef 1, "PoleHeight", "A-Pole Height", 4
    m_oSymbolHelper.AddInputDef 2, "PoleDiameter", "PD-Pole Diameter", 0.0625
    m_oSymbolHelper.AddInputDef 3, "Angle", "C-Angle", Atn(1)
    m_oSymbolHelper.AddInputDef 4, "Wattage", "Wattage", 40
    m_oSymbolHelper.AddInputDef 5, "Length", "L-Length", 1.8
    m_oSymbolHelper.AddInputDef 6, "Width", "W-Width", 0.2
    m_oSymbolHelper.AddInputDef 7, "Depth", "D-Depth", 0.4
    m_oSymbolHelper.AddInputDef 8, "CPx", "P5-Box Width", 0#
    m_oSymbolHelper.AddInputDef 9, "CPy", "P6-Box Length", 0#
    m_oSymbolHelper.AddInputDef 10, "CPz", "Wattage", 0#

' Outputs
    m_oSymbolHelper.NumOutputs = 11
    
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "PoleBody", "Pole Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "PoleCurveBody", "Pole Curve Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Cylinder1atTop", "Cylinder 1 atTop", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Cylinder2atTop", "Cylinder 2 atTop", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Box", "Bottom Curve at Light Position", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "CurveObject", "Center Curve at Light Position", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "TopBox", "Top Curve at Light Position", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "ConduitNozzle", "Conduit Nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "CableNozzle", "Cable Nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "EqpControlPoint", "Contol Point at the Origin", ReferenceGeometry

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
