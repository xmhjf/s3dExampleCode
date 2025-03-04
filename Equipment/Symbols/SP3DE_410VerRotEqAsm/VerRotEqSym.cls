VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VerRotEqSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   VerRotEqSym.cls
'   ProgID:         SP3DE_410VerRotEqAsm.VerlRotatingEqpSym
'   Author:         RUK
'   Creation Date:  Monday, Feb 19 2007
'
'   Description:
'   This is E410 Vertical Rotating Equipment Assembly.
'   This Assembly Uses four Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                            13
'      2                Maintenance                          1
'      3                Insulation                           3
'      4                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "VerRotEqSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_410VerRotEqAsm"
    m_oSymbolHelper.ClassName = "VerRotEqSym"
    
    'Inputs
    m_oSymbolHelper.NumInputs = 10
    
    m_oSymbolHelper.AddInputDef 1, "BottomBodyHeight1", "Bottom Body Height1", 0.1
    m_oSymbolHelper.AddInputDef 2, "BottomBodyHeight2", "Bottom Body Height2", 0.1
    m_oSymbolHelper.AddInputDef 3, "BottomBodyDiameter", "Bottom Body Diameter", 0.1
    m_oSymbolHelper.AddInputDef 4, "ShaftLength", "Shaft Length", 1.5
    m_oSymbolHelper.AddInputDef 5, "ShaftDiameter", "Shaft Diameter", 0.45
    m_oSymbolHelper.AddInputDef 6, "TaperBodyHeight", "Taper Body Height", 0.1
    m_oSymbolHelper.AddInputDef 7, "TopBodyHeight", "Top Body Height", 0.3
    m_oSymbolHelper.AddInputDef 8, "TopBodyDiameter", "Top Body Diameter", 1.5
    m_oSymbolHelper.AddInputDef 9, "BundlePullingLength", "Bundle Pull Length", 0.1
    m_oSymbolHelper.AddInputDef 10, "InsulationThickness", "InsulationThickness", 0.02
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 18
    
    m_oSymbolHelper.AddOutputDef 1, "InsTopBody", "Insulation for Top Body", Insulation
    m_oSymbolHelper.AddOutputDef 2, "InsTaperBody", "Insulation for Taper Body", Insulation
    m_oSymbolHelper.AddOutputDef 3, "InsBottomBody", "Insulation for Bottom Body", Insulation
    m_oSymbolHelper.AddOutputDef 4, "MaintEquipment", "Maintenance for Equipment", Maintenance
    m_oSymbolHelper.AddOutputDef 5, "DefaultSurface", "Surface at Bottom", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "TopSurface", "Surface at Top", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Edge1", "Top Edge of Top Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Edge2", "Bottom Edge of Top Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Edge3", "Top Edge of Bottom Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Edge4", "Bottom Edge of Bottom Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Edge5", "Top Edge of the Shaft", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "Edge6", "Bottom Edge of the Shaft", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "TopBody", "Top Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "TaperBody", "Taper Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "BottomBody", "Bottom Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "EqpFoundationPort", "Foundation Port under Bottom Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "Shaft", "Shaft", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "VerRotatingEquipmentCompControlPoint", "Control Point of Vertiacl Rotating Equipment Component", ReferenceGeometry
        
   'Aspects
    m_oSymbolHelper.NumAspects = 4
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "Maintenance", "Maintenance", Maintenance
    m_oSymbolHelper.AddAspectDef 4, "RefGeometry", "ReferenceGeometry", ReferenceGeometry
    
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
