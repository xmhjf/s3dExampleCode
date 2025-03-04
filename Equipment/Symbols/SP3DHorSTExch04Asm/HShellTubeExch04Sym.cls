VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HShellTubeExch04Sym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:19 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   HShellTubeExch04Sym.cls
'   ProgID:         SP3DHorSTExch04Asm.HShellTubeExch04Sym
'   Author:         svsmylav
'   Creation Date:  Friday, Apr 25, 2003
'
'   This is a Shell and Tube Exchanger(E-182).
'   There are 2 Reference Geometry objects which are Control-Point at origin and a reference plane at support bottom.
'   This is a Dow symbol Equipment
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "HShellTubeExch04Sym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices


Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHorSTExch04Asm"
    m_oSymbolHelper.ClassName = "HShellTubeExch04Sym"
    
' Inputs
    m_oSymbolHelper.NumInputs = 16
    m_oSymbolHelper.AddInputDef 1, "VesselDiameter", "Exchanger Diameter", 0.8382
    m_oSymbolHelper.AddInputDef 2, "VesselTantoTan", "Exchanger  Tangent to Tanget Length", 5.0292
    m_oSymbolHelper.AddInputDef 3, "VesselLength", "Front Head Length", 0.8128
    m_oSymbolHelper.AddInputDef 4, "VesselStartPoint", "Distance of Control Point from Tangent", 0.4064
    m_oSymbolHelper.AddInputDef 5, "FaceDiameter", "Flange Diameter", 0.9398
    m_oSymbolHelper.AddInputDef 6, "FaceThickness", "Flange Thickness", 0.0508
    m_oSymbolHelper.AddInputDef 7, "SupportWidth", "Support Width", 0.1524
    m_oSymbolHelper.AddInputDef 8, "SupportLength", "Support Length", 0.7874
    m_oSymbolHelper.AddInputDef 9, "SupportHeight", "Support Height", 1.6129
    m_oSymbolHelper.AddInputDef 10, "FirstSupportLocation", "First Support Location", 1.016
    m_oSymbolHelper.AddInputDef 11, "SecondSupportLocation", "Second Support Location", 3.9624
    m_oSymbolHelper.AddInputDef 12, "BasePlateThickness", "Base Plate Thickness(E)", 0.016
    m_oSymbolHelper.AddInputDef 13, "BackingPlateThickness", "Backing Plate Thickness(T)", 0.012
    m_oSymbolHelper.AddInputDef 14, "BackingPlateWidth", "Backing Plate Widt(H)", 0.3048
    m_oSymbolHelper.AddInputDef 15, "StiffnerThickness", "Stiffner Thickness(D)", 0.016
    m_oSymbolHelper.AddInputDef 16, "StiffnerThickness1", "Stiffner Thickness(DG)", 0.012

' Outputs
    m_oSymbolHelper.NumOutputs = 20
    m_oSymbolHelper.AddOutputDef 1, "FrontHeadBonnet", "Elliptical Front Head Bonnet Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "FrontHeadBodyCyl", "Cylindrical Front Head Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "FrontHeadFlangeTop", "Top Flange on Front head side", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "FrontHeadFlangeBottom", "Bottom Flange on Front head side", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "ShellBodyCyl", "Cylindrical Exchanger Shell Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "RearHeadBonnet", "Elliptical Front Head Bonnet Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Support1SaddlePlate", "Support1 Saddle Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Support1StiffnerPlate", "Support1 Ends Support Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Support1BackingPlate1", "Support1 Center Support Plate1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Support1BackingPlate2", "Support1 Middle Support Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Support1BackingPlate3", "Support1 Center Support Plate2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "Support1BackingPlate4", "Support1 Bottom Support Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "Support2SaddlePlate", "Support2 Saddle Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "Support2StiffnerPlate", "Support2 Ends Support Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "Support2BackingPlate1", "Support2 Center Support Plate1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "Support2BackingPlate2", "Support2 Middle Support Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "Support2BackingPlate3", "Support2 Center Support Plate2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "Support2BackingPlate4", "Support2 Bottom Support Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "ControlPoint", "ControlPoint at origin", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 20, "DefaultSurface", "Default Surface  at Support", ReferenceGeometry

' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "ReferenceGeometry", "ReferenceGeometry", ReferenceGeometry
    
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






