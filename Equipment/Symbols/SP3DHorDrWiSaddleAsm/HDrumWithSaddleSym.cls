VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HDrumWithSaddleSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:11 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   HDrumWithSaddleSym.cls
'   ProgID:         SP3DHorDrWiSaddleAsm.HDrumWithSaddleSym
'   Author:         svsmylav
'   Creation Date:  Monday, May 26 2003
'   Description:
'   This symbol, Horizontal Drum , is prepared as per DOW Chemical Company manual.
'   All assumed values are as per the DOW documents unless otherwise specified. For more details
'   refer DOW document D-162.
'
'   Change History:
'   dd.mmm.yyyy     who     change description

'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "HDrumWithSaddleSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    Dim index As Long
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHorDrWiSaddleAsm"
    m_oSymbolHelper.ClassName = "HDrumWithSaddleSym"
    index = 0

' Inputs
    m_oSymbolHelper.NumInputs = 16
    '5
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "VesselDiameter", "Vessel Internal Diameter", 0.9
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "VesselTantoTan", "Vessel Tangent to Tangent", 2.7
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "FirstSupportLocation", "First Support Location ", 0.28
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "SecondSupportLocation", "Second Support Location", 1.93
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "InsulationThickness", "Insulation Thickness", 0.025
    
    'Saddle Input parameters
    '11
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "SupportLength", "Support Length(B)", 1.2
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "SupportHeight", "Support Height From Vessel Center to base of saddle", 1.2
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "SupportWidth", "Support Width(G)", 0.15
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "BasePlateThickness", "Base Plate Thickness(E)", 0.016
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "BackingPlateThickness", "Backing Plate Thickness(T)", 0.012
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "BackingPlateWidth", "Backing Plate Widt(H)", 0.25
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "StiffnerThickness", "Stiffner Thickness(D)", 0.016
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "StiffnerThickness1", "Stiffner Thickness(DG)", 0.012
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "StiffnerSpacing", "Stiffner Spacing(C)", 0.406
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "XboltHole", "Bolt Hole Location-X (F)", 0.101
    index = index + 1
    m_oSymbolHelper.AddInputDef index, "YboltHole", "Bolt Hole Location-Y (A)", 1.01

    
    
' Outputs
    index = 0
'Drum Body outputs
    m_oSymbolHelper.NumOutputs = 25
    '3
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "DrumBody", "Drum body as Cylinder", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "DrumHead", "Drum head", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "DrumHead1", "Drum Head 1", SimplePhysical
    
'Saddle 1 Support outputs
    '18
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "SaddlePlate", "Saddle Body", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "SaddlePlate1", "Second Saddle Body", SimplePhysical

    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "StiffnerPlate", "Ends Support Plate", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "StiffnerPlate1", "Second saddle Ends Support Plate", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate1", "Center Support Plate1", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate2", "Second Saddle Center Support Plate", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate3", "Center Support Plate2", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate4", "Saddle Saddle Center Support Plate2", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate5", "Middle Support Plate", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate6", "Second saddle Middle Support Plate", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate7", "Center Support Plate3", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate8", "Second saddle Center Support Plate3", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate9", "Center Support Plate4", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate10", "Second saddle Center Support Plate4", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate11", "Bottom Support Plate", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "BackingPlate12", "Second saddle Bottom Support Plate", SimplePhysical
    
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "EqpFoundationPort", "Foundation Port", SimplePhysical
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "EqpFoundationPort1", "Second Foundation Port", SimplePhysical
    

'Refernce Geometry Outputs
    '3
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "ControlPoint", "Point at origin", ReferenceGeometry
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "DefaultSurface", "Default Surface", ReferenceGeometry
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "HoriLine", "Horizontal Line", ReferenceGeometry
    
'Insulation Aspect
    '1
    index = index + 1
    m_oSymbolHelper.AddOutputDef index, "DrumIns", " Drum Insulation", Insulation
    
' Aspects
    m_oSymbolHelper.NumAspects = 3
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "ReferenceGeometry", "ReferenceGeometry", ReferenceGeometry
        
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




