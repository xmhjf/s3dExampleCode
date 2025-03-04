VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompHorCylSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CompHorCylSym.cls
'   ProgID:         SP3DE_240CompHorCylEqpAsm.CompHorCylSym
'   Author:         VRK
'   Creation Date:  Friday,March 9 2007
'
'   Description:
'   This is E240 Complex Horizontal Cylinder Equipment Assembly.
'   This Assembly Uses three Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                            11
'      2                Insulation                           3
'      3                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CompHorCylSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_240CompHorCylEqpAsm"
    m_oSymbolHelper.ClassName = "CompHorCylSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 25
    
    m_oSymbolHelper.AddInputDef 1, "End1Type", "Vessel Front Head Type", 1
    m_oSymbolHelper.AddInputDef 2, "End2Type", "Vessel Rear Head Type", 1
    m_oSymbolHelper.AddInputDef 3, "VesselTantoTan", "Vessel Tangent to Tangent", 3
    m_oSymbolHelper.AddInputDef 4, "VesselDiameter", "Vessel Diameter", 0.9
    m_oSymbolHelper.AddInputDef 5, "BootHorLocation", "Boot Horizontal Location", 2.175
    m_oSymbolHelper.AddInputDef 6, "BootVertLocation", "Boot Vertical Location", 0.66
    m_oSymbolHelper.AddInputDef 7, "BootDiameter", "Boot Diameter", 0.3
    m_oSymbolHelper.AddInputDef 8, "BootAngularLocation", "Boot Angular Location", 180
    m_oSymbolHelper.AddInputDef 9, "VesselCenterHeight", "Vessel CenterHeight", 0.9
    m_oSymbolHelper.AddInputDef 10, "FirstSupportLocation", "First Support Location", 0.15
    m_oSymbolHelper.AddInputDef 11, "SecondSupportLocation", "Second Support Location", 1.35
    m_oSymbolHelper.AddInputDef 12, "ThirdSupportLocation", "Third Support Location", 1.35
    m_oSymbolHelper.AddInputDef 13, "SupportThickness", "Support Thickness", 0.15
    m_oSymbolHelper.AddInputDef 14, "SupportLength", "Support Length", 0.6
    m_oSymbolHelper.AddInputDef 15, "StiffenerRadius", "Stiffener Radius", 0.6
    m_oSymbolHelper.AddInputDef 16, "EndHeadConeTopDiameter", "End Head Cone Top Diameter", 0.3
    m_oSymbolHelper.AddInputDef 17, "EndHeadConeHeight", "End Head Cone Height", 0.4
    m_oSymbolHelper.AddInputDef 18, "EndHeadKnuckleRadius", "End Head Knuckle Radius", 0.3
    m_oSymbolHelper.AddInputDef 19, "EndHeadDomeradius", "End Head Dome Radius", 0.5
    m_oSymbolHelper.AddInputDef 20, "EndHeadFlangedThick1", "End Head Flanged Thickness 1", 0.2
    m_oSymbolHelper.AddInputDef 21, "EndHeadFlangedThick2", "End Head Flanged Thickness 2", 0.2
    m_oSymbolHelper.AddInputDef 22, "EndHeadSphericalRadius", "End Head Spherical Radius", 0.5
    m_oSymbolHelper.AddInputDef 23, "InsulationThickness", "Insulation Thickness", 0.025
    m_oSymbolHelper.AddInputDef 24, "XboltHole", "Bolt X Hole Location", 0.3
    m_oSymbolHelper.AddInputDef 25, "YboltHole", "Bolt Y Hole Location", 0.075

    'Outputs
    m_oSymbolHelper.NumOutputs = 15
    
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Vessel", "Vessel", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "BootVessel", "Boot vessel", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "BootDome", "Boot Dome", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Support_", "Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Stiffner_", "Stiffners for Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Edges_", "Edges for Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Points_", "Points on Each face of Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "EqpFoundatioPort1", "Foundation Port 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "EqpFoundationPort2", "Foundation Port 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "EqpFoundationPort3", "Foundation Port 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "ComplexHorCylEqpControlPoint", "Control Point of Complex Horizontal Cylinder Vessel", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 13, "InsulatedVessel", "Insulated Vessel", Insulation
    m_oSymbolHelper.AddOutputDef 14, "InsBootVessel", "InsulatedVessel Left Elliptical Head", Insulation
    m_oSymbolHelper.AddOutputDef 15, "InsBootDome", "InsulatedVessel Right Elliptical Head", Insulation
    
    'Aspects
    m_oSymbolHelper.NumAspects = 3
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "RefGeometry", "ReferenceGeometry", ReferenceGeometry
   
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




