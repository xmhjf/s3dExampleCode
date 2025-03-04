VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorCylSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:58 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   SimpHorCylVesselSym.cls
'   ProgID:         SP3DE_245SimpHorCylEqpAsm.HorCylSym
'   Author:         RUK
'   Creation Date:  Wednesday, March 7 2007
'
'   Description:
'   This is E245 Simple Horizontal Cylindrical Equipment Assembly.
'   This Assembly Uses four Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             8
'      2                Insulation                           1
'      3                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "HorCylSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_245SimpHorCylEqpAsm"
    m_oSymbolHelper.ClassName = "HorCylSym"

    'Inputs
    
    m_oSymbolHelper.NumInputs = 19
    
    m_oSymbolHelper.AddInputDef 1, "End1Type", "Vessel Front Head Type", 1
    m_oSymbolHelper.AddInputDef 2, "End2Type", "Vessel Rear Head Type", 1
    m_oSymbolHelper.AddInputDef 3, "VesselTantoTan", "Vessel Tangent to Tangent", 4
    m_oSymbolHelper.AddInputDef 4, "VesselDiameter", "Vessel Diameter", 1.5
    m_oSymbolHelper.AddInputDef 5, "VesselCenterHeight", "Vessel CenterHeight", 1.75
    m_oSymbolHelper.AddInputDef 6, "FirstSupportLocation", "First Support Location", 0.75
    m_oSymbolHelper.AddInputDef 7, "SecondSupportLocation", "Second Support Location", 2.5
    m_oSymbolHelper.AddInputDef 8, "SupportThickness", "Support Thickness", 0.1
    m_oSymbolHelper.AddInputDef 9, "SupportLength", "Support Length", 1
    m_oSymbolHelper.AddInputDef 10, "XBoltHole", "Hole Bolt X Location", 0.5
    m_oSymbolHelper.AddInputDef 11, "YBoltHole", "Hole Bolt Y Location", 0.05
    m_oSymbolHelper.AddInputDef 12, "EndHeadConeTopDiameter", "EndHead Cone Top Diameter", 0.1
    m_oSymbolHelper.AddInputDef 13, "EndHeadConeHeight", "End Head Cone Height", 0.4
    m_oSymbolHelper.AddInputDef 14, "EndHeadKnuckleRadius", "End Head Knuckle Radius", 0.25
    m_oSymbolHelper.AddInputDef 15, "EndHeadDomeradius", "End Head Dome Radius", 1.2
    m_oSymbolHelper.AddInputDef 16, "EndHeadFlangedThick1", "End Head Flanged Thickness 1", 0.6
    m_oSymbolHelper.AddInputDef 17, "EndHeadFlangedThick2", "End Head Flanged Thickness 2", 0.5
    m_oSymbolHelper.AddInputDef 18, "EndHeadSphericalRadius", "End Head Spherical Radius", 0.85
    m_oSymbolHelper.AddInputDef 19, "InsulationThickness", "Insulation Thickness", 0.025
        
    'Outputs
    m_oSymbolHelper.NumOutputs = 10
    
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Vessel", "Vessel", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Support_", "Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Surface_", "Surfaces for Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Edges_", "Edges for Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Points_", "Points on Each face of Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "EqpFoundatioPort1", "Foundation Port 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "EqpFoundationPort2", "Foundation Port 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "SimpleHorCylVesselControlPoint", "Control Point of Simple Horizontal Cylinder Vessel Component", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 10, "InsulatedVessel", "Insulated Vessel", Insulation
    
    'Aspects
    m_oSymbolHelper.NumAspects = 3
    
    m_oSymbolHelper.AddAspectDef 1, "RefGeometry", "ReferenceGeometry", ReferenceGeometry
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "Physical", "Physical", SimplePhysical
    
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


