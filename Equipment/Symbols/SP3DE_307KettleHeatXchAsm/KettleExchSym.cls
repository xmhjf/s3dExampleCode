VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KettleExchSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   KettleExchSym.cls
'   ProgID:         SP3DE_307KettleHeatXchAsm.KettleExchSym
'   Author:         VRK
'   Creation Date:  Monday,March 5 2007
'
'   Description:
'   This is E307 Kettle Heat Exchanger Equipment Assembly.
'   This Assembly Uses four Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                            12
'      2                Maintenance                          1
'      3                Insulation                           3
'      4                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "KettleExchSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_307KettleHeatXchAsm"
    m_oSymbolHelper.ClassName = "KettleExchSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 36

    m_oSymbolHelper.AddInputDef 1, "FrontEndType", "Front End Type", 1
    m_oSymbolHelper.AddInputDef 2, "VesselHeadType", "Rear End Type", 1
    m_oSymbolHelper.AddInputDef 3, "ExchangerLength", "Exchanger Length", 4#
    m_oSymbolHelper.AddInputDef 4, "ExchangerNeckLength", "Exchanger Neck Length ", 0.5
    m_oSymbolHelper.AddInputDef 5, "ExchangerNeckDiameter", "Exchanger Neck Diameter", 0.2
    m_oSymbolHelper.AddInputDef 6, "ExchangerTaperLength", "Exchanger Taper Length", 0.6
    m_oSymbolHelper.AddInputDef 7, "ExchangerDiameter", "Exchanger Diameter", 0.5
    m_oSymbolHelper.AddInputDef 8, "ExchangerFlangeTk1", "Exchanger Flange Thickness 1", 0.065
    m_oSymbolHelper.AddInputDef 9, "BundleFlangeDia", "Bundle Flange Diameter", 0.55
    m_oSymbolHelper.AddInputDef 10, "BundleFlangeTk", "Bundle Flange Thickness", 0.05
    m_oSymbolHelper.AddInputDef 11, "BundlePullingLength", "Bundle Pulling Length", 3#
    m_oSymbolHelper.AddInputDef 12, "SupportLength", "Support Length", 0.45
    m_oSymbolHelper.AddInputDef 13, "BotSupportCenFromPP", "Bottom Support Center from Placement Point", 1.55
    m_oSymbolHelper.AddInputDef 14, "BottomSupportCentoCen", "Bottom Support Center to Center", 1.5
    m_oSymbolHelper.AddInputDef 15, "BottomSupportHeight", "Bottom Support Height", 0.5
    m_oSymbolHelper.AddInputDef 16, "Support1Thickness", "Support 1 Thickness", 0.15
    m_oSymbolHelper.AddInputDef 17, "Support2Thickness", "Support 2 Thickness", 0.12
    m_oSymbolHelper.AddInputDef 18, "FrontEndFlangeDia", "Front End Flange Diameter", 0.55
    m_oSymbolHelper.AddInputDef 19, "FrontEndFlangeTk1", "Front End Flange Thickness 1", 0.065
    m_oSymbolHelper.AddInputDef 20, "FrontEndFlangeTk2", "Front End Flange Thickness 2", 0.05
    m_oSymbolHelper.AddInputDef 21, "FrontEndFlangeTk3", "Front End Flange Thickness 3", 0.065
    m_oSymbolHelper.AddInputDef 22, "FrontEndLength1", "Front End Length 1", 0.75
    m_oSymbolHelper.AddInputDef 23, "FrontEndLength2", "Front End Length 2", 0.3
    m_oSymbolHelper.AddInputDef 24, "FrontEndDiameter", "Front End Diameter", 0.2
    m_oSymbolHelper.AddInputDef 25, "EndHeadConeTopDiameter", "End Head Cone Top Diameter", 0.1
    m_oSymbolHelper.AddInputDef 26, "EndHeadConeHeight", "End Head Cone Height", 0.3
    m_oSymbolHelper.AddInputDef 27, "EndHeadKnuckleRadius", "End Head Knuckle Radius", 0.3
    m_oSymbolHelper.AddInputDef 28, "EndHeadDomeradius", "End Head Dome Radius", 0.3
    m_oSymbolHelper.AddInputDef 29, "EndHeadFlangedThick1", "End Head Flanged Thickness 1", 0.1
    m_oSymbolHelper.AddInputDef 30, "EndHeadFlangedThick2", "End Head Flanged Thickness 2", 0.1
    m_oSymbolHelper.AddInputDef 31, "EndHeadSphericalRadius", "End Head Spherical Radius", 0.5
    m_oSymbolHelper.AddInputDef 32, "InsulationThickness", "Insulation Thickness", 0.025
    m_oSymbolHelper.AddInputDef 33, "FP1XboltHole", "FP1 Bolt X Hole Location", 0.225
    m_oSymbolHelper.AddInputDef 34, "FP1YboltHole", "FP1 Bolt Y Hole Location", 0.075
    m_oSymbolHelper.AddInputDef 35, "FP2XboltHole", "FP2 Bolt X Hole Location", 0.225
    m_oSymbolHelper.AddInputDef 36, "FP2YboltHole", "FP2 Bolt Y Hole Location", 0.06

    'Outputs
    m_oSymbolHelper.NumOutputs = 16
    
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface of Kettle Exchange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "ExchangerBody", "Exchanger Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "ExTaperBody", "Exchanger Taper Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ExneckBody", "Exchanger Neck Portion", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "ExchangerFlange", "Exchanger Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "BundleFlange", "Bundle Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Support1", "Support 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Support2", "Support 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Edges_", "Edges for Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "EqpFoundationPort1", "Foundation Port 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "EqpFoundationPort2", "Foundation Port 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "ExBodyInsul", "ExchangerBody Insulation", Insulation
    m_oSymbolHelper.AddOutputDef 13, "ExTapBodyInsul", "Exchanger Taper Body Insulation", Insulation
    m_oSymbolHelper.AddOutputDef 14, "ExneckBodyInsul", "Exchanger Neck Portion Insulation", Insulation
    m_oSymbolHelper.AddOutputDef 15, "BundlePullCylin", "Bundle Pulling Cylinder", Maintenance
    m_oSymbolHelper.AddOutputDef 16, "KettleExchangerControlPoint", "Control Point of Kettle Exchanger", ReferenceGeometry
    
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
