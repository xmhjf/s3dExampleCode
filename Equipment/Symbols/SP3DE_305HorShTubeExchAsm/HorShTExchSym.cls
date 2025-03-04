VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorShTExchSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:35 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   HorShTExchSym.cls
'   Author: VRK
'   Creation Date:  Monday, Feb 26 2007
'
'   Description:
'   This is E305 Horizontal Shell and Tube Exchanger Assembly with Variable Front and Rear Ends Options.
'   This Symbol Uses four Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                            25
'      2                Insulation                           4
'      3                Maintenance                          1
'      4                ReferenceGeometry                    2
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'******************************************************************************

Option Explicit
Private Const MODULE = "HorShTExchSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_305HorShTubeExchAsm"
    m_oSymbolHelper.ClassName = "HorShTExchSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 40

    m_oSymbolHelper.AddInputDef 1, "FrontEndType", "Front End Type", 1
    m_oSymbolHelper.AddInputDef 2, "RearEndType", "Rear End Type", 1
    m_oSymbolHelper.AddInputDef 3, "ExchangerDiameter", "Exchanger Diameter", 0.5
    m_oSymbolHelper.AddInputDef 4, "ExchangerLength", "Exchanger Length", 2
    m_oSymbolHelper.AddInputDef 5, "ExchangerFlangeTk1", "Exchanger Flange Thickness 1", 0.065
    m_oSymbolHelper.AddInputDef 6, "ExchangerFlangeTk2", "Exchanger Flange Thickness 2", 0.065
    m_oSymbolHelper.AddInputDef 7, "ExpansionJointDiameter", "Expansion Joint Diameter", 0.65
    m_oSymbolHelper.AddInputDef 8, "ExpansionJointPosition", "Expansion Joint Position", 1.5
    m_oSymbolHelper.AddInputDef 9, "ExpansionJointThickness", "Expansion Joint Thickness", 0.2
    m_oSymbolHelper.AddInputDef 10, "BundleFlangeDia", "Bundle Flange Diameter", 0.55
    m_oSymbolHelper.AddInputDef 11, "BundleFlangeTk", "Bundle Flange Thickness", 0.05
    m_oSymbolHelper.AddInputDef 12, "BundlePullingLength", "Bundle Pulling Length", 1.5
    m_oSymbolHelper.AddInputDef 13, "SupportLength", "Support Length", 0.45
    m_oSymbolHelper.AddInputDef 14, "BotSupportCenFromPP", "Bottom Support Center from Placement Point", 1.2
    m_oSymbolHelper.AddInputDef 15, "BottomSupportCentoCen", "Bottom Support Center to Center", 0.8
    m_oSymbolHelper.AddInputDef 16, "BottomSupportHeight", "Bottom Support Height", 0.5
    m_oSymbolHelper.AddInputDef 17, "Support1Thickness", "Support 1 Thickness", 0.15
    m_oSymbolHelper.AddInputDef 18, "Support2Thickness", "Support 2 Thickness", 0.12
    m_oSymbolHelper.AddInputDef 19, "TopSupportHeight", "Top Support Height", 0.5
    m_oSymbolHelper.AddInputDef 20, "TopSupportCenFromPP", "Top Support Center from Placement Point", 1.2
    m_oSymbolHelper.AddInputDef 21, "TopSupportCentoCen", "Top Support Center to Center", 0.8
    m_oSymbolHelper.AddInputDef 22, "FrontEndFlangeDia", "Front End Flange Diameter", 0.55
    m_oSymbolHelper.AddInputDef 23, "FrontEndFlangeTk1", "Front End Flange Thickness 1", 0.065
    m_oSymbolHelper.AddInputDef 24, "FrontEndFlangeTk2", "Front End Flange Thickness 2", 0.05
    m_oSymbolHelper.AddInputDef 25, "FrontEndFlangeTk3", "Front End Flange Thickness 3", 0.065
    m_oSymbolHelper.AddInputDef 26, "FrontEndLength1", "Front End Length 1", 0.5
    m_oSymbolHelper.AddInputDef 27, "FrontEndLength2", "Front End Length 2", 0.3
    m_oSymbolHelper.AddInputDef 28, "RearEndFlangeDia", "Rear End Flange Diameter", 0.7
    m_oSymbolHelper.AddInputDef 29, "RearEndFlangeTk1", "Rear End Flange Thickness 1", 0.05
    m_oSymbolHelper.AddInputDef 30, "RearEndFlangeTk2", "Rear End Flange Thickness 2", 0.065
    m_oSymbolHelper.AddInputDef 31, "RearEndFlangeTk3", "Rear End Flange Thickness 3", 0.06
    m_oSymbolHelper.AddInputDef 32, "RearEndLength", "Rear End Length", 0.6
    m_oSymbolHelper.AddInputDef 33, "ChannelDiameter", "Channel Diameter", 0.3
    m_oSymbolHelper.AddInputDef 34, "FrontEndDiameter", "Front End Diameter", 0.2
    m_oSymbolHelper.AddInputDef 35, "RearEndDiameter", "Rear EndDiameter", 0.2
    m_oSymbolHelper.AddInputDef 36, "InsulationThickness", "Insulation Thickness", 0.025
    m_oSymbolHelper.AddInputDef 37, "FP1XboltHole", "FP1 Bolt X Hole Location", 0.225
    m_oSymbolHelper.AddInputDef 38, "FP1YboltHole", "FP1 Bolt Y Hole Location", 0.075
    m_oSymbolHelper.AddInputDef 39, "FP2XboltHole", "FP2 Bolt X Hole Location", 0.225
    m_oSymbolHelper.AddInputDef 40, "FP2YboltHole", "FP2 Bolt Y Hole Location", 0.06

    'Outputs
    m_oSymbolHelper.NumOutputs = 18

    m_oSymbolHelper.AddOutputDef 1, "ExchangerCyl", "Exchanger Body Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "ExchangerRearFlange", "Exchanger Rear Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "ExchangerFrontFlange1", "Exchanger Front Flange1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ExchangerFrontFlange2", "Exchanger Front Flange 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Expansion_", "Expansion Joint", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Supports_", "Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Point1_", "Top Support Position1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Point2_", "Top Support Position2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Edges_", "Edges of Support1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "EqpFoundationPort1", "EqpFoundation Port1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "EqpFoundationPort2", "EqpFoundation Port2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "ExchangerCylIns", "Exchanger Body Cylinder Insulation", Insulation
    m_oSymbolHelper.AddOutputDef 13, "ExchFrontEndCylIns", "Exchanger Front End Cylinder Insulation", Insulation
    m_oSymbolHelper.AddOutputDef 14, "ExchRearEndCylIns", "Exchanger Rear End Cyl Insulation", Insulation
    m_oSymbolHelper.AddOutputDef 15, "ExpansionJtIns_", "Expansion Joint", Insulation
    m_oSymbolHelper.AddOutputDef 16, "BundlePulling", "BundlePulling", Maintenance
    m_oSymbolHelper.AddOutputDef 17, "DefaultSurface", "Default Surface Horizontal Shell and Tube Exchanger", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 18, "HorShTubeExControlPoint", "Control Point of Horizontal Shell and Tube Exchanger", ReferenceGeometry

    'Aspects
    m_oSymbolHelper.NumAspects = 4
    
    m_oSymbolHelper.AddAspectDef 1, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 2, "Physical", "Physical", SimplePhysical
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
