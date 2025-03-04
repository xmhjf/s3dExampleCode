VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VerShTExchSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:49 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   VerRotEqSym.cls
'   ProgID:         SP3DE_310VerShTubeExchAsm.VerShTExchSym
'   Author:         RUK
'   Creation Date:  Monday, Mar 5 2007
'
'   Description:
'   This is E310 Vertical Shell Tube Exchanger Equipment Assembly.
'   This Assembly Uses four Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             5
'      2                Maintenance                          1
'      3                Insulation                           2
'      4                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "VerShTExchSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_310VerShTubeExchAsm"
    m_oSymbolHelper.ClassName = "VerShTExchSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 34
    
    m_oSymbolHelper.AddInputDef 1, "FrontEndType", "Front End Type", 1
    m_oSymbolHelper.AddInputDef 2, "RearEndType", "Rear End Type", 1
    m_oSymbolHelper.AddInputDef 3, "ExchangerDiameter", "Exchanger Diameter", 0.4
    m_oSymbolHelper.AddInputDef 4, "ExchangerLength", "Exchanger Length", 4
    m_oSymbolHelper.AddInputDef 5, "ExchangerFlangeTk1", "Exchanger Flange Thickness 1", 0.065
    m_oSymbolHelper.AddInputDef 6, "ExchangerFlangeTk2", "Exchanger Flange Thickness 2", 0.065
    m_oSymbolHelper.AddInputDef 7, "ExpansionJointDiameter", "Expansion Joint Diameter", 0.5
    m_oSymbolHelper.AddInputDef 8, "ExpansionJointPosition", "Expansion Joint Position", 4.5
    m_oSymbolHelper.AddInputDef 9, "ExpansionJointThickness", "Expansion Joint Thickness", 0.25
    m_oSymbolHelper.AddInputDef 10, "BundleFlangeDia", "Bundle Flange Diameter", 0.45
    m_oSymbolHelper.AddInputDef 11, "BundleFlangeTk", "SBundle Flange Thickness", 0.05
    m_oSymbolHelper.AddInputDef 12, "BundlePullingLength", "Bundle Pulling Length", 4
    m_oSymbolHelper.AddInputDef 13, "SupportAngularLocation", "Support  Angular Location", 45
    m_oSymbolHelper.AddInputDef 14, "NumberOfSupports", "Number Of Supports", 4
    m_oSymbolHelper.AddInputDef 15, "SupportDepth", "Support Depth", 0.15
    m_oSymbolHelper.AddInputDef 16, "SupportFromOrigin", "Support From Origin", 3.25
    m_oSymbolHelper.AddInputDef 17, "SupportTopWidth", "Support Top Width", 0.5
    m_oSymbolHelper.AddInputDef 18, "SupportBottomWidth", "Support Bottom Widt", 0.65
    m_oSymbolHelper.AddInputDef 19, "SupportThickness", "Support Thickness", 0.2
    m_oSymbolHelper.AddInputDef 20, "FrontEndFlangeDia", "Front End Flange Diameter", 0.5
    m_oSymbolHelper.AddInputDef 21, "FrontEndFlangeTk1", "Front End Flange Thickness 1", 0.05
    m_oSymbolHelper.AddInputDef 22, "FrontEndFlangeTk2", "Front End Flange Thickness 2", 0.05
    m_oSymbolHelper.AddInputDef 23, "FrontEndFlangeTk3", "Front End Flange Thickness 3", 0.065
    m_oSymbolHelper.AddInputDef 24, "FrontEndLength1", "Front End Length 1", 2.5
    m_oSymbolHelper.AddInputDef 25, "FrontEndLength2", "Front End Length 2", 2
    m_oSymbolHelper.AddInputDef 26, "RearEndFlangeDia", "Rear End Flange Diameter", 0.7
    m_oSymbolHelper.AddInputDef 27, "RearEndFlangeTk1", "Rear End Flange Thickness 1", 0.05
    m_oSymbolHelper.AddInputDef 28, "RearEndFlangeTk2", "Rear End Flange Thickness 2", 0.065
    m_oSymbolHelper.AddInputDef 29, "RearEndFlangeTk3", "Rear End Flange Thickness 3", 0.06
    m_oSymbolHelper.AddInputDef 30, "RearEndLength", "Rear End Length", 2.5
    m_oSymbolHelper.AddInputDef 31, "ChannelDiameter", "Channel Diameter", 0.3
    m_oSymbolHelper.AddInputDef 32, "FrontEndDiameter", "Front End Diameter", 0.2
    m_oSymbolHelper.AddInputDef 33, "RearEndDiameter", "Rear End Diameter", 0.2
    m_oSymbolHelper.AddInputDef 34, "InsulationThickness", "Insulation Thickness", 0.025
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 9
    
    m_oSymbolHelper.AddOutputDef 1, "VerShTubeExControlPoint", "Control Point of Vert Shell Tube Exchanger Component", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 2, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "ExchangerBody", "Exchanger Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ExpansionJoint", "Expansion Joint", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Support_", "Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Surface_", "Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "ExchanBodyIns", "Exchanger Body Ins", Insulation
    m_oSymbolHelper.AddOutputDef 8, "ExpansionJointIns", "Expansion Joint Ins", Insulation
    m_oSymbolHelper.AddOutputDef 9, "BundlePulling", "Bundle Pulling Cylinder", Maintenance
    
    'Aspects
    m_oSymbolHelper.NumAspects = 4
    
    m_oSymbolHelper.AddAspectDef 1, "RefGeometry", "ReferenceGeometry", ReferenceGeometry
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 4, "Maintenance", "Maintenance", Maintenance
    
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
