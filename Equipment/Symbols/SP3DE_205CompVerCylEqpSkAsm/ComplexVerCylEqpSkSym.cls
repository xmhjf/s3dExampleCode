VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVCESym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:59 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   ComplexVerCylEqpSkSym.cls
'   ProgID:         SP3DE_205CompVerCylEqpSkAsm.CVCESym
'   Author:         RUK
'   Creation Date:  Monday, March 26 2007
'
'   Description:
'   This is E205 Complex Vertical Cylindrical Equipment With Skirt Assembly.
'   This Assembly Uses four Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             11
'      2                Insulation                           7
'      3                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CVCESym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_205CompVerCylEqpSkAsm"
    m_oSymbolHelper.ClassName = "CVCESym"

    'Inputs
    m_oSymbolHelper.NumInputs = 40
    
    m_oSymbolHelper.AddInputDef 1, "VesselHeight1", "Vessel Height 1", 1
    m_oSymbolHelper.AddInputDef 2, "VesselDiameter1", "Vessel Diameter 1", 1
    m_oSymbolHelper.AddInputDef 3, "SupportHeight1", "Support Height 1", 1
    m_oSymbolHelper.AddInputDef 4, "VesselHeight2", "Vessel Height 2", 1.5
    m_oSymbolHelper.AddInputDef 5, "VesselDiameter2", "Vessel Diameter 2", 2
    m_oSymbolHelper.AddInputDef 6, "SupportHeight2", "Support Height 2", 1.5
    m_oSymbolHelper.AddInputDef 7, "VesselHeight3", "Vessel Height 3", 2.5
    m_oSymbolHelper.AddInputDef 8, "VesselDiameter3", "Vessel Diameter 3", 3.5
    m_oSymbolHelper.AddInputDef 9, "SupportHeight3", "Support Height 3", 2
    m_oSymbolHelper.AddInputDef 10, "VesselHeight4", "Vessel Height 4", 1.5
    m_oSymbolHelper.AddInputDef 11, "VesselDiameter4", "Vessel Diameter 4", 2.7
    m_oSymbolHelper.AddInputDef 12, "SkirtHeight", "Skirt Height", 2.5
    m_oSymbolHelper.AddInputDef 13, "SkirtTopDiameter", "Skirt Top Diameter", 2.5
    m_oSymbolHelper.AddInputDef 14, "SkirtBottomDiameter", "Skirt Bottom Diameter", 4
    m_oSymbolHelper.AddInputDef 15, "VesselReferencePoint", "Vessel Reference Point", 4
    m_oSymbolHelper.AddInputDef 16, "End1Type", "Vessel Front Head Type", 1
    m_oSymbolHelper.AddInputDef 17, "End2Type", "Vessel Rear Head Type 1", 1
    m_oSymbolHelper.AddInputDef 18, "End3Type", "Vessel Rear Head Type 2", 1
    m_oSymbolHelper.AddInputDef 19, "EndHead1ConeTopDiameter", "End Head1 Cone Top Diameter", 0.5
    m_oSymbolHelper.AddInputDef 20, "EndHead1ConeHeight", "End Head1 Cone Height", 0.5
    m_oSymbolHelper.AddInputDef 21, "EndHead1KnuckleRadius", "End Head1 Knuckle Radius", 0.35
    m_oSymbolHelper.AddInputDef 22, "EndHead1Domeradius", "End Head1 Dome Radius", 1.775
    m_oSymbolHelper.AddInputDef 23, "EndHead1FlangedThick1", "End Head1 Flanged Thick1", 0.2
    m_oSymbolHelper.AddInputDef 24, "EndHead1FlangedThick2", "End Head1 Flanged Thick2", 0.2
    m_oSymbolHelper.AddInputDef 25, "EndHead1SphericalRadius", "End Head1 Spherical Radius", 1.8
    m_oSymbolHelper.AddInputDef 26, "EndHead2ConeTopDiameter", "End Head2 Cone Top Diameter", 2
    m_oSymbolHelper.AddInputDef 27, "EndHead2ConeHeight", "End Head2 Cone Height", 1
    m_oSymbolHelper.AddInputDef 28, "EndHead2KnuckleRadius", "End Head2 Knuckle Radius", 0.7
    m_oSymbolHelper.AddInputDef 29, "EndHead2Domeradius", "End Head2 Dome Radius", 2
    m_oSymbolHelper.AddInputDef 30, "EndHead2FlangedThick1", "End Head2 Flanged Thick1", 0.4
    m_oSymbolHelper.AddInputDef 31, "EndHead2FlangedThick2", "End Head2 Flanged Thick2", 0.4
    m_oSymbolHelper.AddInputDef 32, "EndHead2SphericalRadius", "End Head2 Spherical Radius", 2
    m_oSymbolHelper.AddInputDef 33, "EndHead3ConeTopDiameter", "End Head3 Cone Top Diameter", 1.7
    m_oSymbolHelper.AddInputDef 34, "EndHead3ConeHeight", "End Head3 Cone Height", 0.7
    m_oSymbolHelper.AddInputDef 35, "EndHead3KnuckleRadius", "End Head3 Knuckle Radius", 0.5
    m_oSymbolHelper.AddInputDef 36, "EndHead3Domeradius", "End Head3 Dome Radius", 1.6
    m_oSymbolHelper.AddInputDef 37, "EndHead3FlangedThick1", "End Head3 Flanged Thick1", 0.3
    m_oSymbolHelper.AddInputDef 38, "EndHead3FlangedThick2", "End Head3 Flanged Thick2", 0.3
    m_oSymbolHelper.AddInputDef 39, "EndHead3SphericalRadius", "End Head3 Spherical Radius", 1.6
    m_oSymbolHelper.AddInputDef 40, "InsulationThickness", "Insulation Thickness", 0.025
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 19
    
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Cylider1", "Cylinder 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Support1", "Support 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Cylider2", "Cylinder 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Support2", "Support 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Cylider3", "Cylinder 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Support3", "Support 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Cylider4", "Cylinder 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Skirt", "Skirt", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "SurfaceAtBottomOfSkirt", "Surface at Bottom of Skirt", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "CircularFoundationPort", "Circular Foundation Port", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "InsCylider1", "Insulated Cylinder 1", Insulation
    m_oSymbolHelper.AddOutputDef 13, "InsSupport1", "Insulated Support 1", Insulation
    m_oSymbolHelper.AddOutputDef 14, "InsCylider2", "Insulated Cylinder 2", Insulation
    m_oSymbolHelper.AddOutputDef 15, "InsSupport2", "Insulated Support 2", Insulation
    m_oSymbolHelper.AddOutputDef 16, "InsCylider3", "Insulated Cylinder 3", Insulation
    m_oSymbolHelper.AddOutputDef 17, "InsSupport3", "Insulated Support 3", Insulation
    m_oSymbolHelper.AddOutputDef 18, "InsCylider4", "Insulated Cylinder 4", Insulation
    m_oSymbolHelper.AddOutputDef 19, "ComplexVCylEqpSkControlPoint", "Control Point of Complex Vertical Cylindrical Equipment With Skirt", ReferenceGeometry
    
    'Aspects
    m_oSymbolHelper.NumAspects = 3
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "RefGeom", "ReferenceGeometry", ReferenceGeometry
    
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


