VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorRotatEqpSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:02 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   HorRotatEqpSym.cls
'   Author: RUK
'   Creation Date:  Friday, Feb 19 2007
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit

Private Const MODULE = "HorRotatEqpSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_405HoriRotEqpAsm"
    m_oSymbolHelper.ClassName = "HorRotatEqpSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 20

    m_oSymbolHelper.AddInputDef 1, "BaseLength", "Length of Base", 1
    m_oSymbolHelper.AddInputDef 2, "BaseWidth1", "Width1 of Base", 0.25
    m_oSymbolHelper.AddInputDef 3, "BaseWidth2", "Width2 of Base", 0.25
    m_oSymbolHelper.AddInputDef 4, "BaseThickness", "Thickness of Base", 0.1
    m_oSymbolHelper.AddInputDef 5, "DistBaseLeftToDischarge", "Dist between Base Left to Centreline of Discharge", 0.25
    m_oSymbolHelper.AddInputDef 6, "DistEqpLeftToDischarge", "Dist between Eqp Left to Centreline of Discharge", 0.25
    m_oSymbolHelper.AddInputDef 7, "DistEqpRightToDischarge", "Dist between Eqp Right to Centreline of Discharge", 0.4
    m_oSymbolHelper.AddInputDef 8, "EqpWidth1", "Width1 of Equipment", 0.25
    m_oSymbolHelper.AddInputDef 9, "EqpWidth2", "Width2 of Equipment", 0.25
    m_oSymbolHelper.AddInputDef 10, "EqpDepth1", "Depth1 of Equipment", 0.25
    m_oSymbolHelper.AddInputDef 11, "EqpDepth2", "Depth2 of Equipment", 0.25
    m_oSymbolHelper.AddInputDef 12, "RotEqpLength", "Length of Rotating Equipment", 0.3
    m_oSymbolHelper.AddInputDef 13, "RotEqpDiameter", "Diameter of Rotating Equipment", 0.1
    m_oSymbolHelper.AddInputDef 14, "DriverLength", "Length of Driver", 0.3
    m_oSymbolHelper.AddInputDef 15, "DriverWidth1", "Width1 of Driver", 0.25
    m_oSymbolHelper.AddInputDef 16, "DriverWidth2", "Width2 of Driver", 0.25
    m_oSymbolHelper.AddInputDef 17, "DriverHeight", "Height of Driver", 0.3
    m_oSymbolHelper.AddInputDef 18, "XboltHole", "Bolt X Hole Location", 0.025
    m_oSymbolHelper.AddInputDef 19, "YboltHole", "Bolt Y Hole Location", 0.475
    m_oSymbolHelper.AddInputDef 20, "InsulationThickness", "Insulation Thickness", 0.025

    'Outputs
    m_oSymbolHelper.NumOutputs = 30
    
    m_oSymbolHelper.AddOutputDef 1, "Equipment", "Equipment", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Driver", "Driver", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "RotEquip", "Rotating Equipment", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "SupportBodyTopPlane", "Support Body Top Plane", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "DefaultSurface", "Support Body Bottom Plane", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "SupportBodySidePlane1", "Support Body side plane1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "SupportBodySidePlane2", "Support Body side plane2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "SupportBodySidePlane3", "Support Body side plane3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "SupportBodySidePlane4", "Support Body side plane4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "CenterPos1", "FaceCenter Position1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "CenterPos2", "FaceCenter Position2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "CenterPos3", "FaceCenter Position3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "CenterPos4", "FaceCenter Position4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "CenterPos5", "FaceCenter Position5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "CenterPos6", "FaceCenter Position6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "Edge1", "Edge 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "Edge2", "Edge 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "Edge3", "Edge 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "Edge4", "Edge 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "Edge5", "Edge 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Edge6", "Edge 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Edge7", "Edge 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Edge8", "Edge 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "Edge9", "Edge 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "Edge10", "Edge 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "Edge11", "Edge 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "Edge12", "Edge 12", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "EqpFoundationPort", "EqpFoundation Port", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "InsulatedEqp", "Insulated Equipment", Insulation
    m_oSymbolHelper.AddOutputDef 30, "HorRotEqpADvrCompControlPoint", "Control Point of  Horizontal Rotating Equipment And Driver Component", ReferenceGeometry

    'Aspects
    m_oSymbolHelper.NumAspects = 3
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    m_oSymbolHelper.AddAspectDef 3, "RefGeometry", "ReferenceGeometry", ReferenceGeometry

    Exit Sub
ErrorHandler:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    Set m_oSymbolHelper = Nothing
End Sub

'note that the rest of the code is generated by the Symbol Wizard and should not need to be modified
Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    'call symbol services default implementation of this method
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
