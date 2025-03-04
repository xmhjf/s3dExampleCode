VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InducedDrSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   InducedDrSym.cls
'   Author: VRK
'   Creation Date:  Wednesday, Feb 14 2007
'
'   Description:
'   This is E332 Induced Draft AirCooler Assembly.
'   This Assembly Uses three Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             28
'      2                ReferenceGeometry                     1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit

Private Const MODULE = "InducedDrSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_332AirCoolerAsm"
    m_oSymbolHelper.ClassName = "InducedDrSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 14
    
    m_oSymbolHelper.AddInputDef 1, "CoolerWidth", "Cooler Width", 0.75
    m_oSymbolHelper.AddInputDef 2, "CoolerLength", "Cooler Length", 1.5
    m_oSymbolHelper.AddInputDef 3, "CoolerHeight", "Cooler Height", 0.4
    m_oSymbolHelper.AddInputDef 4, "PlenumLength", "Plenum Length", 1.2
    m_oSymbolHelper.AddInputDef 5, "PlenumHeight", "Plenum Height", 0.3
    m_oSymbolHelper.AddInputDef 6, "InletBlockLength", "Inlet Block Length", 0.2
    m_oSymbolHelper.AddInputDef 7, "BlockToptoCoolTop", "Block Top to Cooler Top", 0.1
    m_oSymbolHelper.AddInputDef 8, "InletBlockHeight", "Inlet Block Height", 0.2
    m_oSymbolHelper.AddInputDef 9, "CoolLeftfromPP", "Cooler Left from Placement Point", 0.1
    m_oSymbolHelper.AddInputDef 10, "NumberOfFans", "Number Of Fans", 2
    m_oSymbolHelper.AddInputDef 11, "Fan1CentoPP", "Center Fan 1 to Placement Point", 0.475
    m_oSymbolHelper.AddInputDef 12, "FansCentoCen", "Fans Center to Center", 0.75
    m_oSymbolHelper.AddInputDef 13, "FanDiameter", "Fan Diameter", 0.3
    m_oSymbolHelper.AddInputDef 14, "FanHeight", "Fan Height", 0.1

    'Outputs
    m_oSymbolHelper.NumOutputs = 29
    
    m_oSymbolHelper.AddOutputDef 1, "CoolerBodyPlane1", "Plane1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "CoolerBodyPlane2", "Plane2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "CoolerBodyPlane3", "Plane3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "CoolerBodyPlane4", "Plane4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "CoolerBodyPlane5", "Plane5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "CoolerLeftBox", "Cooler Left Box", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "CoolerRightBox", "Cooler Right Box", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "PlenumBox", "Plenum Box", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Fans_", "Fan", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "CenterPos1", "FaceCenter Position1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "CenterPos2", "FaceCenter Position2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "CenterPos3", "FaceCenter Position3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "CenterPos4", "FaceCenter Position4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "CenterPos5", "FaceCenter Position5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "CenterPos6", "FaceCenter Position6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "Edge1", "Edge 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "Edge2", "Edge 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "Edge3", "Edge 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "Edge4", "Edge 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Edge5", "Edge 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Edge6", "Edge 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Edge7", "Edge 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "Edge8", "Edge 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "Edge9", "Edge 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "Edge10", "Edge 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "Edge11", "Edge 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "Edge12", "Edge 12", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "InducedDraftAirCoolerCompControlPoint", "InducedDraftAirCoolerCompControlPoint", ReferenceGeometry

    'Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "RefGeometry", "ReferenceGeometry", ReferenceGeometry

    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
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

           











