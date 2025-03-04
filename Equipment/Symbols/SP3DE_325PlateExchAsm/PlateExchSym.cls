VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateExchSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:47 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   PlateExchSym.cls
'   Author: VRK
'   Creation Date:  Wednesday, Feb 21 2007
'
'   Description:
'   This is E325 Plate Exchanger Assembly.
'   This Symbol Uses two Aspects
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                            31
'      2                ReferenceGeometry                    1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'******************************************************************************

Option Explicit
Private Const MODULE = "PlateExchSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DE_325PlateExchAsm"
    m_oSymbolHelper.ClassName = "PlateExchSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 19

    m_oSymbolHelper.AddInputDef 1, "BotPlate1toPP", "Bottom Plate 1 to Placement Point", 1
    m_oSymbolHelper.AddInputDef 2, "Plate1Height", "Plate 1 Height", 1.75
    m_oSymbolHelper.AddInputDef 3, "Plate1Thickness", "Plate 1 Thickness", 0.2
    m_oSymbolHelper.AddInputDef 4, "PlatesWidth", "Plates Width", 1
    m_oSymbolHelper.AddInputDef 5, "ExchangerLength", "Exchanger Length", 0.75
    m_oSymbolHelper.AddInputDef 6, "ExchangerHeight", "Exchanger height", 1
    m_oSymbolHelper.AddInputDef 7, "ExchangerWidth", "Exchanger Width", 0.75
    m_oSymbolHelper.AddInputDef 8, "Plate2Height", "Plate 2 Height", 1.5
    m_oSymbolHelper.AddInputDef 9, "Plate2Thickness", "Plate 2 Thickness", 0.2
    m_oSymbolHelper.AddInputDef 10, "RodsLength", "Rods Length", 0.75
    m_oSymbolHelper.AddInputDef 11, "RodsDiameter", "Rods Diameter", 0.2
    m_oSymbolHelper.AddInputDef 12, "BotSidePlatetoRod2Cen", "Bottom Side Plate to Rod 2 Center", 0.5
    m_oSymbolHelper.AddInputDef 13, "RodsCentoCen", "Rods Center to Center", 1
    m_oSymbolHelper.AddInputDef 14, "SidePlateThickness", "Side Plate Thickness", 0.2
    m_oSymbolHelper.AddInputDef 15, "SidePlateWidth", "Side Plate Width", 0.4
    m_oSymbolHelper.AddInputDef 16, "FP1XboltHole", "FP1 Bolt X Hole Location", 0.5
    m_oSymbolHelper.AddInputDef 17, "FP1YboltHole", "FP1 Bolt Y Hole Location", 0.1
    m_oSymbolHelper.AddInputDef 18, "FP2XboltHole", "FP2 Bolt X Hole Location", 0.2
    m_oSymbolHelper.AddInputDef 19, "FP2YboltHole", "FP2 Bolt Y Hole Location", 0.1

    'Outputs
    m_oSymbolHelper.NumOutputs = 32
    
    m_oSymbolHelper.AddOutputDef 1, "Plane1", "Plane 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Plane2", "Plane 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Plane3", "Plane 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Plane4", "Plane 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Plane5", "Plane 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "ExchangerBody", "Exchanger Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Plate2", "Plate 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Rod1", "Rod 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Rod2", "Rod 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "SidePlate", "Side Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "CenterPos1", "FaceCenter Position1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "CenterPos2", "FaceCenter Position2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "CenterPos3", "FaceCenter Position3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "CenterPos4", "FaceCenter Position4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "CenterPos5", "FaceCenter Position5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "CenterPos6", "FaceCenter Position6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "Edge1", "Edge 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "Edge2", "Edge 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "Edge3", "Edge 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Edge4", "Edge 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Edge5", "Edge 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Edge6", "Edge 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "Edge7", "Edge 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "Edge8", "Edge 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "Edge9", "Edge 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "Edge10", "Edge 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "Edge11", "Edge 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "Edge12", "Edge 12", SimplePhysical
    m_oSymbolHelper.AddOutputDef 30, "EqpFoundationPort1", "EqpFoundation Port1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 31, "EqpFoundationPort2", "EqpFoundation Port2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 32, "PlateExchangerCompControlPoint", "Control Point of Plate Exchanger Component", ReferenceGeometry

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

