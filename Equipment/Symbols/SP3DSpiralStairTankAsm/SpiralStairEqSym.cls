VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpiralStairEqSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   SpiralStairEqSym.cls
'   Author: Veena
'   Creation Date:  Wednesday, Dec 20 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit

Private Const MODULE = "SpiralStairEqSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DSpiralStairTankAsm"
    m_oSymbolHelper.ClassName = "SpiralStairEqSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 13
    
    m_oSymbolHelper.AddInputDef 1, "VesselDiameter", "Vessel Diameter", 2
    m_oSymbolHelper.AddInputDef 2, "VesselTantoTan", "Vessel Tangent to Tangent", 6
    m_oSymbolHelper.AddInputDef 3, "StartElevation", "Start Elevation", 0.5
    m_oSymbolHelper.AddInputDef 4, "EndElevation", "End Elevation", 5
    m_oSymbolHelper.AddInputDef 5, "StartAngle", "Start Angle", 45
    m_oSymbolHelper.AddInputDef 6, "AngleBetSupports", "Angle between Supports", 15
    m_oSymbolHelper.AddInputDef 7, "RampAngle", "Ramp Angle", 0
    m_oSymbolHelper.AddInputDef 8, "SupportVerticalDistance", "Support Vertical Distance", 1
    m_oSymbolHelper.AddInputDef 9, "SupportVerticalSpacing", "Support Vertical Spacing", 0.5
    m_oSymbolHelper.AddInputDef 10, "SpiralDirection", "Spiral Direction", 1
    m_oSymbolHelper.AddInputDef 11, "DisplaySpiralPoints", "Display Spiral Points", 1
    m_oSymbolHelper.AddInputDef 12, "XboltHole", "Bolt X Hole Location", 0.05
    m_oSymbolHelper.AddInputDef 13, "YboltHole", "Bolt Y Hole Location", 1
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 4
    m_oSymbolHelper.AddOutputDef 1, "ObjStorageTank", "ObjStorageTank", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Point1_", "Point1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Point2_", "Point2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "EqpFoundationPort", "Foundation Port under support", SimplePhysical

    'Aspects
    m_oSymbolHelper.NumAspects = 1
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical

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
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
    Exit Function
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext
End Function

           









