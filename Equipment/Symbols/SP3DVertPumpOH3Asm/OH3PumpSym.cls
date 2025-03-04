VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OH3PumpSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   OH3PumpSym.cls
'   Author: Veena
'   Creation Date:  Wednesday, Jan 23 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------                 ---                         ------------------
'*****************************************************************************
Option Explicit
Private Const MODULE = "OH3PumpSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DVertPumpOH3Asm"
    m_oSymbolHelper.ClassName = "OH3PumpSym"
    

    'Inputs
    m_oSymbolHelper.NumInputs = 11

    m_oSymbolHelper.AddInputDef 1, "MotorHeight", "MotorHeight", 0.5
    m_oSymbolHelper.AddInputDef 2, "MotorDiameter", "MotorDiameter", 0.35
    m_oSymbolHelper.AddInputDef 3, "BaseplateBottomtoMotorFace", "BaseplateBottomtoMotorFace", 0.7
    m_oSymbolHelper.AddInputDef 4, "BaseplateBottomtoDischargeCL", "BaseplateBottomtoDischargeCL", 0.125
    m_oSymbolHelper.AddInputDef 5, "CLofPumptoSuctionFace", "CLofPumptoSuctionFace", 0.225
    m_oSymbolHelper.AddInputDef 6, "CLofPumptoDischargeFace", "CLofPumptoDischargeFace", 0.2
    m_oSymbolHelper.AddInputDef 7, "BaseplateHeight", "Height of Baseplate", 0.05
    m_oSymbolHelper.AddInputDef 8, "BaseplateWidth", "Width of Baseplate", 0.35
    m_oSymbolHelper.AddInputDef 9, "BaseplateLength", "Length of Baseplate", 0.35
    m_oSymbolHelper.AddInputDef 10, "XboltHole", "Bolt X Hole Location", 0.025
    m_oSymbolHelper.AddInputDef 11, "YboltHole", "Bolt Y Hole Location", 0.325

    'Outputs
    m_oSymbolHelper.NumOutputs = 47
    
    m_oSymbolHelper.AddOutputDef 1, "BodyCylinder1", "BodyCylinder1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "BodyCylinder2", "BodyCylinder2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "BodyCylinder3", "BodyCylinder3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "BodyDish1", "BodyDish1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "BodyDish2", "BodyDish2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Baseplate1", "Baseplate 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Baseplate2", "Baseplate 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Baseplate3", "Baseplate 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Baseplate4", "Baseplate 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Baseplate5", "Baseplate 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Baseplate6", "Baseplate 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "BodyCylinder5", "BodyCylinder5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "BodyCylinder6", "BodyCylinder6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "BodyCylinder7", "BodyCylinder7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "BodyCylinder8", "BodyCylinder8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "BodyCylinder9", "BodyCylinder9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "BodyCylinder10", "BodyCylinder10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "BodyCylinder11", "BodyCylinder11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "BodyCylinder12", "BodyCylinder12", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "Baseplate7", "Baseplate 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Baseplate8", "Baseplate 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Baseplate9", "Baseplate 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Baseplate10", "Baseplate 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "Baseplate11", "Baseplate 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "Baseplate12", "Baseplate 12", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "SuctionNozzle", "SuctionNozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "DischargeNozzle", "DischargeNozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "EqpFoundationPort", "Foundation Port under support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "PumpOrigin", "Control Point", SimplePhysical
    
    m_oSymbolHelper.AddOutputDef 30, "Point1", "Point 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 31, "Point2", "Point 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 32, "Point3", "Point 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 33, "Point4", "Point 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 34, "Point5", "Point 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 35, "Point6", "Point 6", SimplePhysical

    m_oSymbolHelper.AddOutputDef 36, "Edge1", "Edge 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 37, "Edge2", "Edge 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 38, "Edge3", "Edge 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 39, "Edge4", "Edge 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 40, "Edge5", "Edge 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 41, "Edge6", "Edge 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 42, "Edge7", "Edge 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 43, "Edge8", "Edge 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 44, "Edge9", "Edge 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 45, "Edge10", "Edge 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 46, "Edge11", "Edge 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 47, "Edge12", "Edge 12", SimplePhysical
    
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
    ' call symbol services default implementation of this method
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

           







