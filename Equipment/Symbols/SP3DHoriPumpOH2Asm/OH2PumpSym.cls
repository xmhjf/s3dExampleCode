VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OH2PumpSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   OH2PumpSym.cls
'   Author: Veena
'   Creation Date:  Monday, Jan 08 2006
'
'   Description:
'   Description:
'   This is Horizontal Pump Of Type OH2 Assembly.
'   This Assembly Uses one Aspect
'   No.Of Aspects           Name of Aspect                  No.of Outputs
'       1                       Physical                            34
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   14.Jan-2008     VRK             TR-128504:Model Motor with Cylinder
'******************************************************************************

Option Explicit

Private Const MODULE = "OH2PumpSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHoriPumpOH2Asm"
    m_oSymbolHelper.ClassName = "OH2PumpSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 14

    m_oSymbolHelper.AddInputDef 1, "SuctiontoDischargeDistance", "Z-SuctiontoDischargeDistance", 0.11
    m_oSymbolHelper.AddInputDef 2, "BaseplateBottomtoDischargeFace", "H-BaseplateBottomtoDischargeFace", 0.525
    m_oSymbolHelper.AddInputDef 3, "DischargeCLtoSuctionFace", "e-DischargeCLtoSuctionFace", 0.05
    m_oSymbolHelper.AddInputDef 4, "PumpSuctionFacetoCoupling", "P-PumpSuctionFacetoCoupling", 0.6
    m_oSymbolHelper.AddInputDef 5, "CouplingtoMotorEnd", "K-CouplingtoMotorEnd", 0.45
    m_oSymbolHelper.AddInputDef 6, "BaseplateLength", "L-Length of Baseplate", 1.1
    m_oSymbolHelper.AddInputDef 7, "BaseplateHeight", "BU-Height of Baseplate", 0.05
    m_oSymbolHelper.AddInputDef 8, "BaseplateWidth", "B-Width of Baseplate", 0.35
    m_oSymbolHelper.AddInputDef 9, "OffsetfromSuctiontoCLofBase", "O-OffsetfromSuctiontoCLofBaseplate", 0.005
    m_oSymbolHelper.AddInputDef 10, "CouplingLength", "y-CouplingLength", 0.05
    m_oSymbolHelper.AddInputDef 11, "BaseplateBottomtoSuctionCL", "h1-BaseplateBottomtoSuctionCL", 0.35
    m_oSymbolHelper.AddInputDef 12, "MotorDia", "D-Diameter of Motor", 0.18
    m_oSymbolHelper.AddInputDef 13, "XboltHole", "Bolt X Hole Location", 0.025
    m_oSymbolHelper.AddInputDef 14, "YboltHole", "Bolt Y Hole Location", 0.325

    'Outputs
    m_oSymbolHelper.NumOutputs = 34
    
    m_oSymbolHelper.AddOutputDef 1, "BodyBaseplate1", "BodyBaseplate1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "BodyBaseplate2", "BodyBaseplate2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "BodyBaseplate3", "BodyBaseplate3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "BodyBaseplate4", "BodyBaseplate4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "BodyBaseplate5", "BodyBaseplate5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "BodyBaseplate6", "BodyBaseplate6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "DefaultSurface", "DefaultSurface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "BodyCylinder1", "BodyCylinder1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "BodyCylinder2", "BodyCylinder2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "BodyCylinder3", "BodyCylinder3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "BodyCylinder4", "BodyCylinder4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "BodyCylinder5", "BodyCylinder5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "BodyCylinder6", "BodyCylinder6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "BodyCylinder7", "BodyCylinder7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "BodyCylinder8", "BodyMotor", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "BodyCylinder9", "BodyCylinder9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "BodyCylinder10", "BodyCylinder10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "SupportBox", "SupportBox", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "Baseplate1", "Baseplate 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "Baseplate2", "Baseplate 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "Baseplate3", "Baseplate 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "Baseplate4", "Baseplate 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "Baseplate5", "Baseplate 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "Baseplate6", "Baseplate 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 25, "Baseplate7", "Baseplate 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 26, "Baseplate8", "Baseplate 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 27, "Baseplate9", "Baseplate 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 28, "Baseplate10", "Baseplate 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 29, "Baseplate11", "Baseplate 11", SimplePhysical
    m_oSymbolHelper.AddOutputDef 30, "Baseplate12", "Baseplate 12", SimplePhysical
    m_oSymbolHelper.AddOutputDef 31, "EqpFoundationPort", "Foundation Port under support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 32, "PumpOrigin", "Control Point", SimplePhysical
    m_oSymbolHelper.AddOutputDef 33, "Edges_", "Edges For Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 34, "Points_", "Points For Base Plate", SimplePhysical
    
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

           







