VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HPBSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 7/26/2005-9:30:43 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   File:           HPBSym.cls
'   ProgID:         SP3DBlowerHPBAsm.HPBSym
'   Author:         CSA, Inc.
'   Creation Date:  Monday, April 23 2007
'
'   Description:
'   This is High Pressure Blower Model HPB Assembly.
'   This Assembly Uses One Aspect
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             1
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "HPBSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
Private Const E_FAIL = &H80004005
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DBlowerHPBAsm"
    m_oSymbolHelper.ClassName = "HPBSym"

    'Inputs Section
    m_oSymbolHelper.NumInputs = 13

    m_oSymbolHelper.AddInputDef 1, "BottomtoCenter", "A-Bottom to Center", 0.466725
    m_oSymbolHelper.AddInputDef 2, "CentertoEndofOutletFlange", "B-Center to End of Outlet Flange", 0.43815
    m_oSymbolHelper.AddInputDef 3, "CentertoOutletCenter", "D-Center to Outlet Center", 0.2683002
    m_oSymbolHelper.AddInputDef 4, "FacetoFace", "H-Face to Face", 23.75
    m_oSymbolHelper.AddInputDef 5, "CentertoEndofInlet", "J-Center to End of Inlet", 0.60325
    m_oSymbolHelper.AddInputDef 6, "SmallShaftLength", "K-Small Shaft Length", 0.0889
    m_oSymbolHelper.AddInputDef 7, "InsideInletDiameter", "L-Inside Inlet Diameter", 0.168275
    m_oSymbolHelper.AddInputDef 8, "HousingWidth", "M-Housing Width", 0.15875
    m_oSymbolHelper.AddInputDef 9, "DistanceBaseCentertoHole", "T-Distance Base Center to Hole", 0.22225
    m_oSymbolHelper.AddInputDef 10, "DistanceBaseCentertoEnd", "U-Distance Base Center to End", 0.244475
    m_oSymbolHelper.AddInputDef 11, "ShaftDiameter", "SD-Shaft Diameter", 0.0381
    m_oSymbolHelper.AddInputDef 12, "XBoltHole", "HX-Hole Bolt X Location", 0.244475
    m_oSymbolHelper.AddInputDef 13, "YBoltHole", "HY-Hole Bolt Y Location", 0.1666875

    'Outputs Section
    m_oSymbolHelper.NumOutputs = 8

    m_oSymbolHelper.AddOutputDef 1, "BodyCylinder", "Body Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "OutletBox", "Outlet Box", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "InletCylinder", "Inlet Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ShaftCylinder", "Shaft Cylinder", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "NozzleCone", "Nozzle Cone", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "BottomBase", "Bottom Base", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "BasePlate", "Base Plate", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "EqpFoundationPort", "Equipment Foundation Port", SimplePhysical

    'Aspects Section
    m_oSymbolHelper.NumAspects = 1
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    
    Set m_oSymbolHelper = Nothing

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'NOTE: The code below this line was generated by the Symbol Wizard and should not be modified
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    
' This method is in charge of the creation of the symbol definition object
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
