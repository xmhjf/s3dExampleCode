VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WTSupForSmallLiSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:36 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   WTSupForSmallLiSym.cls
'   ProgID:         SP3DWeldTySSLCompAsm.WTSupForSmallLiSym
'   Author:         svsmylav
'   Creation Date:  Tuesday, 22 April 2003
'   Description:
'   The symbol is prepared based on symbol of The DOW Chemical company. "Dow Welded Trunnion Type Pipe Support Detail For Small Lines Through 4 inch Pipe"
'   The symbol consists of 17 Physical outputs. Of Which 10 plates, 2 saddles, One Reference plane, 2 lines and 2 points.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "WTSupForSmallLiSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DWeldTySSLCompAsm"
    m_oSymbolHelper.ClassName = "WTSupForSmallLiSym"
    
' Inputs
    m_oSymbolHelper.NumInputs = 4
    m_oSymbolHelper.AddInputDef 1, "VesselDiameter", "Vessel Diameter", 1.219
    m_oSymbolHelper.AddInputDef 2, "PipeCLtoTowerCL", "Pipe CenterLine to Tower CenterLine-Y", 1.066
    m_oSymbolHelper.AddInputDef 3, "TowertoGuideBase", "Tower Outer shell to Guide Base-K", 0.254
    m_oSymbolHelper.AddInputDef 4, "PipeOD", "Pipe Outer Diameter", 0.101
        
' Outputs
    m_oSymbolHelper.NumOutputs = 17
    m_oSymbolHelper.AddOutputDef 1, "PipeSideSaddle", "Pipe Side Saddle guide", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Plate1", "Plate 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Plate2", "Plate 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Plate3", "Plate 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Plate4", "Plate 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Plate5", "Plate 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "VesselSideSaddle", "Vessel side Saddle guide", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Plate6", "Plate 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Plate7", "Plate 7", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Plate8", "Plate 8", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Plate9", "Plate 9", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "Plate10", "Plate 10", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "DefaultSurface", "Default Surface Reference plane on Tower", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 14, "VerticalPipeLine", "Vertical Line on Centreline of Pipe", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 15, "RadialTowerLine", "Radial line on Tower", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 16, "TowerPoint", "Point on Tower", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 17, "PipePoint", "Point on Pipe CentreLine", ReferenceGeometry
    
' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "ReferenceGeometry", "ReferenceGeometry", ReferenceGeometry
    
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


