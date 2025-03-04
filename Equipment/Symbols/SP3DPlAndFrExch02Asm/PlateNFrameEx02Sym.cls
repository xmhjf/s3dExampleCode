VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateNFrameEx02Sym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:25 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   PlateNFrameEx02Sym.cls
'   ProgID:         SP3DPlAndFrExch02Asm.PlateNFrameEx02Sym
'   Author:         svsmylav
'   Creation Date:  Wed, May 27, 2003
'
'   This is a Plate Exchanger(E-184).
'   In this Symbol there are total 26 outputs.22 outputs are physical out of which 4 are pipe Nozzles and one
'   Foundation Port.
'   Nozzles are fully Parametric.
'   There are 2 Reference Geometry objects which are Control-Point at origin and a reference plane at support bottom.
'   This is a Dow symbol Plate Exchanger.
'
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "PlateNFrameEx02Sym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices


Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DPlAndFrExch02Asm"
    m_oSymbolHelper.ClassName = "PlateNFrameEx02Sym"
    
' Inputs
    m_oSymbolHelper.NumInputs = 13
    m_oSymbolHelper.AddInputDef 1, "EquipmentXLength", "Equipment Length", 1.435
    m_oSymbolHelper.AddInputDef 2, "EquipmentYLength", "Equipment Width", 0.78
    m_oSymbolHelper.AddInputDef 3, "EquipmentZLength", "Equipment Height", 2.244
    m_oSymbolHelper.AddInputDef 4, "FrontPlateThickness", "Front Plate Thickness", 0.1
    m_oSymbolHelper.AddInputDef 5, "RearPlateThickness", "Rear Plate Thickness", 0.07
    m_oSymbolHelper.AddInputDef 6, "PlatePackLength", "Plate Pack Length", 0.177
    m_oSymbolHelper.AddInputDef 7, "CPtoFaceY", "Control Point to start of front plate along width", 0.214
    m_oSymbolHelper.AddInputDef 8, "CPtoFaceZ", "Control Point to start of front plate bottom along Height", 0.38
    m_oSymbolHelper.AddInputDef 9, "SupportLength", "Clamp Length", 0.09
    m_oSymbolHelper.AddInputDef 10, "SupportWidth", "ClampWidth", 0.16
    m_oSymbolHelper.AddInputDef 11, "SupportThickness", "Clamp Thickness", 0.0127
    m_oSymbolHelper.AddInputDef 12, "XboltHole", "Hole Location along Length", 0.045
    m_oSymbolHelper.AddInputDef 13, "YboltHole", "Hole Location along Width", 0.62
    
' Outputs
    m_oSymbolHelper.NumOutputs = 22
    m_oSymbolHelper.AddOutputDef 1, "FrontPlate", "Front Plate Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "PlatePack", "Plate Pack Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "RearPlate", "Rear Plate Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "InternalRod1", "Internal Rod1 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "InternalRod2", "Internal Rod2 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "InternalRod3", "Internal Rod3 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "InternalRod4", "Internal Rod4 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "InternalRod5", "Internal Rod1 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "InternalRod6", "Internal Rod2 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "InternalRod7", "Internal Rod3 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "InternalRod8", "Internal Rod4 Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "Support1", "Horizontal Support1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "Support2", "Horizontal Support2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "Support3", "Vertical Support for Horizontal Supports", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "SupportClamp1", "Support Clamp1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "SupportClamp2", "Support Clamp2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "SupportClamp3", "Support Clamp3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "SupportClamp4", "Support Clamp4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "SupportClamp5", "Support Clamp5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "ExchangerFoundationPort", "Exchanger Foundation Port", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "ControlPoint", "Control Point at origin", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 22, "DefaultSurface", "Default Surface  at Support", ReferenceGeometry
    
' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "PipingAspect Description", SimplePhysical
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




