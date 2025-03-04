VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHexagonalBoltSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2011, Intergraph Corporation. All rights reserved.
'
'   CHexagonalBoltSym.cls
'   ProgID: MHHexagonalBolt.CHexagonalBoltSym
'   Author: Neelima BhupatiRaju
'   Creation Date: 23-Sep-2011
'   Description: This class file has the collection of inputs and outputs
'   of the symbol being created.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   23-Sep-2011     Neelima B               Initial Creation
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CHexagonalBoltSym:"  'Used for error messages
Private m_oSymbolHelper         As IJSymbolHelper
   
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount              As Integer
    Dim EquipmentHole As Long
    EquipmentHole = 524288 'Aspect No 19 for equipment holes
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "MHHexagonalBolt"
    m_oSymbolHelper.ClassName = "CHexagonalBoltSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 17
    
    m_oSymbolHelper.AddInputDef 1, "BoltHeadDiameter", "Bolt Head Diameter", 0.012
    m_oSymbolHelper.AddInputDef 2, "BoltHeadThickness", "Bolt Head Thickness", 0.006
    m_oSymbolHelper.AddInputDef 3, "BoltMaterial", "Bolt Material", "Alloy Steel"
    m_oSymbolHelper.AddInputDef 4, "BoltGrade", "Bolt Grade", "8.8"
    m_oSymbolHelper.AddInputDef 5, "GripLength", "Grip Length", 0.01
    m_oSymbolHelper.AddInputDef 6, "ThreadLength", "Thread Length", 0.03
    m_oSymbolHelper.AddInputDef 7, "ShankLength", "Shank Length", 0.04
    m_oSymbolHelper.AddInputDef 8, "BoltDiameter", "Bolt Diameter", 0.006
    m_oSymbolHelper.AddInputDef 9, "NutDiameter", "Nut Diameter", 0.006
    m_oSymbolHelper.AddInputDef 10, "NutThickness", "Nut Thickness", 0.0048
    m_oSymbolHelper.AddInputDef 11, "NutMaterial", "Nut Material", "Alloy Steel"
    m_oSymbolHelper.AddInputDef 12, "NutGrade", "Nut Grade", "4.4"
    m_oSymbolHelper.AddInputDef 13, "WasherDiameter", "Washer Diameter", 0.006
    m_oSymbolHelper.AddInputDef 14, "WasherThickness", "Washer Thickness", 0.002
    m_oSymbolHelper.AddInputDef 15, "NumberOfWasher", "Number Of Washer", 2
    m_oSymbolHelper.AddInputDef 16, "WasherMaterial", "Washer Material", "Alloy Steel"
    m_oSymbolHelper.AddInputDef 17, "WasherGrade", "Washer Grade", "4.4"

    'Outputs
    m_oSymbolHelper.NumOutputs = 6

    m_oSymbolHelper.AddOutputDef 1, "BoltHead", "Bolt Head", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "BoltShank", "Bolt Shank", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "BoltCenterLine", "Bolt Center Line", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "ConduitPort", "Conduit Port", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "WorkingSpace", "Space for working", Maintenance
    m_oSymbolHelper.AddOutputDef 6, "HoleShape", "Equipment Hole Shape", EquipmentHole

    'Aspects
    m_oSymbolHelper.NumAspects = 3

    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "SimplePhysical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Maintenance", "Maintenance", Maintenance
    m_oSymbolHelper.AddAspectDef 3, "EquipmentHole", "EquipmentHole", EquipmentHole
    
    Exit Sub
    
ErrorHandler:
  ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub Class_Terminate()
    
    Set m_oSymbolHelper = Nothing

End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    
    ' This method is in charge of the creation of the symbol definition object
    ' You can keep the current design unchanged
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    On Error GoTo ErrorHandler
    
    'Call symbol services default implementation of this method
    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolHelper.InstanciateDefinition(CodeBase, defParameters, ActiveConnection)
  
    Exit Function

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
  IJDUserSymbolServices_GetDefinitionName = m_oSymbolHelper.ProjectName + "." + m_oSymbolHelper.ClassName    'm_ProgId

End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    On Error GoTo ErrorHandler
    
    m_oSymbolHelper.InitializeSymbolDefinition pSymbolDefinition
          
  Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
    
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    m_oSymbolHelper.InvokeRepresentation sblOcc, repName, outputcoll, arrayOfInputs
    
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function


