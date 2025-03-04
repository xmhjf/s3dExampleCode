VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHeadSectionSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   CHeadSectionSym.cls
'   ProgID: MHHeadSection.CHeadSectionSym
'   Author: Neelima BhupatiRaju
'   Creation Date: 08-Nov-2010
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   08-Nov-2010     Neelima B               Initial Creation
'   27-Sep-2012     Shireesha M             TR-CP-221650    2D HeasSectionSymbol is changed in Shape & position with changing HSLength
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CHeadSectionSym:"  'Used for error messages
Private m_oSymbolHelper         As IJSymbolHelper
   
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount              As Integer
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "MHHeadSection"
    m_oSymbolHelper.ClassName = "CHeadSectionSym"
    
    'Inputs
    m_oSymbolHelper.NumInputs = 8
    
    'Chaged the default values for TR-CP-221650
    m_oSymbolHelper.AddInputDef 1, "HSLength", "Chute Head Section Length", 3.5
    m_oSymbolHelper.AddInputDef 2, "HSWidth", "Chute Head Section Width", 4#
    m_oSymbolHelper.AddInputDef 3, "HSDepth", "Chute Head Section Depth", 3#
    m_oSymbolHelper.AddInputDef 4, "PulleyCenterPointLength", "Pulley Center Point Length", 0.5
    m_oSymbolHelper.AddInputDef 5, "PulleyCenterPointDepth", "Pulley Center Point Depth", 1.5
    m_oSymbolHelper.AddInputDef 6, "CutOutRadius", "Cut Out Radius", 0.15
    m_oSymbolHelper.AddInputDef 7, "HSOpeningLength", "Head Section Opening Length", 0.8
    m_oSymbolHelper.AddInputDef 8, "HSConvergingAngle1", "Head Section Converging Angle1", 45
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 3
    
    m_oSymbolHelper.AddOutputDef 1, "HeadSectionPart_", "Head Section Part", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "OriginPoint", "Contol Point at the Origin", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Point_", "Contol Point as Reference", SimplePhysical

    'Aspects
    m_oSymbolHelper.NumAspects = 1
    
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "SimplePhysical", SimplePhysical
    
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




