VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBifurSectSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   CBifurSectSym.cls
'   ProgID: MHBifurcatedSection.CBifurSectSym
'   Author: Neelima BhupatiRaju
'   Creation Date: 08-Nov-2010
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   08-Nov-2010     Neelima B               Initial Creation
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CBifurSectSym:"  'Used for error messages
Private m_oSymbolHelper         As IJSymbolHelper
   
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Dim iCount              As Integer
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "MHBifurcatedSection"
    m_oSymbolHelper.ClassName = "CBifurSectSym"
    
    'Inputs
    m_oSymbolHelper.NumInputs = 7
    
    m_oSymbolHelper.AddInputDef 1, "BSLength", "Bifurcated Section Length", 1.25
    m_oSymbolHelper.AddInputDef 2, "BSWidth", "Bifurcated Section Width", 1.75
    m_oSymbolHelper.AddInputDef 3, "BSDepth", "Bifurcated Section Depth", 2.5
    m_oSymbolHelper.AddInputDef 4, "BSDepth1", "Bifurcated Section Depth1", 0.405
    m_oSymbolHelper.AddInputDef 5, "BSAngle1", "Bifurcated Section Angle1", 60
    m_oSymbolHelper.AddInputDef 6, "BSAngle2", "Bifurcated Section Angle2", 70
    m_oSymbolHelper.AddInputDef 7, "BSOpeningLength", "Bifurcated Section Opening Length", 1.25
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 2
    
    m_oSymbolHelper.AddOutputDef 1, "BifurcatedSection", "Bifurcated Section", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Point_", "Point as Reference", SimplePhysical

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




