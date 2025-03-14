VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeVent2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   FreeVent2.cls.cls
'   ProgID:         SP3DFreeVent2.FreeVent2
'   Author:         svsmylav
'   Creation Date:  Wednesday 1, Feb 2006
'   Description:
'    This symbol constructed as per Model 5000 in the Growth corporation's Catalog GrowthC.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   10.12.2007      PS      CR-132021  Modified the default value of Vent Angle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "FreeVent2:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DFreeVent2"
    m_oSymbolHelper.ClassName = "FreeVent2"
    

' Inputs
    m_oSymbolHelper.NumInputs = 3
    m_oSymbolHelper.AddInputDef 1, "PipingSpecialtyWidth", "Piping Specialty Width", 0.26
    m_oSymbolHelper.AddInputDef 2, "PipingSpecialtyHeight", "Piping Specialty Height", 0.17
    m_oSymbolHelper.AddInputDef 3, "VentAngle", "Vent Angle", 4 * Atn(1)
        
' Outputs
    m_oSymbolHelper.NumOutputs = 2
    m_oSymbolHelper.AddOutputDef 1, "FreeeVentBody", "Freee Vent Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "PNoz1", "Nozzle 1", SimplePhysical
    
' Aspects
    m_oSymbolHelper.NumAspects = 1
    m_oSymbolHelper.AddAspectDef 1, "Physical", "PipingAspect Description", SimplePhysical
     
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

