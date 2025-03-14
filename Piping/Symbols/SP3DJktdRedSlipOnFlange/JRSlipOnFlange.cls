VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JRSlipOnFlange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-05, Intergraph Corporation. All rights reserved.
'
'   JRSlipOnFlange.cls.cls
'   ProgID:         SP3DJktdRedSlipOnFlange.JRSlipOnFlange
'   Author:         ACM
'   Creation Date:  Wednesday, July 14 2004
'   Description:
'       This is a Reducing Slip-on Flange for Jacked-Piping.
'       On One side it has,One for Jacketed Pipe and One port of core pipe.
'       Other side it has one port.
'
'   Change History:
'   dd.mmm.yyyy     who               change description
'   -----------     ---               ------------------
'   21.Jul.2005     svsmylav  TR-81593: Corrected InsulationThickness value to 0.025.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Private Const MODULE = "JRSlipOnFlange:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DJktdRedSlipOnFlange"
    m_oSymbolHelper.ClassName = "JRSlipOnFlange"
    
' Inputs
    m_oSymbolHelper.NumInputs = 1
    m_oSymbolHelper.AddInputDef 1, "InsulationThickness", "InsulationThickness", 0.025
    
' Outputs
    m_oSymbolHelper.NumOutputs = 4
    m_oSymbolHelper.AddOutputDef 1, "BodyInsulation", "Insulation Body", Insulation
    m_oSymbolHelper.AddOutputDef 2, "Nozzle1", "Core Pipe Nozzle 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Nozzle2", "Core Pipe Nozzle 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Nozzle3", "Jacketed Pipe Nozzle", SimplePhysical
        
' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "Insulation", "Insulation", Insulation
    
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
    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolHelper.InstanciateDefinition(CodeBase, defParameters, ActiveConnection)
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
  'IJDUserSymbolServices_GetDefinitionName = m_ProgId
    IJDUserSymbolServices_GetDefinitionName = m_oSymbolHelper.ProjectName + "." + m_oSymbolHelper.ClassName
    
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    
  
    ' call symbol services default implementation of this method passing in the array of input types,
    ' AspectTypes and OutputTypes which are defined in the class initialize method of the symbol
    m_oSymbolHelper.InitializeSymbolDefinition pSymbolDefinition
      
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
        
    m_oSymbolHelper.InvokeRepresentation sblOcc, repName, outputcoll, arrayOfInputs
        
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False

End Function


