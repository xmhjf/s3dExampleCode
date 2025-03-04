VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultportG7WayDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CMultportG7WayDia.cls
'   Prog Id :       SP3DMultportG7WayDia.CMultportG7WayDia
'   Author:         ACM/MS
'   Creation Date:  Monday, July 27, 2004
'   Description:
'   This is a multi port 7 way diverver valve symbol. This is prepared based on Saunder's catalog.
'   Source: Gemu  Multiport Valves M600-8-6A, Aseptic valve manifold machined from a single block.
'   M600 valve manifold designs, developed and produced according to customer requirements/specifications
'   The symbol has multiple operators. Each operator for each output port
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CMultportG7WayDia:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DMultportG7WayDia"
    m_oSymbolHelper.ClassName = "CMultportG7WayDia"
    
' Inputs
    m_oSymbolHelper.NumInputs = 11
    m_oSymbolHelper.AddInputDef 1, "ValveBodyWidth", "Valve Body Width", 0.3048
    m_oSymbolHelper.AddInputDef 2, "ValveBodyHeight", "Valve Body Height", 0.254
    m_oSymbolHelper.AddInputDef 3, "Face1toCenter", "Face 1 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 4, "Face2toCenter", "Face 2 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 5, "Face3toCenter", "Face 3 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 6, "Face4toCenter", "Face 4 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 7, "Face5toCenter", "Face 5 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 8, "Face6toCenter", "Face 6 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 9, "Face7toCenter", "Face 7 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 10, "Face8toCenter", "Face 8 to Center", 0.2286
    m_oSymbolHelper.AddInputDef 11, "InsulationThickness", "Insulation Thickness", 0.0254
    
' Outputs
    m_oSymbolHelper.NumOutputs = 28
    m_oSymbolHelper.AddOutputDef 1, "TopTaperBody", "Top Taper part of Valve Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "SteppedBody", "Stepped part of Valve Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "BottomTaperBody", "Bottom Taper part of Valve Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Nozzle1", "Nozzle 1 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Nozzle2", "Nozzle 2 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Nozzle3", "Nozzle 3 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "Nozzle4", "Nozzle 4 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "Nozzle5", "Nozzle 5 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "Nozzle6", "Nozzle 6 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 10, "Nozzle7", "Nozzle 7 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "Nozzle8", "Nozzle 8 with length", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "ValveOperator1", "Port 3 Valve Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "ValveOperator2", "Port 4 Valve Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "ValveOperator3", "Port 5 Valve Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "ValveOperator4", "Port 6 Valve Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "ValveOperator5", "Port 7 Valve Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "ValveOperator6", "Port 8 Valve Operator", SimplePhysical
    m_oSymbolHelper.AddOutputDef 18, "InsTopTaperBody", "Insulation for Top Taper Body", Insulation
    m_oSymbolHelper.AddOutputDef 19, "InsValveBodyStep", "Insulation fo Stepped part", Insulation
    m_oSymbolHelper.AddOutputDef 20, "InsBottomTaperBody", "Insulation Bottom Taper Body", Insulation
    m_oSymbolHelper.AddOutputDef 21, "InsNozzle1", "Ins Nozzle 1", Insulation
    m_oSymbolHelper.AddOutputDef 22, "InsNozzle2", "Ins Nozzle 2", Insulation
    m_oSymbolHelper.AddOutputDef 23, "InsNozzle3", "Ins Nozzle 3", Insulation
    m_oSymbolHelper.AddOutputDef 24, "InsNozzle4", "Ins Nozzle 4", Insulation
    m_oSymbolHelper.AddOutputDef 25, "InsNozzle5", "Ins Nozzle 5", Insulation
    m_oSymbolHelper.AddOutputDef 26, "InsNozzle6", "Ins Nozzle 6", Insulation
    m_oSymbolHelper.AddOutputDef 27, "InsNozzle7", "Ins Nozzle 7", Insulation
    m_oSymbolHelper.AddOutputDef 28, "InsNozzle8", "Ins Nozzle 8", Insulation

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
