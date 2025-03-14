VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlatformTypeFSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   PlatformTypeFSym.cls
'   ProgID:        SP3DPlatformTypeFAsm.PlatformTypeFSym
'   Author:        MS
'   Creation Date: Tuesday, Oct 12 2004
'   Description:
' This symbol is prepared from the PDS Equipment guide Miscellaneous Platform (A003)
'   Change History:
'   dd.mmm.yyyy     who        change description
'   -----------     ---        ------------------
'
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   10.12.2007      PS         CR-132021  Modified the default value of PlatformAngle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "PlatformTypeFSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler

    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DPlatformTypeFAsm"
    m_oSymbolHelper.ClassName = "PlatformTypeFSym"

' Inputs
    m_oSymbolHelper.NumInputs = 8
    m_oSymbolHelper.AddInputDef 1, "EqpDimension1", "Equipment Dimension P1", 0.2032
    m_oSymbolHelper.AddInputDef 2, "EqpDimension2", "Equipment Dimension P2", 0.2032
    m_oSymbolHelper.AddInputDef 3, "EqpDimension3", "Equipment Dimension P3", 0.2032
    m_oSymbolHelper.AddInputDef 4, "EqpDimension4", "Equipment Dimension P4", 0.2032
    m_oSymbolHelper.AddInputDef 5, "EqpDimension5", "Equipment Dimension P5", 0.4032
    m_oSymbolHelper.AddInputDef 6, "EqpDimension6", "Equipment Dimension P6", 0.3032
    m_oSymbolHelper.AddInputDef 7, "PlatformAngle", "Platform Angle P9", (8 / 3) * Atn(1)
    m_oSymbolHelper.AddInputDef 8, "PlatformHeight", "Platform Height", 0.2032
    
' Outputs
    m_oSymbolHelper.NumOutputs = 7
    m_oSymbolHelper.AddOutputDef 1, "Line1", "Line 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "Line2", "Line 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "Line3", "Line 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Line4", "Line 4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Line5", "Line 5", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "Line6", "Line 6", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "ProjBody", "Body of Platform", SimplePhysical

' Aspects
    m_oSymbolHelper.NumAspects = 1
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical

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

