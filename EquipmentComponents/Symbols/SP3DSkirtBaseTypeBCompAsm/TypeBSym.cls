VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeBSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   File:           TypeBSym.cls
'   ProgID:         SP3DSkirtBaseTypeBCompAsm.TypeBSym
'   Author:         MA
'   Creation Date:  Friday, May 4 2007
'
'   Description:
'   This is Type B Skirt Base Plate Component Assembly.
'   This Assembly Uses One Aspect
'   No.Of Aspects       Name of Aspect               No.of Outputs
'      1                Physical                             5
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   19.09.2007      RUK     TR-127065  Problems with skirt with base equipment component symbol.
'                               -Added new attribute  parNoOfBolts and according to that supports are created
'                               -Replaced the attribute parVesselDiamter With parSkirtInsideDiameter
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "TypeBSym:"  'Used for error messages
Private m_oSymbolHelper As IJSymbolHelper
  
'Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DSkirtBaseTypeBCompAsm"
    m_oSymbolHelper.ClassName = "TypeBSym"

    'Inputs
    m_oSymbolHelper.NumInputs = 14
    
    m_oSymbolHelper.AddInputDef 1, "SkirtInsideDiameter", "Vessel Diameter", 0.3449
    m_oSymbolHelper.AddInputDef 2, "A", "A", 0.2088
    m_oSymbolHelper.AddInputDef 3, "B", "B", 0.2088
    m_oSymbolHelper.AddInputDef 4, "C", "C", 0.2088
    m_oSymbolHelper.AddInputDef 5, "D", "D", 0.2088
    m_oSymbolHelper.AddInputDef 6, "E", "E", 0.2088
    m_oSymbolHelper.AddInputDef 7, "F", "F", 0.3449
    m_oSymbolHelper.AddInputDef 8, "H", "H", 0.2088
    m_oSymbolHelper.AddInputDef 9, "T1", "T1", 0.2088
    m_oSymbolHelper.AddInputDef 10, "T2", "T2", 0.2088
    m_oSymbolHelper.AddInputDef 11, "SkirtHeight", "Skirt Height", 0.2088
    m_oSymbolHelper.AddInputDef 12, "SkirtThickness", "Skirt Thickness", 0.2088
    m_oSymbolHelper.AddInputDef 13, "BoltSize", "Bolt Size", 0.2088
    m_oSymbolHelper.AddInputDef 14, "NumberOfBolts", "Number of Bolts", 4
    
    'Outputs
    m_oSymbolHelper.NumOutputs = 5
    
    m_oSymbolHelper.AddOutputDef 1, "DefaultSurface", "Default Surface", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "CircularBase", "Circular Base", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "VesselSkirt", "Vessel Skirt", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Support_", "Support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "FoundationPort", "FoundationPort", SimplePhysical
    
    'Aspects
    m_oSymbolHelper.NumAspects = 1
    
    m_oSymbolHelper.AddAspectDef 1, "Physical", "Physical", SimplePhysical
    
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Sub Class_Terminate()
    Set m_oSymbolHelper = Nothing
End Sub

'Note that the rest of the code is generated by the Symbol Wizard and should not need to be modified
Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    'Call symbol services default implementation of this method
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
    
    'The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
    
    Exit Function
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function






