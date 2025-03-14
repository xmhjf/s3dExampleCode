VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDefVals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   User Defined Values Process Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Const MODULE = "UserDefinedValues: "

Implements IJDMfgTemplateParameters

Private Const sSOURCEFILE           As String = "UserDefVals.cls"
Private m_oTemplateSetObject        As IJDMfgTemplateSet

Private Sub Class_Initialize()

    Set m_oTemplateSetObject = Nothing
    
End Sub

Private Sub Class_Terminate()

    Set m_oTemplateSetObject = Nothing

End Sub

Private Function IJDMfgTemplateParameters_GetAtrributeDefaultValues(ByVal bstrAttrNameColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    Const METHOD As String = "UserDefinedValues: IJDMfgTemplateParameters_GetDoubleValue"
    On Error GoTo ErrorHandler
    
    Dim oOutputValColl As Collection
    Set oOutputValColl = New Collection
    
    Dim iindex  As Long
    For iindex = 1 To bstrAttrNameColl.Count
        Select Case bstrAttrNameColl.Item(iindex)
            Case "MaterialType"
                oOutputValColl.Add "Steel - Carbon", "MaterialType"      ' make sure the value exist in Molded form Material thickness range table
            Case "MaterialGrade"
                oOutputValColl.Add "A", "MaterialGrade"                ' make sure the value exist in Molded form Material thickness range table
            Case "MaterialThickness"
                oOutputValColl.Add 0.0047625, "MaterialThickness"   ' make sure the value exist in Molded form Material thickness range table
            Case Else
                ' Do nothing
        End Select
    Next
    
    Set IJDMfgTemplateParameters_GetAtrributeDefaultValues = oOutputValColl
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3025, , "RULES")
End Function

Private Function IJDMfgTemplateParameters_GetDoubleValue(ByVal bstrParameterName As String) As Double
    Const METHOD As String = "UserDefinedValues: IJDMfgTemplateParameters_GetDoubleValue"
    On Error GoTo ErrorHandler
    
    Select Case bstrParameterName
    Case "MinimumCurveLength"
        ' All marks below the specified value will be skipped in output xml
        IJDMfgTemplateParameters_GetDoubleValue = 0.01  ' In meters

'    Case Else
'        IJDMfgTemplateParameters_GetDoubleValue = -1
        
    End Select

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1055, , "RULES")
End Function

Private Function IJDMfgTemplateParameters_GetLongValue(ByVal bstrParameterName As String) As Long
    Const METHOD As String = "IJDMfgTemplateParameters_GetLongValue"
    On Error GoTo ErrorHandler
    
    Dim oProcessSettings   As IJMfgTemplateProcessSettings
    Set oProcessSettings = m_oTemplateSetObject.GetProcessSettings
    
    Select Case bstrParameterName
        
        Case "ApplyMarginInProcess"

            ' Applying margin to an edges
            ' 0 - Don't apply margin to Template
            ' 1 - Do Apply Margin to Template
    
            IJDMfgTemplateParameters_GetLongValue = 1
            
            
            ' Do not apply margin for all box cases
            If oProcessSettings.TemplateType = "Box" Or _
                oProcessSettings.TemplateType = "UserDefined Box" Or _
                oProcessSettings.TemplateType = "UserDefined Box With Edges" Then
                
                IJDMfgTemplateParameters_GetLongValue = 0
            End If
        Case "TrimMark"
                IJDMfgTemplateParameters_GetLongValue = 1
        Case "UserDefinedNamingRule"
    
            IJDMfgTemplateParameters_GetLongValue = 0
            
            ' Allow editing of template names
            If oProcessSettings.TemplateType = "UserDefined" Or _
                oProcessSettings.TemplateType = "UserDefined Box" Or _
                oProcessSettings.TemplateType = "UserDefined Box With Edges" Then
                
                IJDMfgTemplateParameters_GetLongValue = 1
            End If
        Case Else
           IJDMfgTemplateParameters_GetLongValue = -1
    End Select

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1056, , "RULES")
End Function

Private Function IJDMfgTemplateParameters_GetStringValue(ByVal bstrParameterName As String) As String
    Const METHOD As String = "IJDMfgTemplateParameters_GetStringValue"
    On Error GoTo ErrorHandler
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1096, , "RULES")
End Function

Private Property Let IJDMfgTemplateParameters_Object(ByVal oTemplateSet As Object)
    Const METHOD As String = "IJDMfgTemplateParameters_Object"
    On Error GoTo ErrorHandler
    
    If Not m_oTemplateSetObject Is Nothing Then
        Set m_oTemplateSetObject = Nothing
    End If
    
    Set m_oTemplateSetObject = oTemplateSet

    Exit Property
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property


