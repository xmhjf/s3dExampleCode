Attribute VB_Name = "PreNestingutilFunctions"
'----------------------------------------------------------------------------
' Copyright (C) 2000-2001, Intergraph Corporation.  All rights reserved.
'
' Project
'   All Structural Manufacturing pre-nesting related methods
'
' File
'   PreNestingutilFunctions.bas
'
' Description
'   This file contains helper functions common for pre-nesting
'
' Author
'   Sivaprasad
'
' History:
'   2009-08-20  Sivaprasad     Creation date
'----------------------------------------------------------------------------

Option Explicit
Const MODULE = "PreNestingutilFunctions"
Public Const PRODUCTION_STAGE = 10

Private Function GetActiveConfigShipClass() As IJProjectRoot
    Const METHOD = "GetActiveConfigShipClass"
    On Error GoTo ErrHandler

    Dim oMfgEntHelper As IJMfgEntityHelper
    Set oMfgEntHelper = New MfgEntityHelper

    Dim oActiveProject As IJProjectRoot
    Set oActiveProject = oMfgEntHelper.GetConfigProjectRoot

    Set GetActiveConfigShipClass = oActiveProject

    Exit Function
    
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
                     
Public Function GetProjectRootNestingStage() As Long
    Const METHOD = "GetProjectRootNestingStage"
    On Error GoTo ErrHandler
    
    'Initialize with the value of production nesting
    GetProjectRootNestingStage = 10
    
    Dim oActiveProject As IJProjectRoot
    Set oActiveProject = GetActiveConfigShipClass
        
    Dim oAttribute     As IJDAttribute
    Set oAttribute = GetNestingStageAttribute(oActiveProject)
    
    If Not oAttribute Is Nothing Then
        If Not oAttribute.Value = -1 Then 'if the attribute value is not set or empty
            GetProjectRootNestingStage = oAttribute.Value
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

' Method to get the required attribute from the table
Public Function GetMfgObjectNestingStageArrtibute(oMfgBO As Object) As Long
    Const METHOD = "GetMfgObjectNestingStageArrtibute"
    On Error GoTo ErrHandler
    
    Dim oAttribute         As IJDAttribute
    Set oAttribute = GetNestingStageAttribute(oMfgBO)
    
    If Not oAttribute Is Nothing Then
        GetMfgObjectNestingStageArrtibute = oAttribute.Value
    Else
        GetMfgObjectNestingStageArrtibute = 0
    End If
    Exit Function
   
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

' Method to set the required attribute from the table
Public Function GetNestingStageAttribute(oBO As Object) As IJDAttribute
Const METHOD = "GetNestingStageAttribute"
On Error GoTo ErrorHandler
    
    Dim oIJDAttributes      As IJDAttributes
    Dim oMetaDataHelp       As IJDAttributeMetaData
    Dim oIJDInterfaceInfo   As IJDInterfaceInfo
    Dim vItfId              As Variant
    Dim oIJDAttributesCol   As IJDAttributesCol
    
    Set oMetaDataHelp = oBO
    Set oIJDAttributes = oBO
    
    Dim oIJDPOM             As IJDPOM
    Dim oIJDObject          As IJDObject
    
    Set oIJDObject = oBO
    Set oIJDPOM = oIJDObject.ResourceManager
    
    Dim oMoniker            As IUnknown
    Set oMoniker = oIJDPOM.GetObjectMoniker(oBO)
    
    'If (oIJDPOM.SupportsInterface(oMoniker, "IJMfgPreNestingService")) Then  '-- giving err when oBO is ProjectRoot
        For Each vItfId In oIJDAttributes
            If oIJDAttributes.IsInterfacePublic(vItfId) Then
                Set oIJDInterfaceInfo = oMetaDataHelp.InterfaceInfo(vItfId)
                If Not oIJDInterfaceInfo Is Nothing Then
                    If oIJDInterfaceInfo.Name = "IJMfgPreNestingService" Then
                        If Not oIJDInterfaceInfo.AttributeCollection Is Nothing Then
                            If oIJDInterfaceInfo.AttributeCollection.Count > 0 Then
                                Set oIJDAttributesCol = oIJDAttributes.CollectionOfAttributes(oIJDInterfaceInfo.Type)
                                Set GetNestingStageAttribute = oIJDAttributesCol.Item(1)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next vItfId
    'End If
       
CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    GoTo CleanUp
End Function
