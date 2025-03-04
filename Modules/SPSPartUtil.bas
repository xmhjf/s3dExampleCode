Attribute VB_Name = "SPSPartUtil"
Option Explicit
 
'*******************************************************************
'  Copyright (C) 2007, Intergraph Corporation.  All rights reserved.
'
'  File:  SPSPartUtil.bas
'
'  Description: The file contains the utility functions used for getting attribute and Info Coll for part based objects.
'
'  History:
'   03/08/2000 - AS - Creation
'   09/02/2008 - GG - CR#120707 Added function GetDefAttribInfoColl
'   03/30/2009 - GG - DM#162556 Added function GetOutputCollFromSource
'******************************************************************
    
Private Const MODULE = "SPSLadderMacros.SPSPartUtil"
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Public Const strIJDOutputCollectionUUID = "{15916CAF-6CB5-11D1-A655-00A0C98D7F13}"
Public Const strOutputName = "toOutputs"

Public Sub GetOccAttribInfoColl(oSelObj As IJDObject, bIsOcc As Boolean, oAttrColl As IJDInfosCol)
    Const METHOD = "GetOccAttribInfoColl"
    On Error GoTo ErrorHandler
    
    Dim oPartOcc As IJPartOcc
    Dim oPart As IJDPart
    Dim oAttrs As IJDAttributes
    Dim varPartOccCLSID As Variant
    Dim varPartCLSID As Variant
    'CComBSTR                            strPartOccIID;
    'CComBSTR strPartIID;
    Dim oAttrMetadata As IJDAttributeMetaData
    Dim oResMgr As Object
    
    'get resource mgr
    Set oResMgr = oSelObj.ResourceManager
    
    Set oPartOcc = oSelObj
    If Not oPartOcc Is Nothing Then
        oPartOcc.GetPart oPart
    End If
    
    Set oAttrMetadata = oResMgr

    If Not oPart Is Nothing Then
        If bIsOcc Then
            varPartOccCLSID = oPart.GetPartOccCLSID
            Set oAttrColl = oAttrMetadata.ClassInterfaces(varPartOccCLSID, PublicFlag_ALL)
        Else
            varPartCLSID = oPart.GetPartCLSID
            Set oAttrColl = oAttrMetadata.ClassInterfaces(varPartCLSID, PublicFlag_ALL)
        End If
    End If

    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

Public Sub GetDefAttribInfoColl(oSelObj As IJDObject, oAttrColl As IJDInfosCol)
    Const METHOD = "GetDefAttribInfoColl"
    On Error GoTo ErrorHandler
    
    Dim oPart As IJDPart
    Dim varPartCLSID As Variant
    Dim oAttrMetadata As IJDAttributeMetaData
    
    'get resource mgr
    Set oAttrMetadata = oSelObj.ResourceManager
    
    Set oPart = oSelObj
    
    If Not oPart Is Nothing Then
        varPartCLSID = oPart.GetPartCLSID
        Set oAttrColl = oAttrMetadata.ClassInterfaces(varPartCLSID, PublicFlag_ALL)
    End If

    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

'Use this function to get output collection only if the collection is not available in memory
'The location information in the collection is local, the location information from outputCollection in SymbalDef is global
Public Function GetOutputCollFromSource(oSymbol As IJDSymbol, strRepName As String) As IJDTargetObjectCol
    Const METHOD = "GetOutputCollFromSource"
    On Error GoTo ErrorHandler
    Dim oOutColDisp As Object
    'Need a reference to Ingr SmartPlant 3D Proxy v 1.0 Library
    Dim oProxy As IJDProxy
    Dim oOutCol As IJDOutputCollection
    Dim oUnkCol As IUnknown
    Dim oRelation As IJDAssocRelation
    Set oOutColDisp = oSymbol.BindToOC(strRepName)
    Set oProxy = oOutColDisp
    Set oOutCol = oProxy.Source
    Set oRelation = oOutCol
    Set oUnkCol = oRelation.CollectionRelations(strIJDOutputCollectionUUID, strOutputName)
    Set GetOutputCollFromSource = oUnkCol
    Exit Function

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

