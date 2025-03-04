VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSymbolSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:        This module will perform a mapping from Detailing Symbols
'                   to Mfg Symbols.
'
'
'
'  History:
'       KKamph          May 12 2009     Created
'
'
'
'******************************************************************

Option Explicit

Implements IJMfgSymbolSelector



Private Const MODULE = "StrMfgSymbolSelector.CSymbolSelector"

Private Function IJMfgSymbolSelector_Select(ByVal pMfgObj As Object, ByVal pFeature As Object, ByVal pInputs As Object, ByVal strSymbol As String) As String
    Const METHOD = "IJMfgSymbolSelector_Select"
    On Error GoTo ErrorHandler
    
    Dim lPos As Long
    Dim sTempName As String
    lPos = Len(strSymbol) - InStr(1, StrReverse(strSymbol), "\")
    
    sTempName = Left(strSymbol, lPos + 1) & "Mfg" & Right(strSymbol, Len(strSymbol) - 1 - lPos)
    Dim oFSO As New FileSystemObject
    If oFSO.FileExists(oFSO.BuildPath(GetSymbolSharePath, sTempName)) Then
        IJMfgSymbolSelector_Select = sTempName
    Else
        IJMfgSymbolSelector_Select = strSymbol
    End If

CleanUp:
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetWorkCenter(oAssembly As Object) As String
Const METHOD = "GetWCFromParent"
On Error GoTo ErrorHandler
    Dim oTempPlnAssocCat        As IJPlnAssociateCatalog
    Dim oTempSRDWrkCenter       As IJSRDWorkcenter
    
    Set oTempPlnAssocCat = oAssembly
    
    If (Not oTempPlnAssocCat Is Nothing) Then
        Set oTempSRDWrkCenter = oTempPlnAssocCat.GetWorkCenter
        If (Not oTempSRDWrkCenter Is Nothing) Then
            'Getting WorkCenter Name
            GetWorkCenter = oTempSRDWrkCenter.WorkcenterName
        Else
            'In Case if that Assembly has no proper WorkCenter
            GetWorkCenter = "DefaultWorkCenter"
        End If
    Else
        'In Case the PinJig is created without Assembly
        GetWorkCenter = "DefaultWorkCenter"
    End If
    
CleanUp:
    Set oTempPlnAssocCat = Nothing
    Set oTempSRDWrkCenter = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Public Function GetParentAssembly(oMfgChild As IJMfgChild) As IJAssemblyBase
Const METHOD = "GetParentAssembly"
On Error GoTo ErrorHandler
    Dim oObject As Object
    
    Set oObject = oMfgChild.GetParent
    
    If TypeOf oObject Is IJAssemblyBase Then
        Set GetParentAssembly = oObject
    Else
        Dim oAssemblyChild As IJAssemblyChild
        Set oAssemblyChild = oObject
        Set GetParentAssembly = oAssemblyChild.Parent
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Private Function GetSymbolSharePath() As String
    Const METHOD = "GetSymbolLocationPath"
    On Error GoTo ErrorHandler
    
    Dim oContext As IJContext
    Dim strContextString As String
    Dim strSymbolShare As String
    
    strContextString = "OLE_SERVER"
    
    'Get IJContext
    Set oContext = GetJContext()
    
    If Not oContext Is Nothing Then
        strSymbolShare = oContext.GetVariable(strContextString)
    End If
    GetSymbolSharePath = strSymbolShare
    Set oContext = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function
