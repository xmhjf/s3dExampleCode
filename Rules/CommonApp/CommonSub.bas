Attribute VB_Name = "CommonSub"
'*******************************************************************
'  Copyright (C) 2003, Intergraph Corporation.  All rights reserved.
'
'  Project: CommonSub
'
'  Abstract:
'
'  Author: Jim Fleming
'
'  19-Oct-2003  Moved common methods GetPartName and GetPart to
'               this module
'
'  29-June-2009  GG TR#167244 Split via the Can Rule results in a DesignedMember and BUTube
'                   Get Part directly from IJSmartOccurrence to solve timing issue
'******************************************************************

Option Explicit

Private Const PARTOCCINTERFACE = "IJPartOcc"
Private Const SMARTOCCINTERFACE = "IJSmartOccurrence"

Private Const PARTROLE = "part"
Private Const SMARTITEMROLE = "toSI_ORIG"

Public Const E_FAIL = -2147467259

Public Function GetPartName(ByVal oEntity As Object) As String

    Const METHOD = "GetPartName"
    On Error GoTo label
    
    Dim oPart As IJDPart
    
    Set oPart = GetPart(oEntity)
           
    If Not oPart Is Nothing Then
        GetPartName = oPart.PartNumber
    End If

    Exit Function
label:
    ' log error to middle tier
    GetPartName = ""
    Err.Raise E_FAIL
End Function

Public Function GetPart(ByVal oEntity As Object) As IJDPart

    Const METHOD = "GetPart"
    On Error GoTo label

    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oCollection As IMSRelation.DCollectionHelper 'IJDrelatiopnshipCol
    Dim oPart As Object
    Dim iSO As IJSmartOccurrence

    If Not oEntity Is Nothing Then
        If TypeOf oEntity Is IJSmartOccurrence Then
            Set iSO = oEntity
            Set oPart = iSO.ItemObject
            If oPart Is Nothing Then
                Set oPart = iSO.RootSelectionObject
            End If
            Set iSO = Nothing
        Else
            If TypeOf oEntity Is IJPartOcc Then
                Set oRelationHelper = oEntity
                Set oCollection = oRelationHelper.CollectionRelations(PARTOCCINTERFACE, PARTROLE)
            End If
            If Not oCollection Is Nothing Then
                If (oCollection.Count <> 0) Then
                    Set oPart = oCollection.Item(1)
                End If
            End If
        End If
    End If
    
    If Not oPart Is Nothing Then
        If TypeOf oPart Is IJDPart Then
            Set GetPart = oPart
        End If
    End If

    Exit Function
label:
    ' log error to middle tier
    Set GetPart = Nothing
    Err.Raise E_FAIL
End Function


