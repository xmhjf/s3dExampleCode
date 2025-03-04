VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlaceFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------------------------
' Copyright (C) 2009, Intergraph Corporation. All rights reserved.
'
' Project:
'     AutomaticFeaturePlacement.vbp
'
' File:
'     PlaceFeatures.cls
'
' Description:
'     Used to place corner features on multiple plates within the Execute Detailing
'     command.
'
' Notes:
'     <notes>
'
' Author:
'     Maria White      05/04/2004     Initial creation.
'
'------------------------------------------------------------------------------------------------------------

Option Explicit

Implements IJAutomaticFeaturePlacement

Private Function IJAutomaticFeaturePlacement_PlaceBatchFeatures(ByVal ppDetailedParts As IEnumMoniker) As String

    Exit Function

    If ppDetailedParts Is Nothing Then Exit Function
    
    On Error GoTo ErrorHandler
    
    IJAutomaticFeaturePlacement_PlaceBatchFeatures = ""
    
    ' DIm variables for retrieving the objects of interest
    Dim oPOM As IJDPOM
    Dim oAccessMiddle As IJDAccessMiddle
    Dim oEnumMoniker As IEnumMoniker
    Dim oObjectOfInterest As Object

    Dim varMoniker As Variant
    Dim nTotalNumOfParts As Long
    Dim bCommitObj As Boolean
    
    ppDetailedParts.Reset
    For Each varMoniker In ppDetailedParts
        nTotalNumOfParts = nTotalNumOfParts + 1
    Next
    
    If nTotalNumOfParts = 0 Then Exit Function
    
    ppDetailedParts.Reset

    Dim nObjectIdx As Long
    For Each varMoniker In ppDetailedParts
        bCommitObj = False
        
        GetPOM oPOM, oAccessMiddle
        
        ' output the current item that is being processed in case the MTC dies abruptly
        nObjectIdx = nObjectIdx + 1

        On Error Resume Next
        Set oObjectOfInterest = oPOM.GetObject(varMoniker)
        
        If oObjectOfInterest Is Nothing Then
            oAccessMiddle.AbortTransaction
        Else
            bCommitObj = ExecuteFeatureRule(oObjectOfInterest, oAccessMiddle.GetResourceManagerFromType("Model"))
            
            If bCommitObj = True Then
                Set oObjectOfInterest = Nothing
            
                On Error Resume Next
                oAccessMiddle.CommitTransaction "PlaceBatchFeatures"
                On Error GoTo ErrorHandler
                If Err.Number > 0 Then
                    IJAutomaticFeaturePlacement_PlaceBatchFeatures = "Transaction aborted; error during commit"
                    oAccessMiddle.AbortTransaction
                End If
            End If
        End If
        
        Set oPOM = Nothing
        Set oAccessMiddle = Nothing
    Next varMoniker
    
    ' Clear variables
    Set oEnumMoniker = Nothing
    Set oObjectOfInterest = Nothing
    Set varMoniker = Nothing
    
    GetPOM oPOM, oAccessMiddle
    Set oPOM = Nothing
    
    
    Exit Function
    
ErrorHandler:

End Function

'**************************************************************************
' GetPOM
' Abstract: Returns a POM and the AccessMiddle service
' Input:    strDbType - string of either "MODEL" or "CATALOG"
' Output    oPOM - persistent obj manager
'           oAccessMiddle - accessmiddle service
'**************************************************************************
Private Sub GetPOM(oPOM As IJDPOM, oAccessMiddle As IJDAccessMiddle)
                  
Const METHODNAME = "GetPOM"

    On Error GoTo ErrHandler
    Set oAccessMiddle = GetJContext().GetService("ConnectMiddle")
    Set oPOM = oAccessMiddle.GetResourceManagerFromType("Model")
    
    Exit Sub
    
ErrHandler:
    
End Sub

Private Function IJAutomaticFeaturePlacement_PlaceFeatures(ByVal ppDetailedParts As IJElements) As String

    Dim oElementObj As Object
    Dim oToDoList As IJToDoList
    Dim returnErrorMessage As String
    Dim bCommitObj As Boolean
    Dim oRevMgr As IJRevision
    Dim oPOM As IJDPOM
    Dim oAccessMiddle As IJDAccessMiddle
    
    ' This automated Feature placement rule uses the same logic as the MultiCornerFeature rule to
    ' place corner features on plate parts that satisfy the same requirements as the MultiCornerFeature
    ' rule.  This rule is delivered as an example.
    
    ' To execute this rule, remove the Exit Function code line immediately after this comment
    
    Exit Function
    
    On Error GoTo ErrorHandler
    
    GetPOM oPOM, oAccessMiddle
    'Iterate through collection of detailed plate parts, adding corner features to them
    bCommitObj = False
    For Each oElementObj In ppDetailedParts
        bCommitObj = ExecuteFeatureRule(oElementObj, oPOM)
        
        If bCommitObj = True Then
            ' Compute/commit features to database
            If oRevMgr Is Nothing Then
                Set oRevMgr = New REVISIONLib.JRevision
            End If
            
            On Error Resume Next
            oRevMgr.Compute oToDoList
            
            On Error GoTo ErrorHandler
            If Err.Number <> 0 Then
                ' Error(s) occurred.  Write them to return string for placement in Execute Detailing
                ' log file
            Else
                ' Clear variables
                Set oElementObj = Nothing
                
                oRevMgr.ComputeAndCommit "PlaceAutomaticFeatures", oToDoList

                Dim recCount As Integer
                If oToDoList.Count <> 0 Then
                    ' Error(s) occurred.  Write them to return string for placement in Execute Detailing
                    ' log file
                    For recCount = 1 To oToDoList.Count
                        Dim pToDoRecord As IJToDoRecord
                        Set pToDoRecord = oToDoList.Item(recCount)
                        
                        If Not pToDoRecord Is Nothing Then
                           If pToDoRecord.ToDoType = TODO_ERROR Then
                               Dim itemCount As Long
                               Dim i As Long
                            
                               itemCount = pToDoRecord.ErrorCount
                            
                               For i = 1 To itemCount
                                   Dim pObjectInError As IUnknown
                                   Dim pObjectToUpdate As IUnknown
                                   Dim strErrorMessage As String
                                   Dim bShouldDelete As Boolean
                                
                                   pToDoRecord.ErrorItem i, pObjectInError, strErrorMessage, bShouldDelete, pObjectToUpdate
                                   If Len(returnErrorMessage) = 0 Then
                                      returnErrorMessage = strErrorMessage
                                   Else
                                      returnErrorMessage = returnErrorMessage & vbNewLine & strErrorMessage
                                   End If
                               Next i
                            End If
                        End If
                    Next recCount
                End If
                
            End If
            
            Set oToDoList = Nothing
        End If
    Next
        
    Set oPOM = Nothing
    Set oAccessMiddle = Nothing
    
    IJAutomaticFeaturePlacement_PlaceFeatures = returnErrorMessage
       
    Exit Function
    
ErrorHandler:

End Function


