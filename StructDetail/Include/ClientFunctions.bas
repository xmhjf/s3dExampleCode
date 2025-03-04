Attribute VB_Name = "ClientFunctions"
'-----------------------------------------------------------------------------------------
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'
'  Abstract:
'      Defines various client-tier functions shared among various projects
'
'  Notes:
'      Originally written to ease the transition to SP3Dv3 for error handling
'
'      Projects including this module must reference:
'           M:\CommonApp\Client\Bin\CmnAppErrHndlr.dll
'           M:\CommonApp\Client\Bin\CAppInterfaces.dll
'           X:\Client\Bin\Trader.dll  (Ingr SmartPlant 3D Trader v 1.0 Libary)
'           X:\Client\Bin\IMSICDPInterfacesTlb.dll  (Ingr SmartPlant 3D ICDPInterfaces v 1.0 Library)
'
'  History:
'      Chris Gibble         05/13/2002    Creation.
'      Steve Talbert       04/01/2004    Added method SetHelpContext
'      Veena               05/12/2006     DI-97144:Added  ResIds
'-----------------------------------------------------------------------------------------

Option Explicit
Option Private Module    ' Makes public functions private outside of project

Public Const IDH_GS_CLEAR_HELP_CONTEXT As Long = -1
Public Const IDH_GS_RESET_TO_PREVIOUS_HELP_CONTEXT As Long = -999

'Constant for common ship export services
Public Const TKCOMMONSHIPLOCALIZER = "smLocalizer"
Public m_oToDoList As IJToDoList
Private m_oListOfMonikers As Collection


Private m_oErrHandler As IJCommonError
Private m_lPreviousHelpContext As Long
Private m_oLocalizer As IJLocalizer  ' Module-level variable
Public Function CountObjectsOnToDoList(oColl As Collection, oToDoList As IJToDoList) As Long
    
    Dim oObj As Object
    Dim lCount As Long
    
    
    On Error Resume Next
    
    lCount = 0
    
    If Not oColl Is Nothing And Not oToDoList Is Nothing Then
        
        
        For Each oObj In oColl
            If CheckTransactionForFailure(oObj, oToDoList) Then
                lCount = lCount + 1
            End If
        Next
        
        Set oObj = Nothing
        
    End If
    
    CountObjectsOnToDoList = lCount
End Function
Public Function GetObjectCollFromMonikerColl() As Collection

    Dim oColl As Collection
    Dim monikerCount As Long
    Dim oObjName As IMoniker
    Dim oTrader As New Trader
    Dim oActiveConnection As IJDConnection
    Dim oWorkingSet As IJDWorkingSet
    
    Set oWorkingSet = oTrader.Service(TKWorkingSet, "")
    Set oActiveConnection = oWorkingSet.ActiveConnection
    Set oWorkingSet = Nothing
    Set oTrader = Nothing
    
    Set GetObjectCollFromMonikerColl = Nothing
    
    If m_oListOfMonikers Is Nothing Then
        Exit Function
    End If
    
    Set oColl = New Collection

    For monikerCount = 1 To m_oListOfMonikers.count
        oColl.Add oActiveConnection.GetObject(m_oListOfMonikers.Item(monikerCount))
    Next

    Set GetObjectCollFromMonikerColl = oColl
    Set oColl = Nothing


End Function

Public Sub GetMonikerCollFromObjectColl(oObjColl As IJDObjectCollection)

    
    Dim oObj As Object
    Dim oTrader As New Trader
    Dim oActiveConnection As IJDConnection
    Dim oWorkingSet As IJDWorkingSet
    
    Set oWorkingSet = oTrader.Service(TKWorkingSet, "")
    Set oActiveConnection = oWorkingSet.ActiveConnection
    Set oWorkingSet = Nothing
    Set oTrader = Nothing

    
    
    If oObjColl Is Nothing Then
        Exit Sub
    End If
    
    
    Set m_oListOfMonikers = Nothing
    Set m_oListOfMonikers = New Collection

    For Each oObj In oObjColl
        m_oListOfMonikers.Add oActiveConnection.GetObjectName(oObj)
    Next

    
    Set oActiveConnection = Nothing

End Sub






Public Function CheckTransactionForFailure(oObj As Object, oToDoList As IJToDoList) As Boolean

Dim oToDoRecord As IJToDoRecord
Dim count As Long, errCount As Long
Dim oObjInErr As Object, oObjToUpdate As Object
Dim strErrorMsg As String
Dim bShouldDelete As Boolean

CheckTransactionForFailure = False

If oObj Is Nothing Or oToDoList Is Nothing Then
    Exit Function
End If

If oToDoList.count <> 0 Then
    
    For count = 1 To oToDoList.count
    
        Set oToDoRecord = Nothing
        Set oToDoRecord = oToDoList.Item(count)
        
        For errCount = 1 To oToDoRecord.ErrorCount
        
            Set oObjInErr = Nothing
            Set oObjToUpdate = Nothing
            oToDoRecord.ErrorItem errCount, oObjInErr, strErrorMsg, _
            bShouldDelete, oObjToUpdate
            
            If oObjInErr Is oObj And oToDoRecord.ToDoType = TODO_ERROR Then
                CheckTransactionForFailure = True
                Exit Function
            End If
            Set oObjInErr = Nothing
            Set oObjToUpdate = Nothing
            
        Next
        
    Next
    
    
End If


Set oToDoRecord = Nothing
Set oObjInErr = Nothing
Set oObjToUpdate = Nothing
    
    

End Function

'********************************************************************

' Routine: ReportUnanticipatedError
'
' Description: Forwards call to common app's appropriate error handler
'********************************************************************
Public Sub ReportUnanticipatedError(strModule As String, strMethod As String)
    Dim ErrorHandler As New GSCADUtilities.Utilities
    If Err.Number <> 0 Then
        ErrorHandler.ReportUnanticipatedError strModule, strMethod
       
    
    End If
    
    Set ErrorHandler = Nothing
    
End Sub

'********************************************************************
' Routine: ReportAndRaiseUnanticipatedError
'
' Description: Forwards call to common app's appropriate error handler
'********************************************************************
Public Sub ReportAndRaiseUnanticipatedError(strModule As String, strMethod As String)
    
    Dim lErrNum As Long
    Dim strErrDesc As String
    
    lErrNum = Err.Number
    strErrDesc = Err.Description
    
    If m_oErrHandler Is Nothing Then
        Set m_oErrHandler = New CommonError
    End If

    m_oErrHandler.ReportAndRaiseCriticalError strModule & strMethod, lErrNum, strErrDesc
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure (Sub):
'     SetHelpContext
'
' Description:
'     Uses the IJHelp interface on the OnlineHelp object to set the specified help topic using its
'     context ID.  Typically, the task help file should have already been specified by the task
'     environment manager.
'
' Arguments:
'     lHelpContextID    Long    The help context ID.  This would generally be one of the context
'                                          values listed in the task's ...Help409.bas module OR the value to
'                                          clear the context OR the value to return to the previous context.
'------------------------------------------------------------------------------------------------------------
Public Sub SetHelpContext(lHelpContextID As Long)
    Const METHOD = "SetHelpContext"

    On Error GoTo ErrorHandler
    
    Dim oTrader As Trader
    Dim oHelp As IJHelp
    
    Set oTrader = New Trader
    If Not oTrader Is Nothing Then
        Set oHelp = oTrader.Service(TKHelp, "")
        If Not oHelp Is Nothing Then
        
            If lHelpContextID = IDH_GS_CLEAR_HELP_CONTEXT Then
                oHelp.HelpContext = IDH_GS_CLEAR_HELP_CONTEXT
                
            ElseIf lHelpContextID = IDH_GS_RESET_TO_PREVIOUS_HELP_CONTEXT Then
                oHelp.HelpContext = m_lPreviousHelpContext
                
            Else  'Store current help context, set new help context as specified
                m_lPreviousHelpContext = oHelp.HelpContext
                oHelp.HelpContext = lHelpContextID
            End If
            
        End If
    End If
    
    Set oTrader = Nothing
    Set oHelp = Nothing
    
    Exit Sub
    
ErrorHandler:
    ReportUnanticipatedError "ClientFunctions", METHOD
End Sub


Public Function GetLocalizer(Optional sResourcePath As String = "") As IJLocalizer
    If m_oLocalizer Is Nothing Then
        Set m_oLocalizer = CreateObject("IMSLocalizer.Localizer")
        If sResourcePath = "" Then
            m_oLocalizer.Initialize App.Path & "\" & App.EXEName & ".dll"
        Else
            m_oLocalizer.Initialize App.Path & sResourcePath & App.EXEName & ".dll"
        End If
    End If

    Set GetLocalizer = m_oLocalizer
End Function

Public Sub ShowToDoMessage(strModule As String, strMethod As String, Optional StrErrDsrp As String)
    Dim ErrorHandler As New GSCADUtilities.Utilities
    
    Dim stringLength As Integer
    Dim TempStr As String
    ' Explicitly set the string to Nothing.
    stringLength = Len(StrErrDsrp)

    ' E_FAIL Check for TODO List
    If CountObjectsOnToDoList(GetObjectCollFromMonikerColl(), m_oToDoList) <> 0 Then
        If stringLength <> 0 Then
            TempStr = "( " + StrErrDsrp + " ) "
        End If
        
        Dim oTrader As IJDTrader
        Set oTrader = New Trader
        
        Dim oLocalizerService As Object
        Set oLocalizerService = oTrader.Service(TKCOMMONSHIPLOCALIZER, "")
        
        MsgBox oLocalizerService.GetToDoListMessage1 + TempStr + _
               oLocalizerService.GetToDoListMessage2, _
                vbOKOnly + vbExclamation, strModule + " " + strMethod
                
        Set oLocalizerService = Nothing
    End If
    
    Set m_oToDoList = Nothing
    Set m_oListOfMonikers = Nothing
    Set ErrorHandler = Nothing
    
End Sub
    
Public Function GetBOForPG(bIsFirstClassBOInCreate As Boolean, Optional oDefaultBOForPG As Object = Nothing) As Object
    Const METHOD = "GetBOForPG"
    On Error GoTo ErrorHandler

    'Get appropriate BO for setting the permission group for internal objects (Particularly Symbol here)
    If bIsFirstClassBOInCreate Then
        'During the creation of the first class BO,
        'All internal objects go into appropriate default permission group for the first class BO
        Set GetBOForPG = oDefaultBOForPG
    Else
        'If the firstclass BO is being modified (note:Symbol may be new),
        'the internal objects go into child plate's PG
        Dim oTrader As New Trader
        Dim oSelectSet As IJSelectSet
        
        Set oSelectSet = oTrader.Service(TKSelectSet, "")
        Set GetBOForPG = oSelectSet.Elements.Item(1)
        
        Set oSelectSet = Nothing
        Set oTrader = Nothing
    End If
    
    Exit Function
ErrorHandler:
    ReportUnanticipatedError "ClientFunctions", METHOD
End Function
