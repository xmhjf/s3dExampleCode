VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckChamfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2005 Intergraph Corporation.  All Rights Reserved.
'
' File: CheckChamfers.cls
'
' Author: D.A. Trent
'
' Abstract:
'   Check Manufacturability rule that compares the Tee Weld Chamfer
'   Thickness with the Thickness of Plate on the Other Side of the Connection
'
'   This is for the Tee Weld Chamfer created by the Assembly Connection
'   for Plate Edge by Plate Face case and the Plate Edge is a result of a Split
'
'********************************************************************

Option Explicit

Private Const Module = "CheckMfctyTeeWeldChamfer.CheckChamfers"
Private Const strProgID = "CheckMfctyTeeWeldChamfer.CheckChamfers"

' The following Constants are based on the "Delivered" set of Rules
' The values will required modification if:
'   The Chamfer Question/Answer is changed
'   The Tee Weld Chamfer Root Selector name is changed
Private Const m_sRootTeeChamferSel = "RootTeeChamferSel"
Private Const m_sQuestion_ChamferThickness = "ChamferThickness"
Private Const m_dDeltaThickness = 0.001

Implements IJCheckMfcty

Public Enum ESeverityIndex
    siError = 101
    siWarning = 102
End Enum


Private Sub Class_Initialize()
    Const Method As String = "Class_Initialize"
    On Error GoTo ErrorHandler
    m_bLogError = False
    GoTo Cleanup

ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number

Cleanup:

End Sub

''////////////////////////////////////////////////////////////////////
'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description:
' This method is called by the Check Manufacturability mechanism
' to perform the desired checks for this rule.
'
' Inputs:
'     oElements     :   The selected items in the working set.
'     bStopOnError  :   Operator option to stop after first error
'                       has been encountered.
'     pOptionCodes  :   not used in this method.
'     pCallingObj   :   The object that made the call to this method.
'                       This is used to pass back information to the Check
'                       Manufacturability command.
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck( _
        ByVal oElements As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        pOptionCodes() As Variant, _
        ByVal pCallingObj As Object)
Const Method = "IJCheckMfcty_Execute"
    On Error GoTo ErrorHandler

    Dim nCount As Long
    Dim nMaxCount As Long
    
    Dim lngDummy As Long
    Dim lngCancel As Long

    Dim strMessage As String
    Dim bErrorOccurred As Boolean

    Dim oPOM As IJDPOM
    Dim oObject As Object
    Dim vMoniker As Variant
    Dim oCallBack As IJCheckMfctyCallback
    
    ' Setup the Check Manufacturability call back mechanism.
    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    ' If the working set is empty, do not continue.
    If oElements.Count = 0 Then
        Exit Sub
    End If

    ' Get the Pesistent Object Manager.
    Set oPOM = GetPOM("Model")
    
    'loop
    nCount = 0
    nMaxCount = oElements.Count
    
    m_bLogError = False
    For Each vMoniker In oElements

        ' Get the object from the moniker.
        Set oObject = oPOM.GetObject(vMoniker)
        
        ' Check the values.
        bErrorOccurred = CheckObjectForErrorCondition(oObject, strMessage)
        
        If bErrorOccurred Then
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oObject                the object with error/warning
            '   strProgID              prog id of this rule
            '   ESeverity.siError      severity of the error
            '                          (101 error or 102 warning)
            '   105                    code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            '   strAction              type of repair to be made
            '   strActionTool          progid of the COM object that will make the repair
            oCallBack.OnCheckError oObject, strProgID, _
                                   ESeverityIndex.siWarning, 105, strMessage, _
                                   "COM Repair Object", _
                                   "CheckMfctyTeeWeldChamfer.RepairChamfers"
            
            ' The operator may desire that only the first error be found.
            If bStopOnError Then
                Exit For
            End If
        End If
            
        ' Update the progress bar.  First argument is the progress as a
        ' percentage (i.e. 0 - 100).  The second argument is a cancellation
        ' flag.
        nCount = nCount + 1
        oCallBack.OnProgress nCount * 100 / nMaxCount, 0
        
    Next vMoniker

    If m_bLogError Then
        MsgBox "Processing Complete" & vbCrLf & _
               "Errors (and/or Warnings) were detected" & vbCrLf & _
               "see Sp3d ErrorLog for detailings", , _
               "CheckMfctyTeeWeldChamfer.CheckChamfers"
               
    End If
    
    Set oCallBack = Nothing
    Set vMoniker = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, Module, Method).Number

End Sub


'******************************************************************************
' Routine: CheckObjectForErrorCondition
'
' Abstract:
' Checks if the current Object is a Chamfer Object
'   If a Chamfer:
'       Get the owning Assembly Connection
'       Check if the Assembly Connection is type "Plate Edge to Plate Face"
'       If "Plate Edge to Plate Face" type:
'           If based on Thickness of Part on other side of Connection:
'               Get Other Side of Connection Plate Thickness
'               Get "ChamferThickness" parameter value
'               Compare thickness values
'               if same, Tee Weld Chamfer valid
'               else, need update Assembly Connection
'
' Description:
'
' Inputs:
'
' Outputs:
'
'******************************************************************************
Private Function CheckObjectForErrorCondition(oObject As Object, _
                                              ByRef errorMessage As String) As Boolean
Const Method As String = "CheckObjectForErrorCondition"
    On Error GoTo ErrorHandler

    Dim sMessage As String
    
    Dim lCtx As eUSER_CTX_FLAGS
    Dim sChamferThickness As String
    Dim dChamferThickness As Double
    Dim dOtherSideThickness As Double
    
    Dim oNamedItem As IJNamedItem
    Dim oParentObject As Object
    Dim oChamferSelectorObject As Object
    
    Dim oBoundedPort As IJStructPort
    Dim oBoundingPort As IJStructPort
    Dim oPartOnOtherSide As IJElements
    Dim oSelectorChildren As Collection
    
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn
    
    ' Check if current object is type IJChamfer
    CheckObjectForErrorCondition = False
    If Not TypeOf oObject Is IJChamfer Then
        Exit Function
    End If
    
    ' Current Object is an IJChamfer (a SmartItem)
    ' Get the owning Smart Object ( the Smart Object that created it)
    GetParentSmartObject oObject, oParentObject
    
    ' Check if Parent Smart Object is type IJAssemblyConnection
    If oParentObject Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oParentObject Is IJAssemblyConnection Then
        Exit Function
    End If
    
    ' From the Assembly Connection
    ' Check if Assembly Connection is a Plate Edge to Plate Face (Tee Weld case)
    '   the Bounded Port Must be a Plate Edge
    '   the Bounding Port Must be a Plate Face
    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssemblyConn.object = oParentObject
    
    If TypeOf oSDO_AssemblyConn.ConnectedObject1 Is IJPlate Then
        Set oBoundedPort = oSDO_AssemblyConn.Port1
        lCtx = oBoundedPort.ContextID
        If (lCtx And CTX_LATERAL) = 0 Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    If TypeOf oSDO_AssemblyConn.ConnectedObject2 Is IJPlate Then
        Set oBoundingPort = oSDO_AssemblyConn.Port2
        lCtx = oBoundingPort.ContextID
        If (lCtx And CTX_LATERAL) <> 0 Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    ' Know that this is Plate Edge Bounded by Plate Face case
    ' Want to determine if the Bounded Plate is Split by the Bounding Plate
    ' such that Leaf Part(s) with the same Root Part as the bounded Part
    ' exist on the "other" side of the Bounding Plate
    Set oPartOnOtherSide = oSDO_AssemblyConn.PartsOnOtherSide
    If oPartOnOtherSide Is Nothing Then
        Exit Function
    ElseIf oPartOnOtherSide.Count < 1 Then
        Exit Function
    End If
        
    ' The Chamfer "Nose" is based on the Plate on the Other Side Thickness
    ' Get Thickness of the Plate on the Other Side
    Set oSDO_PlatePart = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart.object = oPartOnOtherSide.Item(1)
    dOtherSideThickness = oSDO_PlatePart.PlateThickness
    
    ' In the delivered set of Rules
    ' The Chamfer Thickness is passed down to the "RootTeeChamferSel"
    ' This value is set to a default value by the Assembly Connection
    ' The Assembly Connection value can be over-ridened by the User
    ' We can get the 'Chamfer Thickness' from the "RootTeeChamferSel"
    ' Note: We do not check if the default value has been over-ridden
    On Error Resume Next
    sChamferThickness = ""
    GetSmartSelector oObject, m_sRootTeeChamferSel, _
                     oSelectorChildren
    If Not oSelectorChildren Is Nothing Then
        If oSelectorChildren.Count > 0 Then
            Set oChamferSelectorObject = oSelectorChildren.Item(1)
            GetSmartOccAnswer oObject, oChamferSelectorObject, _
                              m_sQuestion_ChamferThickness, sChamferThickness
        End If
    End If
    
    ' If for some reason we failed to get the "ChamferThickness" from
    ' the Chamfer's Selector
    ' Attempt to retrieve it from the Assembly Connection
    If Len(Trim(sChamferThickness)) < 1 Then
        ' In the delivered set of Rules
        ' The Chamfer Thickness is set in the Assembly Connection
        ' But this can be over-ridened by the User
        ' We can get the 'Chamfer Thickness' from the Assembly Connection
        ' Note: We do not check if the default value has been over-ridden
        Set oChamferSelectorObject = Nothing
        GetSmartOccAnswer oParentObject, oChamferSelectorObject, _
                          m_sQuestion_ChamferThickness, sChamferThickness
    End If
    
    ' Verify the "ChamferThickness"  is valid
    If Len(Trim(sChamferThickness)) < 1 Then
        sMessage = "Could not retrieve ChamferThickness Answer from Selection Rules"
        
        If TypeOf oObject Is IJNamedItem Then
            Set oNamedItem = oObject
            sMessage = sMessage & vbCrLf & "Chamfer:" & oNamedItem.Name
            Set oNamedItem = Nothing
        End If
        
        If TypeOf oParentObject Is IJNamedItem Then
            Set oNamedItem = oParentObject
            sMessage = sMessage & vbCrLf & "AssemblyConnection:" & oNamedItem.Name
            Set oNamedItem = Nothing
        End If
        
        Err.Raise LogError(Err, Module, Method, sMessage).Number
        
        Exit Function
    End If
    
    ' Compare the OtherSideThickness with the current Chamfer Thickness
    dChamferThickness = sChamferThickness
    If Abs(dChamferThickness - dOtherSideThickness) < m_dDeltaThickness Then
        Exit Function
    End If
        
    ' The Chamfer Thickness does not match the Plate's OtherSideThickness
    ' This could be because the:
    '   Other Plate Thickness has been modified after the Chamfer was created
    '   User manually modified the "Chamfer Thickness" value
    ' We will flag both cases as "Error" conditions
    CheckObjectForErrorCondition = True
    errorMessage = "Chamfer Thickness(" & _
                   Format(dChamferThickness, "0.0000") & ") <> " & _
                   "Other Side Plate Thickness(" & _
                   Format(dOtherSideThickness, "0.0000") & ")"
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, Method, sMessage).Number

End Function
