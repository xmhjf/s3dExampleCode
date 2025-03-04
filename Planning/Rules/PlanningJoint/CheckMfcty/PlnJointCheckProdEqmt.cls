VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnJointCheckProdEqmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2004 Intergraph Corporation.  All Rights Reserved.
'
' File: PlnJointCheckProdEqmt.cls
'
' Author: Dick Swager
'
' Abstract: Check manufacturability rules for Planning Joints.
'
' Description:
'
'********************************************************************

Option Explicit

Private Const Module = "PlnCheckMfcty.PlnJointCheckProdEqmt"
Private Const PROG_ID = "PlnCheckMfcty.PlnJointCheckProdEqmt"

Private Const CATDB = "Catalog"

Implements IJCheckMfcty

'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description:
' Check manufacturability will also need to make sure that the planning joint
' can be manufactured in the assembly it is assigned to. It is possible that
' the user could drag the planning joint to an assembly that cannot perform
' the weld.
'
' PseudoCode
'For each IJPlnJoint
'    Call the GetAssembly() method to retrieve the Assembly
'    From the Assembly, retrieve the Production Equipment (Welding equipment)
'    Check also that the IJPlnJoint does not have a production equipment
'        specified as an override
'    Report a Warning if Welding type (on the PC) is actually Single Sided
'        and the Production Equipment can only do Double-sided welding.
'End for
'
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck _
       (ByVal oCollection As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        pOptionCodes() As Variant, _
        ByVal pCallingObj As Object)

    Const Method = "IJCheckMfcty_Execute"
    On Error GoTo ErrorHandler

    ' Setup the Check Manufacturability call back mechanism.
    Dim oCallBack As IJCheckMfctyCallback
    Dim lngDummy As Long
    Dim lngCancel As Long

    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    ' If the working set is empty, do not continue.
    If oCollection.Count = 0 Then GoTo Cleanup

    ' Get all the planning joints from the collection
    Dim oJoints As IJElements
    GatherPlanningJoints oCollection, oJoints
    
    ' Verify that the manufacturing equipment for each Planning Joint
    ' Assembly is capable to perform the required welds.
    Dim oPlnJoint As IJPlnJoint
    For Each oPlnJoint In oJoints
    
        ' Get the assembly that contains the Planning Joint.
        Dim oAssyPlnJoint As IJAssembly
        Set oAssyPlnJoint = oPlnJoint.GetAssembly
        
        ' Retrieve the production equipment for the Assembly.
        Dim oPlnAssociateCatalog As IJPlnAssociateCatalog
        Set oPlnAssociateCatalog = oAssyPlnJoint
        
        Dim oProdEquipment As IJSRDProdEquipOcc
        Set oProdEquipment = oPlnAssociateCatalog.GetWeldEquipment
                        
        ' Set named items in the event of failure.
        Dim oNamedItemJoint As IJNamedItem
        Set oNamedItemJoint = oPlnJoint
        Dim oNamedItemAssembly As IJNamedItem
        Set oNamedItemAssembly = oAssyPlnJoint
            
        ' Verify that there is production equipment assigned to the
        ' assembly.
        Dim strErrMsg As String
        If oProdEquipment Is Nothing Then
        
            strErrMsg = "Production Equipment is not assigned to the assembly" & vbNewLine _
                    & "Planning Joint: " + oNamedItemJoint.Name & vbNewLine _
                    & "Assembly: " + oNamedItemAssembly.Name

                    
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oPlnJoint              the object with error/warning
            '   PROG_ID                prog id of this rule
            '   ESeverity.siError      severity of the error (101 error or 102 warning)
            '   10                     code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   ESeverity.siError, 10, strErrMsg, "", ""
        
            GoTo Cleanup
        End If
        
        ' Report a Warning if Welding type (on the PC) is actually Single
        ' Sided and the Production Equipment can only do Double-sided
        ' welding.
        strErrMsg = ""
        Dim severity As ESeverity
        If Not CheckWeldSidedness _
                (oPlnJoint, oProdEquipment, strErrMsg, severity) Then
        
            ' An error condition exits.
            strErrMsg = strErrMsg & vbNewLine _
                    & "Planning Joint: " + oNamedItemJoint.Name & vbNewLine _
                    & "Production Equipment Occurence description: " + _
                        oProdEquipment.ProdEquipOccDesc & vbNewLine _
                    & "Production Equipment Occurence name: " + _
                        oProdEquipment.ProdEquipOccName & vbNewLine _
                    & "Production Equipment Part name: " + _
                        oProdEquipment.ProdEquipPartName & vbNewLine _
                    & "Assembly: " + oNamedItemAssembly.Name

                    
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oPlnJoint              the object with error/warning
            '   PROG_ID                prog id of this rule
            '   ESeverity.siError      severity of the error (101 error or 102 warning)
            '   10                     code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   severity, 10, strErrMsg, "", ""
        
        End If
        
'        ' Verify that the production equipment can perform the weld.
'        strErrMsg = ""
'        If Not EquipmentCanPerformWeld(oProdEquipment, strErrMsg) Then
'
'            ' An error condition exits.
'            strErrMsg = strErrMsg & vbNewLine _
'                    & "Planning Joint: " + oNamedItemJoint.Name & vbNewLine _
'                    & "Production Equipment Occurence description: " + _
'                        oProdEquipment.ProdEquipOccDesc & vbNewLine _
'                    & "Production Equipment Occurence name: " + _
'                        oProdEquipment.ProdEquipOccName & vbNewLine _
'                    & "Production Equipment Part name: " + _
'                        oProdEquipment.ProdEquipPartName & vbNewLine _
'                    & "Assembly: " + oNamedItemAssembly.Name
'
'
'            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
'            '   oPlnJoint              the object with error/warning
'            '   PROG_ID                prog id of this rule
'            '   ESeverity.siError      severity of the error (101 error or 102 warning)
'            '   10                     code of this rule (found in spreadsheet
'            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
'            '   strErrMsg              message for the error
'            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
'                                   ESeverity.siError, 10, strErrMsg
'
'        End If
        
        ' Verify that the production equipment matches the workcenter
        ' equipment.
        strErrMsg = ""
        If Not EquipmentMatchesWorkcenterEquipment _
                (oPlnJoint, oProdEquipment, strErrMsg) Then
        
            ' An error condition exits.
            strErrMsg = strErrMsg & vbNewLine _
                    & "Planning Joint: " + oNamedItemJoint.Name & vbNewLine _
                    & "Production Equipment Occurence description: " + _
                        oProdEquipment.ProdEquipOccDesc & vbNewLine _
                    & "Production Equipment Occurence name: " + _
                        oProdEquipment.ProdEquipOccName & vbNewLine _
                    & "Production Equipment Part name: " + _
                        oProdEquipment.ProdEquipPartName & vbNewLine _
                    & "Assembly: " + oNamedItemAssembly.Name

                    
            ' The arguments for IJCheckMfctyCallback.OnCheckError are:
            '   oPlnJoint              the object with error/warning
            '   PROG_ID                prog id of this rule
            '   ESeverity.siError      severity of the error (101 error or 102 warning)
            '   10                     code of this rule (found in spreadsheet
            '                          M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
            '   strErrMsg              message for the error
            oCallBack.OnCheckError oPlnJoint, PROG_ID, _
                                   ESeverity.siError, 10, strErrMsg, "", ""
        
        End If
        
    Next oPlnJoint
        
    GoTo Cleanup
    
ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
    Set oCallBack = Nothing
    Set oJoints = Nothing
    Set oPlnJoint = Nothing
    Set oAssyPlnJoint = Nothing
    Set oPlnAssociateCatalog = Nothing
    Set oProdEquipment = Nothing
    Set oNamedItemJoint = Nothing
    Set oNamedItemAssembly = Nothing

End Sub



'******************************************************************************
' Routine: EquipmentMatchesWorkcenterEquipment
'
' Abstract: Check that the Planning Joint weld equipment matches the weld
'   equipment of the Workcenter.
'
' Description:
' Compare WeldingEquipement assigned to the Work center to which the part
' belongs with Equipment on Planning joint.
'
'******************************************************************************
Private Function EquipmentMatchesWorkcenterEquipment _
        (oPlnJoint As IJPlnJoint, _
         oProdEquipment As IJSRDProdEquipOcc, _
         strMessage As String) As Boolean
    
    Const Method = "EquipmentMatchesWorkcenterEquipment"
    On Error GoTo ErrorHandler

    ' Initial values in case of unexpected error.
    EquipmentMatchesWorkcenterEquipment = False
    strMessage = "Planning Joint weld equipment does not match the equipment in the workcenter."

    ' Get the Assembly from the Planning Joint.
    Dim oAssembly As IJAssembly
    Set oAssembly = oPlnJoint.GetAssembly
    
    ' Get the Workcenter for the Assembly.
    If Not oAssembly Is Nothing Then
        Dim oPlnAssociateCatalog As IJPlnAssociateCatalog
        Set oPlnAssociateCatalog = oAssembly
        
        If Not (oPlnAssociateCatalog Is Nothing) Then
            Dim oWorkcenter As IJSRDWorkcenter
            Set oWorkcenter = oPlnAssociateCatalog.GetWorkcenter
            If oWorkcenter Is Nothing Then
                EquipmentMatchesWorkcenterEquipment = False
                strMessage = "No Workcenter assigned to the Assembly."
                GoTo Cleanup
            End If
        End If
    End If
    
    ' Get the name of the Workcenter.
    Dim strWorkcenterName As String
    strWorkcenterName = oWorkcenter.WorkcenterName
    If Len(strWorkcenterName) = 0 Then
        EquipmentMatchesWorkcenterEquipment = False
        strMessage = "Workcenter name is blank."
        GoTo Cleanup
    Else
        strMessage = strMessage & vbNewLine & _
                "Workcenter name: " & strWorkcenterName
    End If
    
    ' Create a Workcenter Query object to get information about the
    ' workcenter.
    Dim oSRDQuery As IJSRDQuery
    Set oSRDQuery = New SRDQuery
    Dim oWorkcenterQuery As IJSRDWorkcenterQuery
    Set oWorkcenterQuery = oSRDQuery.GetWorkcenterQuery()

    ' Get the equipment that is assigned to the Workcenter.
    Dim strListOfEquipment() As String
    oWorkcenterQuery.GetEquipmentInWorkcenter strWorkcenterName, strListOfEquipment
    
    ' Check each equipment in the Workcenter for a match with the
    ' welding equipment for the Planning Joint.
    
    Dim i As Integer
    For i = LBound(strListOfEquipment) To UBound(strListOfEquipment)
    
        Dim strEquipment As String
        strEquipment = strListOfEquipment(i)
        
        If strEquipment = oProdEquipment.ProdEquipOccName Then
        
' Note: it may be that we want to check more than the name.  If so then
' the following code applies.  When I tried this though I had different
' objects with the same name and the check failed.  CRS 04/01/2004
'        ' Get the actual equipment from the name.
'        Dim oEquipment As IJSRDProdEquipOcc
'        oWorkcenterQuery.GetEquipOccObject strWorkcenterName, _
'                strEquipment, oEquipment
'
'        ' After a match is found there is no need to continue.
'        If oEquipment Is oProdEquipment Then
' End of note.

            EquipmentMatchesWorkcenterEquipment = True
            GoTo Cleanup
        End If
'        Set oEquipment = Nothing
        
    Next i
    
    GoTo Cleanup
    
ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
    Set oAssembly = Nothing
    Set oPlnAssociateCatalog = Nothing
    Set oWorkcenter = Nothing
    Set oSRDQuery = Nothing
    Set oWorkcenterQuery = Nothing
'    Set oEquipment = Nothing

End Function


'******************************************************************************
' Routine: CheckWeldSidedness
'
' Abstract: Check the Welding type on the Physical Connection compared to
'   the Production Equipment with regard to sidedness (i.e. Single Sided /
'   Double Sided).
'
' Description:
'   Not yet implemented.
'******************************************************************************
Private Function CheckWeldSidedness _
        (oPlnJoint As IJPlnJoint, _
         oProdEquipment As IJSRDProdEquipOcc, _
         strMessage As String, _
         severity As ESeverity) As Boolean
    
    Const Method = "CheckWeldSidedness"
    On Error GoTo ErrorHandler

    ' Initialization in the event of an unexpected error.
    CheckWeldSidedness = False
    strMessage = "Weld equipment failed sidedness test."
    severity = ESeverity.siError
    
    ' Create the physical connection wrapper for the Planning Joint.
    Dim oPlnPhysConn As PlanningObjects.PlnPhysConn
    Set oPlnPhysConn = New PlanningObjects.PlnPhysConn
    Set oPlnPhysConn.Object = oPlnJoint.GetPhysicalConnection

    ' Get the weld type.
    Dim strWeldType As String
    strWeldType = oPlnPhysConn.WeldType
    
    ' Determine sidedness out of the weld type.  Typical values for the
    ' weld type are "ButtWeldSingleSided", "CornerWeldSingleSided",
    ' "TeeWeldDoubleSided", "LapWeldSingleSided", "TeeWeldSingleSided"
    ' and "ButtWeldSingleSided"
    Dim bSingleSidedPC As Boolean
    Dim bDoubleSidedPC As Boolean
    bSingleSidedPC = (InStr(strWeldType, "SingleSided") <> 0)
    bDoubleSidedPC = (InStr(strWeldType, "DoubleSided") <> 0)
    
    ' Determine the weld machine capabilities with regard to sidedness.
    ' Prepare to ask for the MachineType of the equipment.
    Const strUserInterfaceName As String = "IJUAWeldingEquipment"
    Const strAttributeName As String = "MachineType"
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String

    ' Use StructDetailObjects.Helper to get the MachineType.
    Dim oHelper As StructDetailObjects.Helper
    Set oHelper = New StructDetailObjects.Helper
        
    Dim oPartDisp       As Object
    Dim strPartName     As String
    Dim oCatalogDMO     As IJDCatalogDMO
    Dim oPOM            As IJDPOM
        
    strPartName = oProdEquipment.ProdEquipPartName
    
    Set oCatalogDMO = New CatalogDMO
    
    GetPOM CATDB, oPOM
    
    'Get the Part Object from the given Part Name
    Set oPartDisp = oCatalogDMO.GetSpecGivenName(oPOM, strPartName)
    
    Call oHelper.GetCustomAttributeParamValue _
            (oPartDisp, strUserInterfaceName, strAttributeName, _
            retValue, retValueType, sTableName, _
            lCodeList, sShortDesc, sLongDesc)
            
    
    ' Check if answer is from code list; if so, get short desc, if not,
    ' use retValue.
    Dim strMachineType As String
    If Len(Trim(sTableName)) > 0 Then
        strMachineType = sShortDesc
    Else
        strMachineType = retValue
    End If
    
    If InStr(strMachineType, "Sided") = 0 Then
        CheckWeldSidedness = False
        strMessage = "Weld equipment failed sidedness test." & vbNewLine & _
                "MachineType sidedness is unknown" & vbNewLine & _
                "MachineType: " & strMachineType & vbNewLine & _
                "WeldType: " & strWeldType
        severity = ESeverity.siWarining
        
        GoTo Cleanup
    End If
    
    ' Compare MachineType to WeldType.
    If (strMachineType = "OneSided" And bDoubleSidedPC) _
            Or (strMachineType = "TwoSided" And bSingleSidedPC) Then
    
        CheckWeldSidedness = False
        strMessage = "Weld equipment failed sidedness test." & vbNewLine & _
                "MachineType: " & strMachineType & vbNewLine & _
                "WeldType: " & strWeldType
        severity = ESeverity.siError
        
        GoTo Cleanup
        
    End If
    
    CheckWeldSidedness = True
    
    GoTo Cleanup
    
ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
    Set oPlnPhysConn = Nothing
    Set oHelper = Nothing
    Set oCatalogDMO = Nothing
    Set oPartDisp = Nothing
    Set oPOM = Nothing
    

End Function


'******************************************************************************
' Routine: EquipmentCanPerformWeld
'
' Abstract: Check that the Planning Joint weld equipment can perform the
'   required weld.
'
' Description:
'   Not yet implemented.
'******************************************************************************
Private Function EquipmentCanPerformWeld _
        (oProdEquipment As IJSRDProdEquipOcc, _
         strMessage As String) As Boolean
    EquipmentCanPerformWeld = False
    strMessage = "Weld equipment cannot perform the required weld."
End Function

Private Sub GetPOM(strDbType As String, oPOM As IJDPOM)
Const METHODNAME = "GetPOM"

    Dim oAccessMiddle   As IJDAccessMiddle

    Set oAccessMiddle = GetJContext().GetService("ConnectMiddle")
    Set oPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    Set oAccessMiddle = Nothing
    
Exit Sub
End Sub
