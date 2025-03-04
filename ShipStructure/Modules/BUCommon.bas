Attribute VB_Name = "BUCommon"
Option Explicit

'******************************************************************
' Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'File
'    DesignedMemberCommon.cls
'
'Author
'    Richard Solomon
'
'Description
'   Common definitions
'
'Notes
' Would prefer that only one 'BUCommon.bas' exist, but there is no common
' folder between dev and shared content, so two versions will have to exist for now
' once this is replaced by the .Net version, there should be no duplication
'
'History:
'
'   22-Sept-2009 GG TR#167167 - DesignedMember does not set PG of plates to its own PG
'   16-Oct-2009 GG DM#173407  BuiltUps write numerous Info messages to error log
'*******************************************************************

Public Const distTol = 0.00001
Public Const PI = 3.14159265358979


Public Const E_FAIL = -2147467259
Public Const E_INVALIDARG = -2147942487#

Public Const IJDAttributes = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"

'============ test code... ====================================================
Public Const IID_IUABuiltUpLengthExt = "IUABuiltUpLengthExt"
Public Const IID_IUABuiltUpCompute = "IUABuiltUpCompute"
Public Const IID_IUABuiltUpWeb = "IUABuiltUpWeb"
Public Const IID_IUABuiltUpTopFlange = "IUABuiltUpTopFlange"
Public Const IID_IUABuiltUpBottomFlange = "IUABuiltUpBottomFlange"
Public Const IID_IUABuiltUpC = "IUABuiltUpC"
Public Const IID_IUABuiltUpL = "IUABuiltUpL"
Public Const IID_IUABUBoxFlangeMajor = "IUABUBoxFlangeMajor"
Public Const IID_IUABuiltUpBoxWebMajor = "IUABuiltUpBoxWebMajor"
Public Const IID_IUABuiltUpBoxComb = "IUABuiltUpBoxComb"
Public Const IID_IUABuiltUpIHaunch = "IUABuiltUpIHaunch"
Public Const IID_IUABuiltUpITaperWeb = "IUABuiltUpITaperWeb"
Public Const IID_IUABuiltUpTube = "IUABuiltUpTube"
Public Const IID_IUABuiltUpCone = "IUABuiltUpCone"
Public Const IID_IUABuiltUpCan = "IUABuiltUpCan"
Public Const IID_IUABuiltUpEndCan = "IUABuiltUpEndCan"
Public Const IDesignedMembPlateChanged = "{47776E5F-7777-49CD-811F-24084BEDEC20}"
Public Const IStructCrossSection = "{F92D39CB-6BE8-48a6-9AEC-0ED19773B3CA}"
Public Const IStructCrossSectionDimensions = "{0F02B4A9-8D07-4dcc-A764-CC7DB12425DE}"
Public Const IStructCrossSectionDesignProperties = "{A2A1F140-CFE4-4e12-A23D-A1245EDFDDE6}"
Public Const IJCurve = "{A7A36610-C800-11d1-9555-0060973D4824}"
Public Const IStructCrossSectionUnitWeight = "{6238E751-24BD-488C-BA5E-93F27E7EA922}"
Public Const IID_IUABuiltUpCone1 = "IUABuiltUpCone1"
Public Const IID_IUABuiltUpCone2 = "IUABuiltUpCone2"
Public Const IID_IJUASMCanRuleResult = "IJUASMCanRuleResult"
Public Const IID_IUABuiltUpRectTube = "IUABuiltUpRectTube"
'===============================================================================

Public Const AssemblyMembers1RelationshipCLSID = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}"
Public Const IJSmartOccurrence = "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}"
Public Const DESINGEDMEMMBERPROGID = "SPSMembers.SPSDesignedMember"

Public Const attrCanType As String = "CanType"
Public Const CanType_InLine = 1
Public Const CanType_StubEnd = 2
Public Const CanType_End = 3
Public Const dTol As Double = 0.000001

'===============================================================================
' Resource IDs
' ... Copied from:
' ... \ShipStructure\Middle\Services\BuiltUpDefinitions\L10N\Include\L10Nresource.bas
Public Const IDS_BUILTUP_ERROR = 104
Public Const IDS_BUILTUP_SCHEMAERROR = 105
Public Const IDS_BUILTUP_VALUE_MUSTBE_POSITIVE = 106
Public Const IDS_BUILTUP_VALUE_MUSTBE_GREATERTHAN_OR_EQUAL_TO_ZERO = 107
Public Const IDS_BUILTUP_VALUE_MUSTBE_IN_RANGE = 108
'*************************************************************************
'Function
'HandleAndRaiseError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and raises it back up the call stack
'
'input
'Module(file) name, method name
'
'Return
'success
'
'Exceptions
'
'***************************************************************************
Public Sub HandleAndRaiseError(sModule As String, sMethod As String)
    
    Dim errNo As Long
    errNo = Err.Number
    Dim ErrDesc As String
    ErrDesc = Err.Description
   
    HandleError sModule, sMethod
    Err.Raise errNo, sModule + "::" + sMethod, ErrDesc
End Sub
'*************************************************************************
'Function
'HandleError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and returns success
'
'input
'Module(file) name, method name
'
'Return
'success
'
'Exceptions
'
'***************************************************************************
Public Sub HandleError(sModule As String, sMethod As String)
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    Set oEditErrors = Nothing
End Sub

'*************************************************************************
'Function
'SPSToDoErrorNotify
'
'Abstract
' Called to notify the SmartOccurrence of a ToDo error that occurred during a
' smart occurrence custom evaluate
'
'History
'
'***************************************************************************
Public Sub SPSToDoErrorNotify(strCodelistTable As String, nToDoListErrorNum As Long, oObjectInError As Object, oObjectToUpdate As Object)
    Const METHOD = "SPSToDoErrorNotify"
    On Error GoTo ErrHandler:
    Dim oToDoListHelper As IJToDoListHelper ' Set ToDoListHelper = pointer to the CAO Object

    Set oToDoListHelper = oObjectInError
    If Not oToDoListHelper Is Nothing Then
        If oObjectToUpdate Is Nothing Then
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum
        Else
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum, oObjectToUpdate
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    ' Report the error but do not return the error since this call maybe in their
    '   error handler
    HandleError "CommonError", METHOD
    Err.Clear
End Sub

Public Sub CopyPermissionGroup(ByRef ObjTarget As Object, ObjSource As Object)
    Const METHOD = "CopyPermissionGroup"
    On Error GoTo ErrHandler:
    Dim opgObjTarget As IJDObject
    Dim opgobjSource As IJDObject
    Set opgObjTarget = ObjTarget
    Set opgobjSource = ObjSource
    If opgObjTarget Is Nothing Or opgobjSource Is Nothing Then
        Err.Raise E_INVALIDARG, "BUCommon", "NULL input arguments"
    End If
    opgObjTarget.PermissionGroup = opgobjSource.PermissionGroup
    Exit Sub
    
ErrHandler:
    ' Report and raise the error. Otherwise, a caller will have no idea if it failed
    
    HandleAndRaiseError "BUCommon", METHOD
    Err.Clear
End Sub


Public Sub GetComponentMaterial(ByVal pProfileObject As Object, ComponentInterfaceName As String, Material As String, Grade As String, Thickness As Double)
    Const METHOD = "GetComponentMaterial"
    On Error GoTo ErrHandler:
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Set oSmartOcc = pProfileObject
    Set oSmartItem = oSmartOcc.ItemObject
    
    Dim oAttributeMetaData As IJDAttributeMetaData
    Set oAttributeMetaData = oSmartItem
    'Dim oInterfaceInfo  As IJDInterfaceInfo
    
    Dim interfaceIID As Variant
    interfaceIID = oAttributeMetaData.IID(ComponentInterfaceName)
    'Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo()
   
    Dim oInfoCollection As IJDInfosCol
    Set oInfoCollection = oAttributeMetaData.InterfaceAttributes(interfaceIID, Mask_Expose_SQLOrCOM, False)
    
    ' get the names of the material, grade and thickness attributes
    Dim oAttrInfo As IJDAttributeInfo
    Dim materialAttr As String
    Dim gradeAttr As String
    Dim thicknessAttr As String
    For Each oAttrInfo In oInfoCollection
        If InStr(oAttrInfo.Name, "Material") > 0 Then
            materialAttr = oAttrInfo.Name
        ElseIf InStr(oAttrInfo.Name, "Grade") > 0 Then
            gradeAttr = oAttrInfo.Name
        ElseIf InStr(oAttrInfo.Name, "Thickness") > 0 Then
            thicknessAttr = oAttrInfo.Name
        End If

    Next
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    Set oAttr = oSmartItem

    Set oAttrCol = oAttr.CollectionOfAttributes(ComponentInterfaceName)
    If Not oAttrCol Is Nothing Then
        Material = oAttrCol.Item(materialAttr).Value
        Grade = oAttrCol.Item(gradeAttr).Value
        Thickness = oAttrCol.Item(thicknessAttr).Value
    End If
    Exit Sub
ErrHandler:
    ' Report and raise the error. Otherwise, a caller will have no idea if it failed
    HandleAndRaiseError "BUCommon", METHOD
End Sub



Public Function CreateComplexStringFromPositions(oPos() As IJDPosition) As IJComplexString
    
    Dim curveElms As IJElements
    Dim idx As Long
    Dim pGeometryFactory As New GeometryFactory
    Dim numPos As Long
    
    numPos = UBound(oPos) - LBound(oPos) + 1
    Set curveElms = New JObjectCollection

    For idx = 1 To numPos
        Dim oLine3d As Line3d
        Dim nextIdx As Integer
        nextIdx = idx Mod numPos
        Set oLine3d = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oPos(idx).x, oPos(idx).y, oPos(idx).z, _
        oPos(nextIdx + 1).x, oPos(nextIdx + 1).y, oPos(nextIdx + 1).z)
        
        curveElms.Add oLine3d
        Set oLine3d = Nothing
    Next idx
    
    Set CreateComplexStringFromPositions = pGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
    Set curveElms = Nothing

End Function


