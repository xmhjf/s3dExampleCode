VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatumShapeNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project: EquipNamingRules
'
'  Abstract: The file contains naming rule implementation for Shape
'
'  Author: Samba
'
'   History:
'       21 Jul, 2005    Samba           Initial Creation
'**************************************************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "DatumShapeNameRule: "
Private Const MODELDATABASE = "Model"
Private Const PARTROLE = "part"

Dim m_oErrors As IJEditErrors
Dim m_oLocalizer As IJLocalizer

Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & App.EXEName
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oLocalizer = Nothing
    Set m_oErrors = Nothing
End Sub


'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the parents
' name and object name.It is something like this: "Base Name" + "Object Name" + Index.
' "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name"
' It is assumed that all Naming Parents and the Object implement IJNamedItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName"
On Error GoTo ErrorHandler

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If
    Dim oNamedItem As IJNamedItem
    Dim strName As String
    Dim strNameBasis As String
    Dim strPartname As String
    Dim nCount As Long
    Dim strExistingName As String
    Dim bIsNameGenRequired As Boolean
    
    Set oNamedItem = oObject
    strExistingName = oNamedItem.Name
    bIsNameGenRequired = False
    ' Use part Name for basename, remove blanks:
    strPartname = m_oLocalizer.GetString(IDS_STRING_DATUMSHAPE, "DP")
    bIsNameGenRequired = IsExistingNameBad(strExistingName, strPartname)
    
    If bIsNameGenRequired = True Then
        ' no parents - use just the strPartName (DP) as the base
        strNameBasis = oActiveEntity.NamingParentsString
        'Check if New parent name string constructed and old parent name string existing are same
        'if they are the same, we do not need to generate a new name
        If strPartname <> strNameBasis Then
            oActiveEntity.NamingParentsString = strPartname
            nCount = GetRelatedObjects(oObject)
            strName = strPartname + Format(nCount)
            oNamedItem.Name = strName
        End If
    End If
    Set oNamedItem = Nothing
    
Exit Sub
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' For the Generic rule, the naming parent would be the part object if the named object
' is a part occurrence object.  However, the naming parent relationship is not
' defined to cross database boundaries.  For now there will not be any naming parent
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oUnkParentEqp As IUnknown
    Dim oShape As IJShape

    On Error Resume Next
    Set oShape = oEntity
    oShape.GetParent oUnkParentEqp

    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    If Not (oUnkParentEqp Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oUnkParentEqp)
    End If

    Set oUnkParentEqp = Nothing
    Set oShape = Nothing
Exit Function

ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

Private Function GetRelatedObjects(oObject As IJDObject) As Long
    Const METHODNAME = "GetRelatedObjects"
    On Error GoTo ErrHandler
    
    Dim oColl As IJDObjectCollection
    Dim oShape As IJShape
    Dim oEquipUnk As IUnknown
    Dim oDesignObject As IJGenericDesignObject
    Dim oNI As IJNamedItem
    Dim oShapeUnk As IUnknown
    Dim strDatumShape As String
    Dim iFirstFound As Integer
    Dim strName As String
    Dim lTemp As Long, lCount As Long
    
    GetRelatedObjects = 1
    Set oShape = oObject
    oShape.GetParent oEquipUnk
    Set oDesignObject = oEquipUnk
    oDesignObject.GetChildren oColl
    
    On Error Resume Next
    
    strDatumShape = m_oLocalizer.GetString(IDS_STRING_DATUMSHAPE, "DP")
    For Each oShapeUnk In oColl
        Set oShape = oShapeUnk
        If Not oShape Is Nothing Then
            Set oNI = oShape
            strName = oNI.Name
            iFirstFound = InStr(1, strName, strDatumShape)
            If iFirstFound = 1 Then
                lTemp = CInt(Mid(strName, 3, Len(strName) - 2))
                If lCount < lTemp Then
                    lCount = lTemp
                End If
            End If
            Set oNI = Nothing
            Set oShape = Nothing
        End If
    Next oShapeUnk
    
    Set oShapeUnk = Nothing
    Set oDesignObject = Nothing
    Set oEquipUnk = Nothing
    Set oColl = Nothing
    GetRelatedObjects = lCount + 1
    
Exit Function
ErrHandler:
    m_oErrors.Add Err.Number, "GetRelatedObjects", Err.Description & " : " & METHODNAME
End Function


Private Function IsExistingNameBad(ByVal strCurrName As String, ByVal strPreFix As String) As Boolean
Const METHODNAME = "IsExistingNameBad"
On Error GoTo ErrHandler
IsExistingNameBad = True ' we assume every name is bad so as to regenerate
    If strCurrName Like "" Then
        Exit Function
    End If
    Dim lFound As Long
    Dim lUsedIndex As Long
    'Check with the base name Name of DP in this rule is DP + Integer number of DP shape.
    lFound = CLng(InStr(1, strCurrName, strPreFix, vbTextCompare))
    If lFound = 1 Then
        lUsedIndex = CLng(Mid(strCurrName, (Len(strPreFix) + 1)))
        If IsNumeric(lUsedIndex) Then
            IsExistingNameBad = False
        End If
    End If
Exit Function
ErrHandler:
    IsExistingNameBad = True
    m_oErrors.Add Err.Number, "IsExistingNameBad", Err.Description & " : " & METHODNAME
End Function
