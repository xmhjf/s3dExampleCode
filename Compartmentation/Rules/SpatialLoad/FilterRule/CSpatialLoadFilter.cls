VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSpatialLoadFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'Abstract:
' Filter Function for Creation of spatial load
'Description: This class implements the Filter function for spatial load.
'
'Notes:
'
'History
'       AppaRao         17th Jun 04                 Creation.
'       Arnab           23rd Jul 04                 Major Modification
'*********************************************************************

Option Explicit

Private Const MODULE = "CSpatialLoadFilter"
Private Const CATALOG = "Catalog"
Private Const PARTCLASS = "PartClass"

Private m_vEnumAllowableLoads               As IJDEnumMoniker
Private m_oPOM                              As IJDPOM

Implements IJSpatialFilter

Private Function IJSpatialFilter_FilterObject(ByVal Object As Variant) As Boolean
Const METHOD = "IJSpatialFilter_FilterObject"
On Error GoTo ErrorHandler

    Dim vMoniker                As Variant
    Dim oAllowableobj           As Object
    Dim oPartClass              As IJDPartClass
    Dim oClassItem              As IJClassItem
    
    IJSpatialFilter_FilterObject = False
    
    Set oClassItem = Object
    
    If oClassItem Is Nothing Then   'The input is not a ClassItem
        GoTo wrapup
    End If
    
    If m_vEnumAllowableLoads Is Nothing Then
        IJSpatialFilter_FilterObject = True 'No Argument set. So pass the Filter
        GoTo wrapup
    End If
    
    If m_vEnumAllowableLoads.Count = 0 Then     'Nothing is allowable. Fail the Filter
        GoTo wrapup
    End If
    
    If m_oPOM Is Nothing Then   'The function SetFilterArgument may not have been called.
        Set m_oPOM = GetPOM(CATALOG)
    End If
    
    Set oPartClass = GetPartClass(oClassItem)
    
    'if the part class name is present in the load type collection allow the part
    For Each vMoniker In m_vEnumAllowableLoads
        Set oAllowableobj = m_oPOM.GetObject(vMoniker)  'Get each allowable PartClass Object
        If oPartClass Is oAllowableobj Then
            IJSpatialFilter_FilterObject = True
            Exit For
        End If
    Next
    
wrapup:
    Set vMoniker = Nothing
    Set oAllowableobj = Nothing
    Set oPartClass = Nothing
    Set oClassItem = Nothing
        
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

'This sub will be called from the command.
'BAsed on the compartment part type this method gets the
'Available load types and add the collection to m_vEnumAllowableLoads
'Gives the Load type collection based on the compartment type
Private Sub IJSpatialFilter_SetFilterArgument(ByVal Object As Variant)
Const METHOD = "IJSpatialFilter_SetFilterArgument"
On Error GoTo ErrorHandler

    Dim RelationHelper          As IMSRelation.DRelationHelper
    Dim CollectionHelper        As IMSRelation.DCollectionHelper
    Dim oPartClass              As IJDPartClass
    Dim strCompartType          As String
    
    If Object Is Nothing Then
        Set m_vEnumAllowableLoads = Nothing
        Exit Sub
    End If
    
    Set RelationHelper = Object
    
    'Get the PartClass of the Compartment Part
    Set CollectionHelper = RelationHelper.CollectionRelations("IJDPart", PARTCLASS)
    Set oPartClass = CollectionHelper.Item(1)
    strCompartType = oPartClass.Name
    
    'Gives the Load type collection based on the compartment type
    GetLoadTypeBasedOntheRule (strCompartType)
    
    Set CollectionHelper = Nothing
    Set RelationHelper = Nothing
    Set oPartClass = Nothing

Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub


Private Sub GetLoadTypeBasedOntheRule(strCompartType As String)
Const METHOD = "GetLoadTypeBasedOntheRule"
On Error GoTo ErrorHandler

    Dim oSqlCmd             As IJADOMiddleCommand
    Dim strQuery            As String
    Dim strConnectionType       As String
    
    Set oSqlCmd = New JMiddleCommand
    
    Set m_oPOM = GetPOM(CATALOG)
    
    strConnectionType = m_oPOM.ConnectionType
    
    If StrComp(strConnectionType, "ORACLE", vbTextCompare) = 0 Then
        strQuery = "Select ObjectOid from CoreNamedObjects where objectname in"
        strQuery = strQuery & " (Select value from CoreBstrAttribute where iid = HEXTORAW('CFAEBE90070646D8A28CBB347D0647D8') and dispid = 3"
        strQuery = strQuery & " and oid in (Select oid from CoreBstrAttribute where iid = HEXTORAW('CFAEBE90070646D8A28CBB347D0647D8')"
        strQuery = strQuery & " and dispid = 2 and value = '" & strCompartType & "'))"
    Else
        strQuery = "Select ObjectOid from CoreNamedObjects where objectname in"
        strQuery = strQuery & " (Select value from CoreBstrAttribute where iid = 'CFAEBE90-0706-46D8-A28C-BB347D0647D8' and dispid = 3"
        strQuery = strQuery & " and oid in (Select oid from CoreBstrAttribute where iid = 'CFAEBE90-0706-46D8-A28C-BB347D0647D8'"
        strQuery = strQuery & " and dispid = 2 and value = '" & strCompartType & "'))"
    End If
    
  
    'Execute the sql Query string to get Allowable PartClasses of Loads
    With oSqlCmd
        .ActiveConnection = m_oPOM.DatabaseID
        .QueryLanguage = LANGUAGE_SQL
        .CommandText = strQuery
        Set m_vEnumAllowableLoads = .SelectObjects
    End With
    
    Set oSqlCmd = Nothing
    
Exit Sub
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub


Private Function GetPOM(strDbType As String) As IJDPOM
    Const METHOD = "GetPOM"
    On Error GoTo ErrHandler
    
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    
    Exit Function
ErrHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Private Function GetPartClass(oClassItem As IJClassItem) As IJDPartClass
    Const METHOD = "GetPartClass"
    On Error GoTo ErrHandler
    
    Dim RelationHelper          As IMSRelation.DRelationHelper
    Dim CollectionHelper        As IMSRelation.DCollectionHelper
    
    Set RelationHelper = oClassItem

    If RelationHelper Is Nothing Then
        GoTo ErrHandler
    End If

    'get the part class name
    Set CollectionHelper = RelationHelper.CollectionRelations("IJClassItem", PARTCLASS)
    Set GetPartClass = CollectionHelper.Item(1)
    
    Set RelationHelper = Nothing
    Set CollectionHelper = Nothing

    Exit Function
ErrHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
