VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjInside"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2005, Intergraph Corporation.  All rights reservd.
''
''  Project     : ObjectsInside
''  File        : CObjectsInside.cls
''
''  Description : QueryInterpreter for Objects Inside the Compartments Report
''
''  Author      : Thandur Raghuveer
''
''**************************************************************************************
Option Explicit

Private Const MODULE = "ObjectsInside.CObjInside:"

Implements IJQueryInterpreter

Private m_bEvaluateOnly             As Boolean
Private m_oADORecordset             As ADODB.Recordset
Private m_oFinalADORecordset        As ADODB.Recordset
Private m_oFilterResult             As Object
Private m_oParameters               As IJQueriesParameters
Private m_oQParams                  As IJQueryParameters

Private m_oCacheController          As IJCacheController
Private m_oMetaDataQuery            As IJDRepositoryQuery
Private m_oDBCommand                As ADODB.Command

Private m_rsCompartData             As ADODB.Recordset
Private m_oElements                 As IJElements

Private m_oDatabaseHelper           As ReportsDatabaseHelperInterfaceLib.IDatabaseHelper

Private m_strOid             As String



'/////////////////////////////////////////////////////////////////////////////
'Class Constructor/Destructor
'/////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()

    Set m_oADORecordset = Nothing
     Set m_oCacheController = Nothing
    Set m_oMetaDataQuery = Nothing
        
    'Set m_oErrorServices = Nothing
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
' Implementation of IJQueryInterpreter
'/////////////////////////////////////////////////////////////////////////////
Public Property Set IJQueryInterpreter_CacheController _
            (RHS As SP3DReportsRuntimeInterfaces.IJCacheController)
                                
            Set m_oCacheController = RHS
            Set m_oMetaDataQuery = m_oCacheController.MetadataQuery
            
End Property

Public Property Get IJQueryInterpreter_CacheController() _
            As SP3DReportsRuntimeInterfaces.IJCacheController
            
            Const METHOD = "IJQueryInterpreter_CacheController"
            
    
End Property

Public Property Let IJQueryInterpreter_EvaluateOnly(ByVal RHS As Boolean)
Const METHOD = "IJQueryInterpreter_EvaluateOnly"
On Error GoTo ErrorHandler

            m_bEvaluateOnly = RHS
                
Exit Property
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Property Get IJQueryInterpreter_EvaluateOnly() As Boolean
Const METHOD = "IJQueryInterpreter_EvaluateOnly"
On Error GoTo ErrorHandler
    
    IJQueryInterpreter_EvaluateOnly = m_bEvaluateOnly
    
Exit Property
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Sub IJQueryInterpreter_Execute(Command As String, Argument As String)
Const METHOD = "IJQueryInterpreter_Execute"
On Error GoTo ErrorHandler
    
    ' Create the Recordset that is to be returned.  It will be open so
    ' that records may be added.
    If Not m_oADORecordset Is Nothing Then
        Set m_oADORecordset = Nothing
    End If
        
    Set m_oADORecordset = CreateOpenRecordset
        
    GetRecordSetData
        
    ProcessData
    
    ' In the design mode only the empty recordset will be returned.  The
    ' available fields are what is required in that case rather than
    ' actual data.
    
    If m_bEvaluateOnly = True Then
        Exit Sub
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Public Property Set IJQueryInterpreter_FilterResult(ByVal RHS As Object)
Const METHOD = "IJQueryInterpreter_Parameters"
On Error GoTo ErrorHandler


    Set m_oFilterResult = RHS
    
    
Exit Property
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Property Get IJQueryInterpreter_FilterResult() As Variant
Const METHOD = "IJQueryInterpreter_Parameters"
On Error GoTo ErrorHandler


    Set IJQueryInterpreter_FilterResult = m_oFilterResult
    
    
Exit Property
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Property Set IJQueryInterpreter_Parameters _
            (RHS As SP3DReportsObjectsInterfaces.IJQueryParameters)
Const METHOD = "IJQueryInterpreter_Parameters"
On Error GoTo ErrorHandler

            
    Set m_oQParams = RHS
    
Exit Property
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Property Get IJQueryInterpreter_Parameters() _
            As SP3DReportsObjectsInterfaces.IJQueryParameters
Const METHOD = "IJQueryInterpreter_Parameters"
On Error GoTo ErrorHandler
            
    Set IJQueryInterpreter_Parameters = m_oQParams

Exit Property
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Property Get IJQueryInterpreter_ParticipatingObjects() _
            As SP3DReportsRuntimeInterfaces.IJElements
            
End Property

Public Property Set IJQueryInterpreter_Query _
            (RHS As SP3DReportsObjectsInterfaces.IJQuery)
            
End Property

Public Property Get IJQueryInterpreter_Query() _
            As SP3DReportsObjectsInterfaces.IJQuery
    
End Property

Public Property Set IJQueryInterpreter_QueryFilter _
            (RHS As SP3DReportsObjectsInterfaces.IJQueryFilter)
    
End Property

Public Property Get IJQueryInterpreter_QueryFilter() _
            As SP3DReportsObjectsInterfaces.IJQueryFilter
    
End Property

Public Property Get IJQueryInterpreter_QueryResult() As Variant
    
    Set IJQueryInterpreter_QueryResult = m_oFinalADORecordset
    
End Property


'******************************************************************************
' Method        : CreateOpenRecordset
' Description   : Create Recordset and populate the property names
'
'******************************************************************************
Private Function CreateOpenRecordset() As ADODB.Recordset
Const METHOD = "CreateOpenRecordset"
On Error GoTo ErrorHandler
    
    
    ' Create the recordset
    Dim oRecordset As ADODB.Recordset
    Set oRecordset = New ADODB.Recordset
    
    m_strOid = "Oid"
    
    ' Specify the fields in the recordset.
    With oRecordset.Fields
        .Append m_strOid, adVarWChar, 256, adFldIsNullable
    End With
    
    ' Open the recordset to enable addition of records.
    oRecordset.Open , , , adLockOptimistic
    
    Set m_oADORecordset = oRecordset
    
CleanUp:
    Set oRecordset = Nothing
    Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Function

Private Sub GetRecordSetData()
Const METHOD = "GetRecordSetData"
On Error GoTo ErrorHandler

    Dim rsTemp As ADODB.Recordset
    Dim rsUValues As ADODB.Recordset
    Dim strCompartQuery As String
    Dim Cnxn1 As ADODB.Connection
    Dim oPOm As IJDPOM
    
    Set m_oDBCommand = New ADODB.Command
    
    Set m_oDBCommand.ActiveConnection = m_oCacheController.DBConnection
            
    Set oPOm = GetPOM("Model")
    
    If StrComp(oPOm.ConnectionType, "ORACLE", vbTextCompare) = 0 Then
        strCompartQuery = "Select CAST(oid as varchar2(32)) oid from JCompartEntity"
    Else
        strCompartQuery = "Select oid from JCompartEntity"
    End If

    
    
    m_oDBCommand.CommandText = strCompartQuery
    
    Set m_oADORecordset = m_oDBCommand.Execute
    
    Set Cnxn1 = Nothing
    Set rsTemp = Nothing
    Set rsUValues = Nothing
    Set oPOm = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Sub ProcessData()
Const METHOD = "ProcessData"
On Error GoTo ErrorHandler

    If m_oADORecordset Is Nothing Then Exit Sub
        
    Dim lRecCount                   As Long
    Dim lIndex                      As Long
    Dim strOid                      As String
    Dim oIMoniker                   As IMoniker
    Dim oPOm                        As IJDPOM
    Dim oCompartObj                 As Object
    Dim oCompartEntity              As IJCompartEntity
    Dim oInsideObjColl              As IJDObjectCollection
    Dim oCompartNamedItem           As IJNamedItem
    Dim oInsideObjNamedItem         As IJNamedItem
    Dim oInsideObj                  As Object
    
    Dim bCompartRequired            As Boolean
    Dim oCompartPart                As IJDPart
    Dim oQueryModel                 As CompartAttributeHelper.Model  'CompartHelper.Model
            
    If m_oFilterResult Is Nothing Then
        lRecCount = m_oADORecordset.RecordCount
    Else
        lRecCount = m_oFilterResult.Count
    End If
    
    Set m_oFinalADORecordset = New Recordset

    m_oFinalADORecordset.Fields.Append "ObjectOid", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Name", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "InsideObject", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "PartDesc", adVarWChar, 256, adFldIsNullable
    
    m_oFinalADORecordset.Open
    
    Set oPOm = GetPOM("Model")
    
    If oPOm Is Nothing Then Exit Sub
        
    For lIndex = 1 To lRecCount
    
        bCompartRequired = True
    
        strOid = m_oADORecordset(0)  'Oid
        
        Set oIMoniker = oPOm.MonikerFromDbIdentifier(strOid)
            
            If m_oFilterResult Is Nothing Then
                Set oCompartObj = oPOm.GetObject(oIMoniker)
            Else
                Set oCompartObj = m_oFilterResult.Item(lIndex)
            End If
        
        Set oCompartEntity = oCompartObj
        
        Set oCompartPart = oCompartEntity.CompartDefinition
        
        Set oCompartNamedItem = oCompartObj
        
        Set oQueryModel = New CompartAttributeHelper.Model 'New Model
        
        Set oInsideObjColl = oQueryModel.ObjectsWithInCompartment(oCompartObj)
        
        For Each oInsideObj In oInsideObjColl
        
        
            If (Not oPOm.SupportsInterface(oPOm.GetObjectMoniker(oInsideObj), "IJShpStrSystem")) And (Not oPOm.SupportsInterface(oPOm.GetObjectMoniker(oInsideObj), "ISPSMemberSystem")) Then
                                        
                Set oInsideObjNamedItem = oInsideObj
                
                m_oFinalADORecordset.AddNew
                
                m_oFinalADORecordset.Fields("ObjectOid").Value = GetOID(oInsideObj)
                
                    If bCompartRequired Then
                        m_oFinalADORecordset.Fields("Name").Value = oCompartNamedItem.Name
                        bCompartRequired = False
                    End If
                
                m_oFinalADORecordset.Fields("PartDesc").Value = oCompartPart.PartDescription
                
                    If Not oInsideObjNamedItem Is Nothing Then
                        m_oFinalADORecordset.Fields("InsideObject").Value = oInsideObjNamedItem.Name
                    End If
                                            
                Set oInsideObjNamedItem = Nothing
                
            End If
            
        Next oInsideObj
        
        Set oCompartNamedItem = Nothing
        Set oCompartPart = Nothing
        
        If oInsideObjColl.Count <> 0 Then
            m_oFinalADORecordset.AddNew
            m_oFinalADORecordset.Fields("ObjectOid").Value = vbNullString
            m_oFinalADORecordset.Fields("Name").Value = vbNullString
            m_oFinalADORecordset.Fields("PartDesc").Value = vbNullString
            m_oFinalADORecordset.Fields("InsideObject").Value = vbNullString
        End If
        
        Set oInsideObjColl = Nothing
        m_oADORecordset.MoveNext
        
    Next lIndex
    
    Set oCompartObj = Nothing
    Set oCompartEntity = Nothing
    Set oInsideObjColl = Nothing
    Set oCompartNamedItem = Nothing
    Set oInsideObjNamedItem = Nothing
    Set oInsideObj = Nothing
    
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

Private Function GetOID(oObject As Object) As String
On Error GoTo ErrorHandler
Const METHOD = "GetOID"
    
    Dim oMoniker        As IMoniker
    Dim oPOm As IJDPOM

    Set oPOm = GetPOM("Model")
    
    Set oMoniker = oPOm.GetObjectMoniker(oObject)
    GetOID = oPOm.DbIdentifierFromMoniker(oMoniker)
    
    Set oMoniker = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
