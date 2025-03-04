VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompZonesQI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2005, Intergraph Corporation.  All rights reserved.
''
''  Project     : CmpConnZonesQI
''  File        : CCompZonesQI.cls
''
''  Description : QueryInterpreter for Compartments and Connected Zones Report
''
''  Author      : Thandur Raghuveer
''
''  13-Oct-06   Raghuveer  TR-CP·107848  Build Fails due to the unavailability of dependancy dll
''
''**************************************************************************************
Option Explicit

Private Const MODULE = "CmpConnZonesQI.CCompZonesQI:"


Implements IJQueryInterpreter

Private m_bEvaluateOnly             As Boolean
Private m_oADORecordset             As ADODB.Recordset
Private m_oFinalADORecordset        As ADODB.Recordset
Private m_oFilterResult             As Object
Private m_oQParams                  As IJQueryParameters

Private m_oCacheController          As IJCacheController
Private m_oMetaDataQuery            As IJDRepositoryQuery
Private m_oDBCommand                As ADODB.Command

Private m_strOids                   As String

'/////////////////////////////////////////////////////////////////////////////
'Class Constructor/Destructor
'/////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

    Set m_oADORecordset = Nothing
    Set m_oCacheController = Nothing
    Set m_oMetaDataQuery = Nothing
    Set m_oFinalADORecordset = Nothing
    
End Sub

'/////////////////////////////////////////////////////////////////////////////
' Implementation of IJQueryInterpreter
'/////////////////////////////////////////////////////////////////////////////
Public Property Set IJQueryInterpreter_CacheController _
            (RHS As SP3DReportsRuntimeInterfaces.IJCacheController)
Const METHOD = "IJQueryInterpreter_CacheController"
On Error GoTo ErrorHandler
                                
            Set m_oCacheController = RHS
            Set m_oMetaDataQuery = m_oCacheController.MetadataQuery
            
Exit Property
ErrorHandler:
        Err.Raise CompartLogError(Err, MODULE, METHOD)
End Property

Public Property Get IJQueryInterpreter_CacheController() _
            As SP3DReportsRuntimeInterfaces.IJCacheController
            
    
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

    Set m_oFilterResult = RHS
    
End Property

Public Property Get IJQueryInterpreter_FilterResult() As Variant

    Set IJQueryInterpreter_FilterResult = m_oFilterResult
    
End Property

Public Property Set IJQueryInterpreter_Parameters _
            (RHS As SP3DReportsObjectsInterfaces.IJQueryParameters)
            
    Set m_oQParams = RHS
    
End Property

Public Property Get IJQueryInterpreter_Parameters() _
            As SP3DReportsObjectsInterfaces.IJQueryParameters
            
    Set IJQueryInterpreter_Parameters = m_oQParams
    
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
    
    m_strOids = "CompartmentOids"
    
    ' Specify the fields in the recordset.
    With oRecordset.Fields
        .Append m_strOids, adVarWChar, 256, adFldIsNullable
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

    Dim rsTemp As ADODB.Recordset
    Dim rsUValues As ADODB.Recordset
    Dim strCompartQuery As String
    Dim Cnxn1 As ADODB.Connection
    Dim oPom As IJDPOM
    Set m_oDBCommand = New ADODB.Command
    
    Set m_oDBCommand.ActiveConnection = m_oCacheController.DBConnection
            
    Set oPom = GetPOM("Model")
    
    If StrComp(oPom.ConnectionType, "ORACLE", vbTextCompare) = 0 Then
        strCompartQuery = "Select CAST(oid as varchar2(32)) oid from JCompartEntity"
    Else
        strCompartQuery = "Select oid from JCompartEntity"
    End If
    
    m_oDBCommand.CommandText = strCompartQuery
    
    Set m_oADORecordset = m_oDBCommand.Execute
    
    Set Cnxn1 = Nothing
    Set rsTemp = Nothing
    Set rsUValues = Nothing
    Set oPom = Nothing
End Sub
Private Sub ProcessData()
Const METHOD = "ProcessData"
On Error GoTo ErrorHandler

    If m_oADORecordset Is Nothing Then Exit Sub
        
    Dim oField                      As Field
    Dim colFields                   As Collection
    Dim lRecCount                   As Long
    Dim lIndex                      As Long
    Dim oTempRecordset              As Recordset
    Dim strOid                      As String
    Dim oIMoniker                   As IMoniker
    Dim oPom                        As IJDPOM
    Dim oNamedItem                  As IJNamedItem
    Dim oCompartObj                 As Object
    Dim oCompartEntity              As IJCompartEntity
    Dim oResCompartEntity           As IJCompartEntity
    Dim oQueryModel                 As CompartAttributeHelper.Model 'CompartHelper.Model
    Dim oResObjColl                 As IJDObjectCollection
    Dim lReslIndex                  As Long
    Dim oResNamedItem               As IJNamedItem
    Dim oCompartNamedItem           As IJNamedItem
    Dim oResObj                     As Object
    Dim oResPart                    As IJDPart
    Dim oCompartPart                As IJDPart
    Dim strCompartType              As String
    Dim strResCompartType           As String
    Dim bNeedCompartName            As Boolean
    Dim bNeedEmptyRecordSet         As Boolean
    
    lRecCount = m_oADORecordset.RecordCount
    
    Set m_oFinalADORecordset = New Recordset
    
    m_oFinalADORecordset.Fields.Append "Compartments", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CompartmentPartDescription", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Type", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ConnectedZone", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ConnectedType", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ZoneType", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Open
        
    Set oPom = GetPOM("Model")
    
    If oPom Is Nothing Then Exit Sub
    
    Set oQueryModel = New CompartAttributeHelper.Model 'New Model
                
    For lIndex = 1 To lRecCount
    
        bNeedCompartName = True
        
        strOid = m_oADORecordset(0)  'Oid
        
        Set oIMoniker = oPom.MonikerFromDbIdentifier(strOid)
        
        Set oCompartObj = oPom.GetObject(oIMoniker)
        Set oCompartNamedItem = oCompartObj
        Set oCompartEntity = oCompartObj
        Set oCompartPart = oCompartEntity.CompartDefinition
                
        strCompartType = vbNullString
        
        'Get the Part Class Type
        strCompartType = GetClassPropertyValue(oCompartPart, MODULE)
                    
        If strCompartType = "Compartment" Then
                        
            Set oResObjColl = oQueryModel.AdjacentCompartments(oCompartObj)
                    
            For Each oResObj In oResObjColl
                      
                Set oResNamedItem = oResObj
                Set oResCompartEntity = oResObj
                Set oResPart = oResCompartEntity.CompartDefinition
                
                strResCompartType = vbNullString
                
                strResCompartType = GetClassPropertyValue(oResPart, MODULE)
                
                If strResCompartType = "ShipZone" Then
                                                        
                    m_oFinalADORecordset.AddNew
                    
                    If bNeedCompartName Then
                        m_oFinalADORecordset.Fields("Compartments").Value = oCompartNamedItem.Name
                        m_oFinalADORecordset.Fields("CompartmentPartDescription").Value = oCompartPart.PartDescription
                        m_oFinalADORecordset.Fields("Type").Value = strCompartType
                        m_oFinalADORecordset.Fields("ConnectedType").Value = strResCompartType
                        bNeedCompartName = False
                    End If
                                    
                    m_oFinalADORecordset.Fields("ConnectedZone").Value = oResNamedItem.Name
                    m_oFinalADORecordset.Fields("ZoneType").Value = GetPartClassName(oResPart)
                    
                    bNeedEmptyRecordSet = True
                    
                End If  'End of ShipZone Type Check
                    
            Next oResObj
            
            If bNeedEmptyRecordSet Then
            
                m_oFinalADORecordset.AddNew
                m_oFinalADORecordset.Fields("Compartments").Value = vbNullString
                m_oFinalADORecordset.Fields("CompartmentPartDescription").Value = vbNullString
                m_oFinalADORecordset.Fields("Type").Value = vbNullString
                m_oFinalADORecordset.Fields("ConnectedZone").Value = vbNullString
                m_oFinalADORecordset.Fields("ConnectedType").Value = vbNullString
                m_oFinalADORecordset.Fields("ZoneType").Value = vbNullString
                bNeedEmptyRecordSet = False
            End If
            
            
         End If  ' End of Compartment Check
                    
            Set oCompartObj = Nothing
            Set oIMoniker = Nothing
          
            m_oADORecordset.MoveNext
                        
        Next lIndex
        
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

Private Function GetPartClassName(ByVal oSpacePart As IJDPart) As String
Const METHOD = "GetPartClassName"
On Error GoTo ErrorHandler
    
    Dim oPartClass                  As IJDPartClass
    Dim oPartRelationHelper         As IMSRelation.DRelationHelper
    Dim oPartCollection             As IMSRelation.DCollectionHelper
    Dim strClass                    As String
    
    GetPartClassName = vbNullString
    
    If oSpacePart Is Nothing Then Exit Function
        
        If oSpacePart.PartDescription <> vbNullString Then
        
            'Get the Relation Helper
            Set oPartRelationHelper = oSpacePart
            
            'get the partclass
            Set oPartCollection = oPartRelationHelper.CollectionRelations(IID_IJDPart, PARTCLASS)
            
            Set oPartClass = oPartCollection.Item(1)
            
            GetPartClassName = oPartClass.Name
            
        End If
            
    Set oPartClass = Nothing
    Set oPartCollection = Nothing
    Set oPartRelationHelper = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
