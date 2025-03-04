VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadCombQI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2005, Intergraph Corporation.  All rights reserved.
''
''  Project     : LoadCombQI
''  File        : CLoadCombQI.cls
''
''  Description : QueryInterpreter for Compartment Load Combination Report
''
''  Author      : Thandur Raghuveer
''
''**************************************************************************************
Option Explicit

Private Const MODULE = "LoadCombQI.CLoadCombQI:"

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

Private m_strVolumeName             As String
Private m_strStructTightness        As String
Private m_strCompartTightness       As String



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
            m_bEvaluateOnly = RHS
    
            Const METHOD = "IJQueryInterpreter_EvaluateOnly"
    
End Property

Public Property Get IJQueryInterpreter_EvaluateOnly() As Boolean
    
    IJQueryInterpreter_EvaluateOnly = m_bEvaluateOnly
    
    Const METHOD = "IJQueryInterpreter_EvaluateOnly"
        
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
   ' m_oErrorServices.ReportUnanticipatedError MODULE, METHOD
    
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
    
    m_strVolumeName = "Name"
    m_strCompartTightness = "CompartTightness"
    m_strStructTightness = "StructTightness"
    
    ' Specify the fields in the recordset.
    With oRecordset.Fields
        .Append m_strVolumeName, adVarWChar, 256, adFldIsNullable
        .Append m_strCompartTightness, adVarWChar, 256, adFldIsNullable
        .Append m_strStructTightness, adVarWChar, 256, adFldIsNullable
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
    Dim oPOm As IJDPOM
    
    Set m_oDBCommand = New ADODB.Command
    
    Set m_oDBCommand.ActiveConnection = m_oCacheController.DBConnection
    
    Set oPOm = GetPOM("Model")
            
    If StrComp(oPOm.ConnectionType, "ORACLE", vbTextCompare) = 0 Then
            strCompartQuery = "Select CAST(jni.oid as varchar2(32)) oid ,  CAST(jlf.oid as varchar2(32)) oid from JNamedItem jni,JnamedItem jlf, XCompartLoadHierarchy Xch Where jni.oid = xch.oiddestination And jlf.oid = xch.oidorigin"
    Else
        strCompartQuery = "Select jni.oid,jlf.oid from JNamedItem jni,JnamedItem jlf, XCompartLoadHierarchy Xch Where jni.oid = xch.oiddestination And jlf.oid = xch.oidorigin"
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
        
    Dim oField                      As Field
    Dim colFields                   As Collection
    Dim lRecCount                   As Long
    Dim lIndex                      As Long
    Dim strOid                      As String
    Dim oIMoniker                   As IMoniker
    Dim oPOm                        As IJDPOM
    Dim oNamedItem                  As IJNamedItem
    Dim oLoadCombObj                As Object
    Dim oLoadFolderObj              As Object
    Dim oLoadCombination            As IJSpatialLoadCombination
    Dim oSpatialLoadColl            As IJElements
    Dim oUnitLoadColl               As IJElements
    Dim lSpatialLoadCount           As Long
    Dim lUnitLoadCount              As Long
    Dim oLoadCombName               As IJNamedItem
    Dim oLoadFolderName             As IJNamedItem
    Dim oSpatialLoadName            As IJNamedItem
    Dim oUnitLoadName               As IJNamedItem
    Dim oSpatialLoadObj             As Object
    Dim oPlate                      As IJPlate
    Dim bLoadCombRequired            As Boolean
    Dim oSpatialLoadOp                As IJSpatialLoadOperation
    Dim oCompartEntity              As IJCompartEntity
    Dim oSpatialCompartPart         As IJDPart
    
    
    
    lRecCount = m_oADORecordset.RecordCount
    
    Set m_oFinalADORecordset = New Recordset
    
    m_oFinalADORecordset.Fields.Append "LoadCombination", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "LoadFolder", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "SpatialLoad", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ConnectedCompartment", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ConnectedCompartmentPartDescription", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "UnitLoad", adVarWChar, 256, adFldIsNullable
    
    m_oFinalADORecordset.Open
    
    Set oPOm = GetPOM("Model")
    
    If oPOm Is Nothing Then Exit Sub
        
    For lIndex = 1 To lRecCount
    
        bLoadCombRequired = True
    
        strOid = m_oADORecordset(0)  'Oid
        
        Set oIMoniker = oPOm.MonikerFromDbIdentifier(strOid)
        
        Set oLoadCombObj = oPOm.GetObject(oIMoniker)
        
        Set oLoadCombination = oLoadCombObj
        
        Set oLoadCombName = oLoadCombination
        
        Set oIMoniker = Nothing
        strOid = vbNullString
        
        strOid = m_oADORecordset(1)  'Oid
        
        Set oIMoniker = oPOm.MonikerFromDbIdentifier(strOid)
        
        Set oLoadFolderObj = oPOm.GetObject(oIMoniker)
        
        Set oLoadFolderName = oLoadFolderObj
                            
        Set oSpatialLoadColl = oLoadCombination.GetSpatialLoadCollection
        
        Set oUnitLoadColl = oLoadCombination.GetUnitLoadCollection
        
        If Not oSpatialLoadColl Is Nothing Then
            lSpatialLoadCount = oSpatialLoadColl.Count
        Else
            lSpatialLoadCount = 0
        End If
        
        If Not oUnitLoadColl Is Nothing Then
            lUnitLoadCount = oUnitLoadColl.Count
        Else
            lUnitLoadCount = 0
        End If
    
        Dim lCurrentLoad As Long
        Dim lMaxCount   As Long
        
        lCurrentLoad = 1
                                    
        If lSpatialLoadCount <> lUnitLoadCount Then
            
            If lSpatialLoadCount > lUnitLoadCount Then
                lMaxCount = lSpatialLoadCount
            Else
                lMaxCount = lUnitLoadCount
            End If
        
            For lCurrentLoad = 1 To lMaxCount
            
                m_oFinalADORecordset.AddNew

                If bLoadCombRequired Then
                    m_oFinalADORecordset.Fields("LoadCombination").Value = oLoadCombName.Name
                    m_oFinalADORecordset.Fields("LoadFolder").Value = oLoadFolderName.Name
                    bLoadCombRequired = False
                End If
                                        
                If lCurrentLoad <= lSpatialLoadCount Then   'Handle the SpatialLoad cases
                
                    Set oSpatialLoadObj = oSpatialLoadColl.Item(lCurrentLoad)
                    Set oSpatialLoadName = oSpatialLoadObj
                    
                    If Not oSpatialLoadName Is Nothing Then
                    
                       m_oFinalADORecordset.Fields("SpatialLoad").Value = oSpatialLoadName.Name
                       
                       Dim objSpatialLoad  As Object
                       Dim oCompartName    As IJNamedItem
                       
                       Set oSpatialLoadOp = oSpatialLoadObj

                       Set objSpatialLoad = oSpatialLoadOp.Compartment

                        If Not objSpatialLoad Is Nothing Then  'No Connected Compartment to the SpatialLoad
    
                            Set oCompartEntity = objSpatialLoad
                            Set oSpatialCompartPart = oCompartEntity.CompartDefinition
                            Set oCompartName = oCompartEntity
    
                            m_oFinalADORecordset.Fields("ConnectedCompartment").Value = oCompartName.Name
                            m_oFinalADORecordset.Fields("ConnectedCompartmentPartDescription").Value = oSpatialCompartPart.PartDescription
    
                            Set oCompartEntity = Nothing
                            Set oSpatialCompartPart = Nothing
                            Set oCompartName = Nothing
                                                    
                        Else  ' No Compartment Gets Connected to Spatial Load
                        
                            m_oFinalADORecordset.Fields("ConnectedCompartment").Value = vbNullString
                            m_oFinalADORecordset.Fields("ConnectedCompartmentPartDescription").Value = vbNullString
                                                    
                        End If
                    
                    Else  ' No Spatial Load Name
                    
                        m_oFinalADORecordset.Fields("SpatialLoad").Value = vbNullString
                        m_oFinalADORecordset.Fields("ConnectedCompartment").Value = vbNullString
                        m_oFinalADORecordset.Fields("ConnectedCompartmentPartDescription").Value = vbNullString
                                                
                    End If
                    
                End If ' End for Spatial Load Cases only
                
                
                If lCurrentLoad <= lUnitLoadCount Then 'Handle the Unit Load Case
                
                    Set oUnitLoadName = oUnitLoadColl.Item(lCurrentLoad)
                    
                    If Not oUnitLoadName Is Nothing Then
                        m_oFinalADORecordset.Fields("UnitLoad").Value = oUnitLoadName.Name
                    Else
                        m_oFinalADORecordset.Fields("UnitLoad").Value = vbNullString
                    End If
                    
                Else  ' If no Unit Loads Exits
                    m_oFinalADORecordset.Fields("UnitLoad").Value = vbNullString
                    
                End If  ' End for Unit Load cases only
            
            Next lCurrentLoad
                            
        
        ElseIf lSpatialLoadCount = lUnitLoadCount And lUnitLoadCount <> 0 Then  ' both are same not Zero
            
             For lCurrentLoad = 1 To lUnitLoadCount

                Set oUnitLoadName = oUnitLoadColl.Item(lCurrentLoad)
                Set oSpatialLoadName = oSpatialLoadColl.Item(lCurrentLoad)

                m_oFinalADORecordset.AddNew

                If bLoadCombRequired Then
                    m_oFinalADORecordset.Fields("LoadCombination").Value = oLoadCombName.Name
                    m_oFinalADORecordset.Fields("LoadFolder").Value = oLoadFolderName.Name
                    bLoadCombRequired = False
                End If

                m_oFinalADORecordset.Fields("SpatialLoad").Value = oSpatialLoadName.Name
                m_oFinalADORecordset.Fields("UnitLoad").Value = oUnitLoadName.Name

                Set oSpatialLoadName = Nothing
                Set oUnitLoadName = Nothing

            Next lCurrentLoad
    
        Else
            
             m_oFinalADORecordset.AddNew

            If bLoadCombRequired Then
                m_oFinalADORecordset.Fields("LoadCombination").Value = oLoadCombName.Name
                m_oFinalADORecordset.Fields("LoadFolder").Value = oLoadFolderName.Name
                bLoadCombRequired = False
            End If
            m_oFinalADORecordset.Fields("SpatialLoad").Value = vbNullString
            m_oFinalADORecordset.Fields("UnitLoad").Value = vbNullString
            
        End If
        
        m_oFinalADORecordset.AddNew
        m_oFinalADORecordset.Fields("LoadCombination").Value = vbNullString
        m_oFinalADORecordset.Fields("LoadFolder").Value = vbNullString
        m_oFinalADORecordset.Fields("SpatialLoad").Value = vbNullString
        m_oFinalADORecordset.Fields("UnitLoad").Value = vbNullString
        
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

