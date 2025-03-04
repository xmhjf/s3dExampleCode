VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCmpTightnessQI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2005, Intergraph Corporation.  All rights reserved.
''
''  Project     : CmpartTightnessQI
''  File        : CCmpartTightnessQI.cls
''
''  Description : QueryInterpreter for Compartment Tightness Report
''
''  Author      : Thandur Raghuveer
''
''**************************************************************************************
Option Explicit

Private Const MODULE = "CompartTightnessQI.CQryInterpreter:"

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
        strCompartQuery = "Select CAST(oid as varchar2(32)) oid ,CompartTightness,StructTightness from JCompartTightness"
      Else
        strCompartQuery = "Select oid,CompartTightness,StructTightness from JCompartTightness"
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
    Dim oCompartObj                 As Object
    Dim oCompartEntity              As IJCompartEntity
    Dim oBoundigFacesColl           As IJDObjectCollection
    Dim lBoundingCount              As Long
    Dim oCompartNamedItem           As IJNamedItem
    Dim oPlateNamedItem             As IJNamedItem
    Dim oPlateObj                   As Object
    Dim oPlate                      As IJPlate
    Dim bCompartRequired            As Boolean
    Dim oCompartPart                As IJDPart
    
    
    lRecCount = m_oADORecordset.RecordCount
    
    Set m_oFinalADORecordset = New Recordset
    
    
    m_oFinalADORecordset.Fields.Append "Name", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "PartDescription", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CompartTightness", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "StructTightness", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Plate", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "PlateTightness", adVarWChar, 256, adFldIsNullable
    
    m_oFinalADORecordset.Open
    
    Set oPOm = GetPOM("Model")
    
    If oPOm Is Nothing Then Exit Sub
        
    For lIndex = 1 To lRecCount
    
        bCompartRequired = True
    
        strOid = m_oADORecordset(0)  'Oid
        
        Set oIMoniker = oPOm.MonikerFromDbIdentifier(strOid)
        
        Set oCompartObj = oPOm.GetObject(oIMoniker)
        
        Set oCompartEntity = oCompartObj
        
        Set oCompartNamedItem = oCompartObj
        
        Set oCompartPart = oCompartEntity.CompartDefinition
            
        Set oBoundigFacesColl = oCompartEntity.GetBoundingFaces
        
        lBoundingCount = oBoundigFacesColl.Count
                    
        If lBoundingCount <> 0 Then
        
            For Each oPlateObj In oBoundigFacesColl
            
                If TypeOf oPlateObj Is IJPlate Then
                    Set oPlate = oPlateObj
                    Set oPlateNamedItem = oPlate
                    m_oFinalADORecordset.AddNew
                    
                    If bCompartRequired Then
                        m_oFinalADORecordset.Fields("Name").Value = oCompartNamedItem.Name
                        m_oFinalADORecordset.Fields("PartDescription").Value = oCompartPart.PartDescription
                        m_oFinalADORecordset.Fields("CompartTightness").Value = GetCompartTightness(m_oADORecordset(1))
                        m_oFinalADORecordset.Fields("StructTightness").Value = GetStructTightness(m_oADORecordset(2))
                       bCompartRequired = False
                    End If

                    m_oFinalADORecordset.Fields("Plate").Value = oPlateNamedItem.Name
                    m_oFinalADORecordset.Fields("PlateTightness").Value = GetStructTightness(oPlate.Tightness)
                    Set oPlate = Nothing
                    Set oPlateNamedItem = Nothing
                End If
            Next oPlateObj
        End If
        
        m_oFinalADORecordset.AddNew
        m_oFinalADORecordset.Fields("Name").Value = vbNullString
        m_oFinalADORecordset.Fields("PartDescription").Value = vbNullString
        m_oFinalADORecordset.Fields("CompartTightness").Value = vbNullString
        m_oFinalADORecordset.Fields("StructTightness").Value = vbNullString
        m_oFinalADORecordset.Fields("Plate").Value = vbNullString
        m_oFinalADORecordset.Fields("PlateTightness").Value = vbNullString
                                
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

Private Function GetCompartTightness(strCompartTightness) As String
On Error GoTo ErrorHandler
Const METHOD = "SetTightnessValue"

    GetCompartTightness = vbNullString

    Select Case strCompartTightness
        Case "1":
            GetCompartTightness = "Undefined"
        Case "5":
            GetCompartTightness = "None"
        Case "10":
            GetCompartTightness = "User Defined"
        Case "20":
            GetCompartTightness = "Unspecified"
        Case "100":
            GetCompartTightness = "Non-Tight"
        Case "110":
            GetCompartTightness = "Fume-Tight"
        Case "120":
            GetCompartTightness = "Air-Tight"
        Case "130":
            GetCompartTightness = "Oil-Tight"
        Case "140":
            GetCompartTightness = "Water-Tight"
        Case "150":
            GetCompartTightness = "Weather-Tight"
    End Select
        
        
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Private Function GetStructTightness(strStructTightness As Variant) As String
On Error GoTo ErrorHandler
Const METHOD = "SetTightnessValue"

    GetStructTightness = vbNullString

    Select Case strStructTightness
        Case "0":
            GetStructTightness = "Unspecified"
        Case "1":
            GetStructTightness = "NonTight"
        Case "2":
            GetStructTightness = "WaterTight"
        Case "3":
            GetStructTightness = "OilTight"
        Case "4":
            GetStructTightness = "AirTight"
        Case 0:
            GetStructTightness = "Unspecified"
        Case 1:
            GetStructTightness = "NonTight"
        Case 2:
            GetStructTightness = "WaterTight"
        Case 3:
            GetStructTightness = "OilTight"
        Case 4:
            GetStructTightness = "AirTight"
    End Select
        
        
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

