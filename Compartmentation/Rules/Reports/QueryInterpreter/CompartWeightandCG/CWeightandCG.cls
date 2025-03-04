VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWeightandCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2008, Intergraph Corporation.  All rights reservd.
''
''  Project     : CompartWEightandCGQI
''  File        : CWeightandCG.cls
''
''  Description : QueryInterpreter for Compartment weight and CG Report
''
''  Author      : Apparao.N
''**************************************************************************************
Option Explicit

Private Const MODULE = "CompartWeightandCGQI.CWeightandCG:"

Implements IJQueryInterpreter

Private m_bEvaluateOnly                 As Boolean
Private m_oADORecordset                 As ADODB.Recordset
Private m_oFinalADORecordset            As ADODB.Recordset
Private m_oFilterResult                 As Object
Private m_oQParams                      As IJQueryParameters
Private m_oCacheController              As IJCacheController
Private m_oDBCommand                    As ADODB.Command
Private m_oPom                          As IJDPOM


'/////////////////////////////////////////////////////////////////////////////
'Class Constructor/Destructor
'/////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()

    Set m_oADORecordset = Nothing
    Set m_oCacheController = Nothing
    Set m_oPom = Nothing
        
End Sub

'/////////////////////////////////////////////////////////////////////////////
' Implementation of IJQueryInterpreter
'/////////////////////////////////////////////////////////////////////////////
Public Property Set IJQueryInterpreter_CacheController _
            (RHS As SP3DReportsRuntimeInterfaces.IJCacheController)
                                
    Set m_oCacheController = RHS
            
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
        
    'Get the modelDb pointer
    Set m_oPom = GetPOM("Model")
    
    If m_oPom Is Nothing Then Exit Sub


    GetRecordSetData
    ProcessCompartments
    
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
    Dim oRecordset  As ADODB.Recordset
    Dim strOId      As String

    Set oRecordset = New ADODB.Recordset
    
    strOId = "Oid"
    
    ' Specify the fields in the recordset.
    With oRecordset.Fields
        .Append strOId, adVarWChar, 256, adFldIsNullable
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

    Dim strCompartQuery     As String
    Dim lRecCount           As Long
    Dim lIndex              As Long
    Dim oCompartObj        As Object
    Dim strOId              As String
    Dim oIMoniker           As IMoniker
    Dim strConnType         As String
    
    
    'If no object is selected prior to start Run report command
    If m_oFilterResult Is Nothing Then
    
        Set m_oDBCommand = New ADODB.Command
        
        Set m_oDBCommand.ActiveConnection = m_oCacheController.DBConnection
                
        strConnType = m_oPom.ConnectionType
        
        If StrComp(strConnType, "ORACLE", vbTextCompare) = 0 Then
             strCompartQuery = "Select CAST(oid as varchar2(32)) oid from JCompartEntity"
        Else
            'create query to get all the compartments
            strCompartQuery = "Select oid from JCompartEntity"
        End If
        
        m_oDBCommand.CommandText = strCompartQuery
        
        Set m_oADORecordset = m_oDBCommand.Execute
        
        lRecCount = m_oADORecordset.RecordCount
    
        For lIndex = 1 To lRecCount
        
            strOId = m_oADORecordset(0)  'Oid
            Set oIMoniker = m_oPom.MonikerFromDbIdentifier(strOId)
                    
            Set oCompartObj = m_oPom.GetObject(oIMoniker)
            
            If m_oFilterResult Is Nothing Then
                
                Set m_oFilterResult = New Collection
            End If
            
            m_oFilterResult.Add oCompartObj
            
            Set oCompartObj = Nothing
            
            m_oADORecordset.MoveNext
            
        Next lIndex
        
    End If
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Sub ProcessCompartments()
Const METHOD = "ProcessCompartments"
On Error GoTo ErrorHandler

    If m_oFilterResult Is Nothing Then Exit Sub
        
    Dim lRecCount                   As Long
    Dim lIndex                      As Long
    Dim strOId                      As String
    Dim oIMoniker                   As IMoniker
    Dim oCompartObj                 As Object
    Dim oCompartNamedItem           As IJNamedItem
    Dim bCompartRequired            As Boolean
    Dim oFilter                     As Object
    Dim oFactory                    As IJMiddleFiltersFactory
    Dim oObject                     As Object
    Dim oParentFolder               As IJFilterFolderChild
    Dim oSQLFilter                  As IJSqlFilter
    Dim strConnType                 As String

    Set oFactory = New MiddleFiltersFactory
    
    'Create an sql filter which basically filters Plates and Profiles
    Set oParentFolder = oFactory.GetRootFolder(MFMyFiltersRoot, True, Nothing)
    Set oSQLFilter = oFactory.CreateSqlFilter(oParentFolder)
    
    strConnType = m_oPom.ConnectionType
    
    If StrComp(strConnType, "ORACLE", vbTextCompare) = 0 Then
        oSQLFilter.QueryText = "select CAST(oid as varchar2(32)) oid from JPlatePart union select CAST(oid AS varchar2(32)) oid from JProfilePart"
    Else
        oSQLFilter.QueryText = "select oid from dbo.JPlatePart union select oid from dbo.JProfilePart"
    End If
    
    Set oFilter = oSQLFilter
    
    Set m_oFinalADORecordset = New Recordset
    
'            -Compartment
'            -Object Name
'            -ObjectType
'            -NamingCategory
'            -Oid
'            -Condition
'            -Weight
'            -CogX
'            -CogY
'            -CogZ

    
    m_oFinalADORecordset.Fields.Append "Compartment", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Name", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ObjectType", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "NamingCategory", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Oid", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Condition", adVarWChar, 256, adFldIsNullable
     m_oFinalADORecordset.Fields.Append "Volume", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Density", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Weight", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CogX", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CogY", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CogZ", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CompartmentName", adVarWChar, 256, adFldIsNullable
    
    
    m_oFinalADORecordset.Open
    
    lRecCount = m_oFilterResult.Count

    For lIndex = 1 To lRecCount
    
        bCompartRequired = True
        
        Set oCompartObj = m_oFilterResult.Item(lIndex)
        
        If Not TypeOf oCompartObj Is IJCompartEntity Then
            GoTo Continue
        End If
        
        Set oCompartNamedItem = oCompartObj
        
        
        'Fill report with objects which cross compartment
        Call GeneratePartSurfaceAreaReport(oCompartObj, oFilter, lIndex)

       'm_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name

      ' m_oFinalADORecordset.Fields("Name").Value = "Total Area"
        
      
        Set oCompartNamedItem = Nothing
        Set oIMoniker = Nothing
        Set oCompartObj = Nothing


Continue:
        
    Next lIndex
    
    'delete the created sql filter
    Dim oObj As IJDObject
    
    Set oObj = oFilter
    oObj.Remove
    
    Set oObj = oSQLFilter
    oObj.Remove
    
    Set oObj = Nothing
    Set oFilter = Nothing
    Set oSQLFilter = Nothing
    
    
Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Function GetPOM(strDbType As String) As IJDPOM
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    
Exit Function
ErrHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function


'Fills the Object Type(ClassInfo) and Oid of the objects
Private Sub FillObjectInformation(oObject As Object)
On Error GoTo ErrorHandler
Const METHOD = "FillObjectInformation"
    
    Dim oClassInfo      As IJDClassInfo
    Dim oMetaData       As IJDAttributeMetaData
    Dim strClassID      As String
    Dim oMoniker        As IMoniker
    Dim strClassName    As String
    Dim strOId          As String
    
    Set oMetaData = oObject
    
    If oMetaData Is Nothing Then Exit Sub
    
    Set oMoniker = m_oPom.GetObjectMoniker(oObject)
    strOId = m_oPom.DbIdentifierFromMoniker(oMoniker)

    If oMoniker Is Nothing Then Exit Sub
    
    strClassID = m_oPom.ClsidOf(oMoniker)
        
    If strClassID <> vbNullString Then
        Set oClassInfo = oMetaData.ClassInfo(strClassID)
        If oClassInfo Is Nothing Then Exit Sub
        
        strClassName = oClassInfo.UserName
        m_oFinalADORecordset.Fields("ObjectType").Value = strClassName
        m_oFinalADORecordset.Fields("Oid").Value = strOId
         m_oFinalADORecordset.Fields("NamingCategory").Value = oClassInfo.Namespace
        
    End If
    
    Set oClassInfo = Nothing
    Set oMetaData = Nothing
    Set oMoniker = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub


Private Sub GeneratePartSurfaceAreaReport(oCompartment As Object, oFilter As IJSimpleFilter, lIndex As Long)
Const METHOD As String = "GeneratePartSurfaceAreaReport"
On Error GoTo ErrorHandler

    Dim oQueryData                  As IJDQueryData
    Dim oObjectsWhichCross          As IJDObjectCollection
    Dim oObjectsWhichTouchOutside   As IJDObjectCollection
    Dim oObjectsWhichTouchInside    As IJDObjectCollection
    Dim oObjectsInside              As IJDObjectCollection
    Dim oName                       As IJNamedItem
    Dim oObj                        As Object
    Dim eAdjacencyType              As AdjacencyType
    Dim dLength As Double, dArea As Double, dVolume As Double
    Dim bCoatingAreaAvbl            As Boolean
    Dim oCompartNamedItem           As IJNamedItem
    
    Set oCompartNamedItem = oCompartment
    
   
    Set oQueryData = New CompartQueryData

    'get all the objects crossing the Compartment
    With oQueryData
        .InputObject = oCompartment
        .QueryCriteria = Overlapping
        If Not oFilter Is Nothing Then
            .QueryFilter = oFilter
        End If
        .CompartQueryRequired = False
    End With


    'The final Result is following
    Set oObjectsWhichCross = oQueryData.ExecuteQuery

    Set oQueryData = New CompartQueryData

    'get the objects touching inside the compartment
    With oQueryData
        .InputObject = oCompartment
        .QueryCriteria = InsideAndTouching
        If Not oFilter Is Nothing Then
            .QueryFilter = oFilter
        End If
        .CompartQueryRequired = False
    End With

    'The final Result is following
    Set oObjectsWhichTouchInside = oQueryData.ExecuteQuery

    Set oQueryData = New CompartQueryData

    'get the objects inside the compartment
    With oQueryData
        .InputObject = oCompartment
        .QueryCriteria = Inside
        If Not oFilter Is Nothing Then
            .QueryFilter = oFilter
        End If
        .CompartQueryRequired = False
    End With

    'The final Result is following
    Set oObjectsInside = oQueryData.ExecuteQuery
    
    
    Set oQueryData = New CompartQueryData

    'get the objects inside the compartment
    With oQueryData
        .InputObject = oCompartment
        .QueryCriteria = OutsideAndTouching
        If Not oFilter Is Nothing Then
            .QueryFilter = oFilter
        End If
        .CompartQueryRequired = False
    End With

    'The final Result is following
    Set oObjectsWhichTouchOutside = oQueryData.ExecuteQuery


    'fill the overlapping objects info
    If Not oObjectsWhichCross Is Nothing Then
        For Each oObj In oObjectsWhichCross
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then
                 m_oFinalADORecordset.AddNew

                m_oFinalADORecordset.Fields("Name").Value = oName.Name
                
                ReportWeightandCG oCompartment, oObj


                FillObjectInformation oObj

                m_oFinalADORecordset.Fields("Condition").Value = "Overlapping"

                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name & lIndex
                m_oFinalADORecordset.Fields("CompartmentName").Value = oCompartNamedItem.Name
            End If
        Next
    End If
    
         'fill objects touching outside info
    If Not oObjectsWhichTouchOutside Is Nothing Then
        For Each oObj In oObjectsWhichTouchOutside
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then
                    m_oFinalADORecordset.AddNew

                    m_oFinalADORecordset.Fields("Name").Value = oName.Name
    
                    ReportWeightandCG oCompartment, oObj
    
                    FillObjectInformation oObj
    
                    m_oFinalADORecordset.Fields("Condition").Value = "OutsideAndTouching"
    
    
                    m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name & lIndex
                m_oFinalADORecordset.Fields("CompartmentName").Value = oCompartNamedItem.Name


            End If

        Next
    End If


     'fill objects  touching inside info
    If Not oObjectsWhichTouchInside Is Nothing Then
        For Each oObj In oObjectsWhichTouchInside
        
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then
                m_oFinalADORecordset.AddNew

                m_oFinalADORecordset.Fields("Name").Value = oName.Name

                ReportWeightandCG oCompartment, oObj

                m_oFinalADORecordset.Fields("Condition").Value = "InsideAndTouching"
                

                FillObjectInformation oObj
                
                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name & lIndex
                m_oFinalADORecordset.Fields("CompartmentName").Value = oCompartNamedItem.Name

            End If
        Next
    End If

    'fill completely inside objects info
    If Not oObjectsInside Is Nothing Then
        For Each oObj In oObjectsInside
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then
                m_oFinalADORecordset.AddNew

                m_oFinalADORecordset.Fields("Name").Value = oName.Name


                ReportWeightandCG oCompartment, oObj

                FillObjectInformation oObj

                m_oFinalADORecordset.Fields("Condition").Value = "Inside"
                
                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name & lIndex
                m_oFinalADORecordset.Fields("CompartmentName").Value = oCompartNamedItem.Name

            End If
        Next
    End If
    
    Set oQueryData = Nothing
    Set oObjectsWhichCross = Nothing
    Set oObjectsWhichTouchOutside = Nothing
    Set oObjectsWhichTouchInside = Nothing
    Set oObjectsInside = Nothing
    Set oName = Nothing
    Set oObj = Nothing

    Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub
                    


'get the only overlapped area with the compartment
Private Sub ReportWeightandCG(oCompartment As Object, oOverlapObj As Object)
Const METHOD = "ReportWeightandCG"
On Error GoTo ErrorHandler

    Dim oIntersectionGeom               As IJDModelBody
    Dim dAcc As Double, dNewAcc         As Double
    Dim oSpatialService                 As IJDCompartSpatialService
    Dim oGeomService                    As IJDCompartGeomService
    Dim eAdjacency                      As AdjacencyType
    Dim dVolume                         As Double
    Dim dCogX As Double, dCogY As Double, dCogZ As Double
    Dim dDensity As Double
    
    Set oSpatialService = New CompartGeomService
    
    Set oGeomService = New CompartGeomService

    On Error Resume Next
    
 
    Set oIntersectionGeom = oSpatialService.OverlapGeometryAsSolid(oCompartment, oOverlapObj)
    
   
    oGeomService.GetMassProperties oIntersectionGeom, dCogX, dCogY, dCogZ, dVolume
    
    m_oFinalADORecordset.Fields("CogX").Value = dCogX
    m_oFinalADORecordset.Fields("CogY").Value = dCogY
    m_oFinalADORecordset.Fields("CogZ").Value = dCogZ
    
    On Error GoTo ErrorHandler
    
    Set oSpatialService = Nothing
    Set oIntersectionGeom = Nothing
    
    Dim oStructMaterial As IJStructureMaterial
    
    Set oStructMaterial = oOverlapObj
    
    Dim oMaterial As IJDMaterial
    
    Set oMaterial = oStructMaterial.Material
    
    dDensity = oMaterial.Density
    
     m_oFinalADORecordset.Fields("Weight").Value = dDensity * dVolume
    m_oFinalADORecordset.Fields("Volume").Value = dVolume
    m_oFinalADORecordset.Fields("Density").Value = dDensity
       
Exit Sub
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub



