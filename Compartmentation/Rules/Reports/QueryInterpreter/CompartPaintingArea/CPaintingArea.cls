VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaintingArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**************************************************************************************
''  Copyright (C) 2008, Intergraph Corporation.  All rights reservd.
''
''  Project     : CompartPaintingAReaQI
''  File        : CPaintingArea.cls
''
''  Description : QueryInterpreter for Compartment Painting area Report
''
''  Author      : Apparao.N
''15th Jan 2009 Apparao TR-CP-153986  Paint Area Report includes object out side the compartment
' 11th Aug 2009 Apparao TR-CP-168940  When a volume is created by only 6 plates, painting area report does not show a
' 25th Mar 2010 Apparao DM-180418   Plate painting area duplicated/incorrectly reported in certain cases
' 17th April 2010 Apparao TR-CP-177187  Surface Area and Covering Area calculation is wrong in Painting Area report
''**************************************************************************************
Option Explicit

Private Const MODULE = "CompartPaintingAreaQI.CPaintingArea:"
Private Const m_dTolerance As Double = 0.001

Implements IJQueryInterpreter

Private m_bEvaluateOnly                 As Boolean
Private m_oADORecordset                 As ADODB.Recordset
Private m_oFinalADORecordset            As ADODB.Recordset
Private m_oFilterResult                 As Object
Private m_oQParams                      As IJQueryParameters
Private m_oCacheController              As IJCacheController
Private m_oDBCommand                    As ADODB.Command
Private m_oPom                          As IJDPOM
Private m_oBoundaryColl                 As IJDObjectCollection
Private m_dTotalSurfaceArea             As Double
Private m_dTotalCoveringArea            As Double
Private m_lCompartCoatingType           As String
Private m_lCompartCoatingColor          As String
Private m_lCompartCoatingLevel          As Long
Private m_lSurfaceAreaFactor            As Long

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

Public Sub IJQueryInterpreter_Execute(Command As String, argument As String)
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


    GetRecordSetData (argument)
    ProcessCompartments (argument)
    
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

Private Sub GetRecordSetData(argument As String)
Const METHOD = "GetRecordSetData"
On Error GoTo ErrorHandler

    Dim strCompartQuery     As String
    Dim lRecCount           As Long
    Dim lIndex              As Long
    Dim oCompartObj        As Object
    Dim strOId              As String
    Dim oIMoniker           As IMoniker
    Dim strconnType         As String
       
    
    If Not m_oFilterResult Is Nothing Then Exit Sub
    
    'If no object is selected prior to start Run report command
   
    Set m_oDBCommand = New ADODB.Command
    
    Set m_oDBCommand.ActiveConnection = m_oCacheController.DBConnection
    
    strconnType = m_oPom.ConnectionType

    If argument <> vbNullString Then
        If StrComp(strconnType, "ORACLE", vbTextCompare) = 0 Then
            'create query to get all the spaces
            strCompartQuery = "Select CAST(oid as varchar2(32)) oid from JSpaceEntity"
        Else
            'create query to get all the spaces
            strCompartQuery = "Select oid from JSpaceEntity"
        End If
    Else
        If StrComp(strconnType, "ORACLE", vbTextCompare) = 0 Then
            'create query to get all the compartments
            strCompartQuery = "Select CAST(oid as varchar2(32)) oid from JCompartEntity"
        Else
            'create query to get all the compartments
            strCompartQuery = "Select oid from JCompartEntity"
        End If
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
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Sub ProcessCompartments(argument As String)
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
    Dim oCompartByBound             As IJCompartByBound
    Dim oCompartCreation            As IJCompartCreation
    Dim oObject                     As Object
    Dim oParentFolder               As IJFilterFolderChild
    Dim oSQLFilter                  As IJSqlFilter
    Dim strconnType                 As String

    Set oFactory = New MiddleFiltersFactory
    
    strconnType = m_oPom.ConnectionType
    
    'Create an sql filter which basically filters Plates and Profiles
    Set oParentFolder = oFactory.GetRootFolder(MFMyFiltersRoot, True, Nothing)
    Set oSQLFilter = oFactory.CreateSqlFilter(oParentFolder)
    If StrComp(strconnType, "ORACLE", vbTextCompare) = 0 Then
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
'            -CoatingType
'            -Coating Color
'            -CoatingLevel
'            -SurfaceArea
'            -CoveringArea
    
    m_oFinalADORecordset.Fields.Append "Compartment", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Name", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "ObjectType", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "NamingCategory", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Oid", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "Condition", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CoatingType", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CoatingColor", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CoatingLevel", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "SurfaceArea", adVarWChar, 256, adFldIsNullable
    m_oFinalADORecordset.Fields.Append "CoveringArea", adVarWChar, 256, adFldIsNullable
    
    m_oFinalADORecordset.Open
    
    lRecCount = m_oFilterResult.Count
        
    For lIndex = 1 To lRecCount
    
        bCompartRequired = True
        
        m_dTotalSurfaceArea = 0
        m_dTotalCoveringArea = 0
        m_lSurfaceAreaFactor = 1
        m_lCompartCoatingLevel = 0
        
        Set oCompartObj = m_oFilterResult.Item(lIndex)
        ' We are using same QI project for PaintingAreaReports and VolumeInsideObjects
        'For "VolumeInsideObjects" reports argument is "Volume"
        'For PaintingAreaReports argument is null string
        If argument <> vbNullString Then
        
            If Not TypeOf oCompartObj Is IJSpaceEntity Then
        
             GoTo Continue
            End If
        Else
            If Not TypeOf oCompartObj Is IJCompartEntity Then

             GoTo Continue
            End If
        End If
        
        Set oCompartNamedItem = oCompartObj
        
        'get the boundaries of the compartement
        On Error Resume Next
        Set oCompartCreation = oCompartObj
        
        'get the active entity
        If Not oCompartCreation Is Nothing Then
        Set oCompartByBound = oCompartCreation.GenerationAE
        End If
        Err.Clear
        
        On Error GoTo ErrorHandler
        m_oFinalADORecordset.AddNew
        
        

        'Gets CompartCoating area attributes
        GetCompartCoatingAreaAttributes oCompartObj
        
        Set m_oBoundaryColl = New JObjectCollection
        
        If Not oCompartByBound Is Nothing Then
            Dim oFaceColl       As New Collection 'IJElements
            Dim oFace           As Object
            Dim ii              As Long

           
            'get the compartment boundaries
            Set oFaceColl = GetCompartBoundaries(oCompartObj)
            
            For ii = 1 To oFaceColl.Count

                Set oFace = oFaceColl.Item(ii)
                
                
                If TypeOf oFace Is IJPlateSystem Then
                    Dim oSystem             As IJSystem
                    Dim oChildrenPS         As IJDTargetObjectCol
                    Dim index               As Integer
                    Dim oPlatePart          As IJPlatePart
                    Dim oChildPS            As IJSystem
                    Dim oChildParts         As IJDTargetObjectCol
                    
                    
                    Set oSystem = oFace
    
                        Set oChildrenPS = oSystem.GetChildren
                        
                        For index = 1 To oChildrenPS.Count
                            If TypeOf oChildrenPS.Item(index) Is IJPlateSystem Then
                            
                                On Error Resume Next
                                Set oChildPS = oChildrenPS.Item(index)
    
                                Set oChildParts = oChildPS.GetChildren
                               
                                
                                If oChildParts.Count = 1 Then
                                    Set oPlatePart = oChildParts.Item(1)
                                End If
                                
                                If Not oPlatePart Is Nothing Then
                                    
                                    m_oBoundaryColl.Add oPlatePart
                                End If
                                
                                Err.Clear
                                On Error GoTo ErrorHandler
                                
                            End If
                            
                        Next index
                    
                End If

            Next ii
        End If

        'Fill report with objects which cross compartment
        Call GeneratePartSurfaceAreaReport(oCompartObj, oFilter)

        m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name

        m_oFinalADORecordset.Fields("Name").Value = "Total Area"
        m_oFinalADORecordset.Fields("SurfaceArea").Value = m_dTotalSurfaceArea
        m_oFinalADORecordset.Fields("CoveringArea").Value = m_dTotalCoveringArea
        
      
        Set oCompartNamedItem = Nothing
        Set oCompartByBound = Nothing
        Set oIMoniker = Nothing
        Set oCompartObj = Nothing
        Set oCompartCreation = Nothing
        Set oSystem = Nothing
        Set oChildrenPS = Nothing
        Set oPlatePart = Nothing
        Set oChildPS = Nothing
        Set oChildParts = Nothing
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
'Fills the report with the Compartment Coating Type and Coating Level attributes
Private Sub FillCompartCoatingAttributes()
On Error GoTo ErrorHandler
Const METHOD = "FillCompartCoatingAttributes"
    
    m_oFinalADORecordset.Fields("CoatingType").Value = m_lCompartCoatingType
    
    If m_lCompartCoatingLevel <> 0 Then
        m_oFinalADORecordset.Fields("CoatingLevel").Value = m_lCompartCoatingLevel
    End If
    
    m_oFinalADORecordset.Fields("CoatingColor").Value = m_lCompartCoatingColor
        
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

'Fills the report with the Compartment Coating Type and Coating Level attributes
Private Sub GetCompartCoatingAreaAttributes(oCompart As Object)
On Error GoTo ErrorHandler
Const METHOD = "GetCompartCoatingAreaAttributes"
    
    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim oCodelist                   As IJDCodeListMetaData
                    
    Set oCodelist = m_oPom

    Set oAttributeMetadata = oCompart
    Set oAttrHelper = oCompart
    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.IID("IJCompartCoating"))
    
    If Not oInterfaceInfo Is Nothing Then
        On Error Resume Next
        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
        Err.Clear
        On Error GoTo ErrorHandler
        If Not oAttributesCollection Is Nothing Then
        
            For Each oAttribute In oAttributesCollection
                If oAttribute.AttributeInfo.Name = "CoatingType" Then
                    m_lCompartCoatingType = oCodelist.LongStringValue(oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                ElseIf oAttribute.AttributeInfo.Name = "CoatingLevel" Then
                    If oAttribute.Value <> 0 Then
                        m_lCompartCoatingLevel = oAttribute.Value
                        m_lSurfaceAreaFactor = m_lCompartCoatingLevel
                    End If
                ElseIf oAttribute.AttributeInfo.Name = "CoatingColor" Then
                    m_lCompartCoatingColor = oCodelist.LongStringValue(oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                End If
            Next
        End If
    End If
    
    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing
    Set oCodelist = Nothing
        
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub
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

'Get the compartment boundaries
Private Function GetCompartBoundaries(oCompart As IJCompartCreation) As Collection
Const METHOD = "GetCompartBoundaries"
On Error GoTo ErrorHandler

    Dim i                           As Integer
    Dim vDummy()                    As Variant
    Dim oCompartFace                As IJCompartByBoundFace
    Dim oCompartFaceCollection      As New Collection 'IJElements
    Dim oFaceColl                   As New Collection 'IJElements
    
    'get the compart faces

    oCompart.GetInputs oCompartFaceCollection, vDummy

    'get compart boundaries
    For i = 1 To oCompartFaceCollection.Count
        Set oCompartFace = oCompartFaceCollection.Item(i)
        oFaceColl.Add oCompartFace.Face
    Next i

    Set GetCompartBoundaries = oFaceColl
        
    Set oCompartFaceCollection = Nothing
    Set oCompartFace = Nothing
    Set oFaceColl = Nothing
    
   Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Private Sub GeneratePartSurfaceAreaReport(oCompartment As Object, oFilter As IJSimpleFilter)
Const METHOD As String = "GeneratePartSurfaceAreaReport"
On Error GoTo ErrorHandler

    Dim oQueryData                  As IJDQueryData
    Dim oObjectsWhichCross          As IJDObjectCollection
    Dim oObjectsWhichTouchOutside   As IJDObjectCollection
    Dim oObjectsWhichTouchInside    As IJDObjectCollection
    Dim oObjectsInside              As IJDObjectCollection
    Dim oCompartService             As CompartAttributeHelper.Compartment ' Compartment
    Dim oName                       As IJNamedItem
    Dim oObj                        As Object
    Dim eAdjacencyType              As AdjacencyType
    Dim dLength As Double, dArea As Double, dVolume As Double
    Dim bCoatingAreaAvbl            As Boolean
    Dim oCompartNamedItem           As IJNamedItem
    
    Set oCompartNamedItem = oCompartment
    Set oCompartService = New CompartAttributeHelper.Compartment
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

    ' Remove the boundaries from the other collections, because
    ' boundary information is already added to the report

'    If m_oBoundaryColl.Count <> 0 Then
'        oObjectsWhichCross.SetSubtract m_oBoundaryColl
''        oObjectsWhichTouchOutside.SetSubtract m_oBoundaryColl
'        oObjectsWhichTouchInside.SetSubtract m_oBoundaryColl
'        oObjectsInside.SetSubtract m_oBoundaryColl
'    End If

    'fill the overlapping objects info
    If Not oObjectsWhichCross Is Nothing Then
        For Each oObj In oObjectsWhichCross
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then
                m_oFinalADORecordset.Fields("Name").Value = oName.Name
'                oCompartService.CompartIntersectionArea oCompartment, oObj, eAdjacencyType, dLength, dArea, dVolume
                OverlapPaintArea oCompartment, oObj, dArea

                m_oFinalADORecordset.Fields("SurfaceArea").Value = dArea
                m_dTotalSurfaceArea = m_dTotalSurfaceArea + dArea

                FillPartCoatingAreaInfo oObj, bCoatingAreaAvbl

                If bCoatingAreaAvbl = False Then
                    m_oFinalADORecordset.Fields("CoveringArea").Value = m_lSurfaceAreaFactor * dArea
                    m_dTotalCoveringArea = m_dTotalCoveringArea + m_lSurfaceAreaFactor * dArea
                End If
                
                FillObjectInformation oObj
                
                If m_oBoundaryColl.Contains(oObj) Then
                    m_oFinalADORecordset.Fields("Condition").Value = "Boundary-Overlapping"
                Else
                    m_oFinalADORecordset.Fields("Condition").Value = "Overlapping"
                End If
                
                FillCompartCoatingAttributes
                
                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name
                m_oFinalADORecordset.AddNew
            End If
        Next
    End If
    
         'fill objects touching outside info
    If Not oObjectsWhichTouchOutside Is Nothing Then
        For Each oObj In oObjectsWhichTouchOutside
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then
                m_oFinalADORecordset.Fields("Name").Value = oName.Name

                oCompartService.CompartIntersectionArea oCompartment, oObj, eAdjacencyType, dLength, dArea, dVolume

                m_oFinalADORecordset.Fields("SurfaceArea").Value = dArea
                m_dTotalSurfaceArea = m_dTotalSurfaceArea + dArea

                FillPartCoatingAreaInfo oObj, bCoatingAreaAvbl

                If bCoatingAreaAvbl = False Then
                    m_oFinalADORecordset.Fields("CoveringArea").Value = m_lSurfaceAreaFactor * dArea
                    m_dTotalCoveringArea = m_dTotalCoveringArea + m_lSurfaceAreaFactor * dArea
                End If

                FillObjectInformation oObj

                If m_oBoundaryColl.Contains(oObj) Then
                    m_oFinalADORecordset.Fields("Condition").Value = "Boundary-OutsideAndTouching"
                Else
                    m_oFinalADORecordset.Fields("Condition").Value = "OutsideAndTouching"
                End If

                FillCompartCoatingAttributes

                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name
                m_oFinalADORecordset.AddNew
            End If
        Next
    End If


     'fill objects  touching inside info
    If Not oObjectsWhichTouchInside Is Nothing Then
        For Each oObj In oObjectsWhichTouchInside
        
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then

                m_oFinalADORecordset.Fields("Name").Value = oName.Name

'                oCompartService.CompartIntersectionArea oCompartment, oObj, eAdjacencyType, dLength, dArea, dVolume
                OverlapPaintArea oCompartment, oObj, dArea

                If m_oBoundaryColl.Contains(oObj) Then
                    m_oFinalADORecordset.Fields("Condition").Value = "Boundary-InsideAndTouching"
                    'dArea = dArea / 2
                Else
                    m_oFinalADORecordset.Fields("Condition").Value = "InsideAndTouching"
                End If
                
                m_oFinalADORecordset.Fields("SurfaceArea").Value = dArea
                m_dTotalSurfaceArea = m_dTotalSurfaceArea + dArea

                FillPartCoatingAreaInfo oObj, bCoatingAreaAvbl

                If bCoatingAreaAvbl = False Then
                    m_oFinalADORecordset.Fields("CoveringArea").Value = m_lSurfaceAreaFactor * dArea
                    m_dTotalCoveringArea = m_dTotalCoveringArea + m_lSurfaceAreaFactor * dArea
                End If

                FillObjectInformation oObj
                
                FillCompartCoatingAttributes
                
                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name
                m_oFinalADORecordset.AddNew
            End If
        Next
    End If

    'fill completely inside objects info
    If Not oObjectsInside Is Nothing Then
        For Each oObj In oObjectsInside
            Set oName = oObj
            If (Not TypeOf oObj Is ISPGFrameEntity) And (Not TypeOf oObj Is IJPlateSystem) Then

                m_oFinalADORecordset.Fields("Name").Value = oName.Name

                'for Plate and profiles, the surafce area is an attribute and hence get area using that attribute
                If TypeOf oObj Is IJPlatePart Then
                    Dim Oplate As IJPlatePart
                    Set Oplate = oObj
                    dArea = Oplate.Area
                    Set Oplate = Nothing

                ElseIf TypeOf oObj Is IJProfilePart Then
                    Dim oProfile As IJProfilePart

                    Set oProfile = oObj
                    dArea = oProfile.Area
                    Set oProfile = Nothing
                Else
'                    oCompartService.CompartIntersectionArea oCompartment, oObj, eAdjacencyType, dLength, dArea, dVolume
                    OverlapPaintArea oCompartment, oObj, dArea
                End If

                m_oFinalADORecordset.Fields("SurfaceArea").Value = dArea
                m_dTotalSurfaceArea = m_dTotalSurfaceArea + dArea

                FillPartCoatingAreaInfo oObj, bCoatingAreaAvbl

                If bCoatingAreaAvbl = False Then
                    m_oFinalADORecordset.Fields("CoveringArea").Value = m_lSurfaceAreaFactor * dArea
                    m_dTotalCoveringArea = m_dTotalCoveringArea + m_lSurfaceAreaFactor * dArea
                End If

                FillObjectInformation oObj
                If m_oBoundaryColl.Contains(oObj) Then
                    m_oFinalADORecordset.Fields("Condition").Value = "Boundary-Inside"
                Else
                    m_oFinalADORecordset.Fields("Condition").Value = "Inside"
                End If
                
                FillCompartCoatingAttributes
                
                m_oFinalADORecordset.Fields("Compartment").Value = oCompartNamedItem.Name
                m_oFinalADORecordset.AddNew
            End If
        Next
    End If
    
    Set oQueryData = Nothing
    Set oObjectsWhichCross = Nothing
    Set oObjectsWhichTouchOutside = Nothing
    Set oObjectsWhichTouchInside = Nothing
    Set oObjectsInside = Nothing
    Set oCompartService = Nothing
    Set oName = Nothing
    Set oObj = Nothing

    Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub
                    
Private Sub GenerateCompartBoundaryReport(oCompartment As Object, oBoundary As Object)
Const METHOD As String = "GenerateCompartBoundaryReport"
On Error GoTo ErrorHandler

    Dim oCompartService             As Object 'New Compartment
    Dim eAdjacencyType              As AdjacencyType
    Dim dLength As Double, dArea As Double, dVolume As Double
    Dim oCompartName                As IJNamedItem
    Dim oBoundaryName               As IJNamedItem
    
    Set oBoundaryName = oBoundary
    
    
    Set oCompartService = SP3DCreateObject("VolumeAttributeHelper.Compartment") '
                
    If Not oBoundaryName Is Nothing Then
        m_oFinalADORecordset.Fields("Name").Value = oBoundaryName.Name
    End If

'    oCompartService.CompartIntersectionArea oCompartment, oBoundary, eAdjacencyType, dLength, dArea, dVolume
    OverlapPaintArea oCompartment, oBoundary, dArea
    
    m_oFinalADORecordset.Fields("SurfaceArea").Value = dArea
    m_oFinalADORecordset.Fields("CoveringArea").Value = m_lSurfaceAreaFactor * dArea
    
    m_dTotalCoveringArea = m_dTotalCoveringArea + m_lSurfaceAreaFactor * dArea
    m_dTotalSurfaceArea = m_dTotalSurfaceArea + dArea
    
    'Fill the object type information of Compartment boundaries
    FillObjectInformation oBoundary
    
    m_oFinalADORecordset.Fields("Condition").Value = "Boundary"
    
    'Fill compartment coating attributes information
    FillCompartCoatingAttributes
    
    Set oCompartName = oCompartment
    m_oFinalADORecordset.Fields("Compartment").Value = oCompartName.Name

    Set oCompartName = Nothing
    Set oBoundaryName = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Sub FillPartCoatingAreaInfo(oObj As Object, bCoatingAreaAvlbl As Boolean)
Const METHOD As String = "FillPartCoatingAreaInfo"
On Error GoTo ErrorHandler
   
    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttributesCollection       As IJDAttributesCol
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim oCodelist                   As IJDCodeListMetaData
                    
    Set oCodelist = m_oPom
                    
    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
    
    bCoatingAreaAvlbl = False
    
    If oAttributeMetadata Is Nothing Then Exit Sub
    
    If oAttrHelper Is Nothing Then Exit Sub
    
    Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(oAttributeMetadata.IID("IJCoatingInfo"))
    
    If Not oInterfaceInfo Is Nothing Then
        On Error Resume Next

        Set oAttributesCollection = oAttrHelper.CollectionOfAttributes(oInterfaceInfo.Type)
        Err.Clear
        
        On Error GoTo ErrorHandler

        If Not oAttributesCollection Is Nothing Then
            For Each oAttribute In oAttributesCollection
                If oAttribute.AttributeInfo.Name = "CoatingArea" Then
                    If oAttribute.Value <> 0 Then
                        bCoatingAreaAvlbl = True
                        m_oFinalADORecordset.Fields("CoveringArea").Value = m_lSurfaceAreaFactor * oAttribute.Value
                        m_dTotalCoveringArea = m_dTotalCoveringArea + m_lSurfaceAreaFactor * oAttribute.Value
                    End If
'                ElseIf oAttribute.AttributeInfo.Name = "CoatingType" Then
'                    m_oFinalADORecordset.Fields("CoatingType").Value = oCodelist.LongStringValue(oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
'                ElseIf oAttribute.AttributeInfo.Name = "CoatingColor" Then
'                    m_oFinalADORecordset.Fields("CoatingColor").Value = oCodelist.LongStringValue(oAttribute.AttributeInfo.CodeListTableName, oAttribute.Value)
                End If
            Next
        End If
    End If
    
    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttributesCollection = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing
    Set oCodelist = Nothing
        
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub
'get the only overlapped area with the compartment
Private Sub OverlapPaintArea(oCompartment As Object, oOverlapObj As Object, dArea As Double)
Const METHOD = "OverlapPaintArea"
On Error GoTo ErrorHandler

    If TypeOf oOverlapObj Is IJPlatePart Then
        dArea = GetPlateOverlapPaintArea(oCompartment, oOverlapObj)
        Exit Sub
    End If

    Dim oIntersectionGeom       As IJDModelBody
    Dim dNewAcc                 As Double
    Dim oGeomService            As IJDCompartSpatialService
    Dim oGTTool                 As IJDTopologyToolBox
    Dim oBodies                 As IJElements
    Dim oCompartFaces           As IJElements
    Dim oObject                 As Object
    Dim i                       As Long
    Dim j                       As Long
    Dim dTArea                  As Double
    Dim oModel                  As IJDModelBody
    Dim oIJDTopologyIntersect   As IJDTopologyIntersect
    Dim oOverlapGeom            As Object
    Dim oFace                   As Object
    Dim bconsider               As Boolean
    Dim eAdjacency              As AdjacencyType
    
    Set oGeomService = New CompartGeomService
    
    Set oCompartFaces = New JObjectCollection
    Set oBodies = New JObjectCollection
    
    Set oIJDTopologyIntersect = New DGeomOpsIntersect
    Set oGTTool = New DGeomOpsToolBox
    
    On Error Resume Next
    
    Set oIntersectionGeom = oGeomService.OverlapGeometry(oCompartment, oOverlapObj, eAdjacency)
    
    If Not oCompartment Is Nothing Then
        Set oCompartFaces = oGeomService.Faces(oCompartment)
    End If
    
    If Not oIntersectionGeom Is Nothing Then
        oGTTool.ExplodeSurfaceBodyByFaces Nothing, oIntersectionGeom, oBodies
    End If
    
    dArea = 0#
    
    bconsider = True
    
   ' get the plate base and offset ports

    If (TypeOf oOverlapObj Is IJPlatePart) Then
        Dim oStructConnectbale As IJStructConnectable
        Dim oBasePort As IJPort
        Dim oOffsetPort As IJPort
        Dim oLateralPortsCol As IJElements
        
        Set oStructConnectbale = oOverlapObj
    
        oStructConnectbale.GetBaseOffsetLateralPorts vbNullString, False, oBasePort, oOffsetPort, oLateralPortsCol
    End If
        
    
    For i = 1 To oBodies.Count
        bconsider = True
        Set oObject = oBodies.Item(i)
        For j = 1 To oCompartFaces.Count
            Set oFace = oCompartFaces.Item(j)

            oIJDTopologyIntersect.PlaceOverlappingObject Nothing, oFace, oObject, Nothing, oOverlapGeom
            
            If Not oOverlapGeom Is Nothing Then
                bconsider = False
                Set oOverlapGeom = Nothing
                Exit For
            End If
            
            Set oOverlapGeom = Nothing
            
        Next j
        
        'ignore the plate edges in surface area calculation.
        
        If (TypeOf oOverlapObj Is IJPlatePart) And (bconsider = True) Then
            
            Dim bStatus As Boolean
            
            oIJDTopologyIntersect.HasOverlappingGeometry oObject, oBasePort.Geometry, bStatus

            If bStatus = False Then
                oIJDTopologyIntersect.HasOverlappingGeometry oObject, oOffsetPort.Geometry, bStatus
            End If
            
            bconsider = bStatus

        End If
        
        
        If (bconsider = True) Then
            Set oModel = oObject
            oModel.GetDimMetrics m_dTolerance, dNewAcc, 0#, dTArea, 0#
            dArea = dArea + dTArea
        End If
        
    Next i
    
    On Error GoTo ErrorHandler
    
    Set oGeomService = Nothing
    Set oIntersectionGeom = Nothing
    Set oObject = Nothing
    Set oFace = Nothing
    Set oOverlapGeom = Nothing
    Set oBodies = Nothing
    Set oCompartFaces = Nothing

Exit Sub
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Function GetPlateOverlapPaintArea(oCompartment As Object, oOverlapPlate As Object) As Double
Const METHOD = "GetPlateOverlapPaintArea"
On Error GoTo ErrorHandler

     Dim oBasePort                                  As IJPort
     Dim oOffsetPort                                As IJPort
     Dim oLateralPortsCol                           As IJElements
     Dim oStructConnectbale                         As IJStructConnectable
     Dim oBasePortSurfaceBody                       As IJSurfaceBody
     Dim oOffsetPortSurfaceBody                     As IJSurfaceBody
     Dim oGeomService                               As IJDCompartSpatialService
     Dim oIntersectionGeom                          As IJDModelBody
     Dim eAdjacency                                 As AdjacencyType
     Dim dBaseArea                                  As Double
     Dim dOffsetArea                                As Double
     
     Set oGeomService = New CompartGeomService
     Set oStructConnectbale = oOverlapPlate
     
     'get the  base and offset ports of the plate
     oStructConnectbale.GetBaseOffsetLateralPorts vbNullString, False, oBasePort, oOffsetPort, oLateralPortsCol
     
     'get the base and offset ports geomerty
     Set oBasePortSurfaceBody = oBasePort.Geometry
     Set oOffsetPortSurfaceBody = oOffsetPort.Geometry
     
     'get the overlap geomerty between compartment and base port geometry of plate
     'Intersection Geometry may not have
     On Error Resume Next
     Set oIntersectionGeom = oGeomService.OverlapGeometry(oCompartment, oBasePortSurfaceBody, eAdjacency)
     On Error GoTo ErrorHandler
     
     If Not oIntersectionGeom Is Nothing And eAdjacency = FaceAdjacency Then
        'get the overlap surface area between compartment and intersection geometry
        dBaseArea = GetOverlapSurfaceArea(oCompartment, oIntersectionGeom)
        Set oIntersectionGeom = Nothing
     End If
     
     'get the overlap geomerty between compartment and offset port geometry of plate
     'Intersection Geometry may not have
     On Error Resume Next
     Set oIntersectionGeom = oGeomService.OverlapGeometry(oCompartment, oOffsetPortSurfaceBody, eAdjacency)
     On Error GoTo ErrorHandler
     
     If Not oIntersectionGeom Is Nothing And eAdjacency = FaceAdjacency Then
        'get the overlap surface area between compartment and intersection geometry
        dOffsetArea = GetOverlapSurfaceArea(oCompartment, oIntersectionGeom)
     End If
     
    GetPlateOverlapPaintArea = dBaseArea + dOffsetArea
    
    Set oGeomService = Nothing
    Set oStructConnectbale = Nothing
    Set oBasePortSurfaceBody = Nothing
    Set oOffsetPortSurfaceBody = Nothing
    Set oIntersectionGeom = Nothing
     
Exit Function
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
     
End Function

Private Function GetOverlapSurfaceArea(oCompartment As Object, oIntersectionGeom As IJDModelBody) As Double
Const METHOD = "GetOverlapSurfaceArea"
On Error GoTo ErrorHandler
    
    Dim oCompartFaces           As IJElements
    Dim oGeomService            As IJDCompartSpatialService
    Dim oIJDTopologyIntersect   As IJDTopologyIntersect
    
    Dim bconsider               As Boolean
    Dim dTempArea               As Double
    Dim dNewAcc                 As Double
    Dim j                       As Integer
    
    Set oIJDTopologyIntersect = New DGeomOpsIntersect
    Set oGeomService = New CompartGeomService
    Set oIJDTopologyIntersect = New DGeomOpsIntersect
    bconsider = True
   
   'get the compartment faces
    If Not oCompartment Is Nothing Then
        Set oCompartFaces = oGeomService.Faces(oCompartment)
    End If
    
    For j = 1 To oCompartFaces.Count
        Dim oFace                   As Object
        Set oFace = oCompartFaces.Item(j)
        'check the overlap between comartment faces and intersection geometry.
        'if overlap geomerty is nothing then only get the surface area for the intersection geometry
        Dim oOverlapGeom            As Object
        
        'Overlap Geometry may not have
        On Error Resume Next
        oIJDTopologyIntersect.PlaceOverlappingObject Nothing, oFace, oIntersectionGeom, Nothing, oOverlapGeom
        On Error GoTo ErrorHandler
        
        If Not oOverlapGeom Is Nothing Then
            bconsider = False
            Exit For
        End If
    Next j
    
    'get the area for intersection geometry.
    If bconsider = True Then
        oIntersectionGeom.GetDimMetrics m_dTolerance, dNewAcc, 0#, dTempArea, 0#
    End If
    
    GetOverlapSurfaceArea = dTempArea
    
    Set oIJDTopologyIntersect = Nothing
    Set oGeomService = Nothing
    Set oCompartFaces = Nothing
    Set oIJDTopologyIntersect = Nothing
    Set oOverlapGeom = Nothing
Exit Function
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
