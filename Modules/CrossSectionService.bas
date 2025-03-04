Attribute VB_Name = "CrossSectionService"
Option Explicit
'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : CrossSectionService.bas
'
'Author : A. Patil
'
'Description :
'    SmartPlant Structural Cross Section Services vb bas file
'
'History:
'
' 05/21/03   J.Schwartz     Commmented out msgbox
'
' 10 Sept 2004 Manish Dhotre Additional argument "SKinning Option" added to CreateProjectionFromCSProfile()
' Which be passed to CreateSurface for for appropriate creation of surface as per requirement through
' CreateBySingleSweep() function. For details see Documentaion of CreateBySingleSweep & CreateSurface
' 06/06/05  J. Schwartz     Fix for DI#79298, created IMSErrorLog.JServerErrors
'                           within the error handler of the GetCSAttribData
'                           method so TaskHost doesn't crash.
'********************************************************************
Private Const MODULE = "CrossSectionService::"
Private Const E_FAIL = -2147467259  'Manish- TR#41148
Private Const S_FALSE = &H1
Private bSymFileMsgDisplayed As Boolean
Private oLocalizer As IJLocalizer

'*************************************************************************
'Function
'GetCSProfile
'
'Abstract
'Use this method to get the cross section object from cache
'Arguments:
'Section Standard, Section Name, cache resource manager
'Return
'Cross section object
'
'Exceptions
'
'***************************************************************************
Public Function GetCSProfile(pOC As IJDOutputCollection, _
                              SecStandard As String, _
                              SecName As String, ByVal CatResMgr As IUnknown) As Object
                                   
Const METHOD = "GetCSProfile"

    Dim strCatlogDB As String
    Dim CatalogDef  As Object
    Dim oErrorColl As IJEditErrors
    Dim oError  As IJError
    
    Dim m_xService As SP3DStructGenericTools.CrossSectionServices

    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

    Dim sectype As String
    
    Set oErrorColl = GetJContext().GetService("Errors")
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & "SPSStairMacros"

    On Error GoTo exitGetCS
    m_xService.GetStructureCrossSectionDefinition CatResMgr, SecStandard, sectype, SecName, CatalogDef
         
    ' tr 51433 - get the cross section always and delete it later
    ' if bAlwaysPlaceNew is TRUE then handrail macros need to delete
    ' the cross section occurrence which is being done right now
    ' otherwise cross section occurrence should not be deleted
    

    Dim sSymName As String
    If DoesSymbolFileExist(CatalogDef, sSymName) Then ' call this only if have access to sym file
        If (Not pOC Is Nothing) Then
            m_xService.PlaceCrossSectionOccurrence pOC.ResourceManager, CatalogDef, True, GetCSProfile
        Else
            m_xService.PlaceCrossSectionOccurrence Nothing, CatalogDef, True, GetCSProfile  ' Create transient
        End If
    Else
        If bSymFileMsgDisplayed = False Then
        ' Symbols should not give out any message boxes and hence commented out the following msgbox code
'            MsgBox oLocalizer.GetString(IDS_SYMFILE_NOTACCESSSIBLE, ".Sym file of some cross-sections is not accessible"), , oLocalizer.GetString(IDS_SMARTPLANT3D, "SmartPlant 3D")
            bSymFileMsgDisplayed = True
        End If
    End If

    Set oErrorColl = Nothing
    Set oError = Nothing
    Set m_xService = Nothing
    Set CatalogDef = Nothing
    Set oLocalizer = Nothing
    Exit Function
exitGetCS:
    Dim oErrors As New IMSErrorLog.JServerErrors
    If Not oErrors Is Nothing Then
        For Each oError In oErrorColl
            If oError.ErrorContext = "IJCrossSectionServices" Then
                If oError.Number <> 0 Then
                    oErrors.Add oError.Number, METHOD, oError.Description
                    Set oError = Nothing
                    Exit For
                End If
            End If
        Next
    End If
    Set oError = Nothing
    Set oErrorColl = Nothing
    Set m_xService = Nothing
    Set oLocalizer = Nothing
End Function

Public Function CreateProjectionFromCSProfile(pOC As IJDOutputCollection, _
                                                ByVal pTraceCurve As Object, _
                                                ByVal oCSOcc As Object, _
                                                m_cp As Integer, _
                                                m_rotation As Double, _
                                                m_mirror As Integer, _
                                                StNorm As Variant, _
                                                EndNorm As Variant, _
                                                skinningOption As Long) As IJElements
                                   
Const METHOD = "CreateProjectionFromCSProfile"
           
    On Error GoTo ErrorHandler
    Dim pProfiles           As IJElements
    Dim oTrans4x4           As IJDT4x4
    Dim Occurrence          As Object
    Dim m_cpx               As Double
    Dim m_cpy               As Double
 
        
    Dim m_xService As SP3DStructGenericTools.CrossSectionServices



    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

    m_xService.GetProfiles oCSOcc, "SimplePhysical", pProfiles     '     DetailedPhysical
 
    m_xService.GetCardinalPoint oCSOcc, m_cp, m_cpx, m_cpy

    Set oTrans4x4 = New DT4x4
    m_xService.ComputeTransformForProjection pTraceCurve, m_mirror, m_rotation, m_cpx, m_cpy, oTrans4x4
                
    Dim i As Integer
    Dim sectionprof As SectionProfile
    For i = 1 To pProfiles.Count
        Set sectionprof = pProfiles.Item(i)
        sectionprof.Holes = Nothing
    Next i
    
    ' "bCapped" & "brkCrv" values sent to CreateBySingleSweepFunction() based on skinningOption are as below
    ' skinningOption     "bCapped"    brkCrv
    '   0                    0          0
    '   1                    1          0
    '   2                    0          1
    '   3                    1          1
    '   4                    0          2
    '   5                    1          2
    '   6                    0          3
    '   7                    1          3
'    BrkCv = 0 means the caller does not want any breaks in the cross section or trace curve
'    BrkCv = 1 means the caller would like the skinning to break the cross section, which is a complexstring, into individual curves and then skin.
'    BrkCv = 2 means the caller would like the skinning to break the trace curve, which is a complexstring, into individual curves and then skin.
'    BrkCv = 3 means the caller would like the skinning to break the cross section and the trace curve into individual curves before skinning.
    If (Not pOC Is Nothing) Then
        m_xService.CreateSurfaces pOC.ResourceManager, _
                              oTrans4x4, pProfiles, pTraceCurve, skinningOption, StNorm, EndNorm, CreateProjectionFromCSProfile
    Else
        ' Create transient
        m_xService.CreateSurfaces Nothing, _
                              oTrans4x4, pProfiles, pTraceCurve, skinningOption, StNorm, EndNorm, CreateProjectionFromCSProfile
    End If

    If Not Occurrence Is Nothing Then
        Dim otmp As iJDObject
        Set otmp = Occurrence
        otmp.Remove
        Set otmp = Nothing
    End If
    
    Set oTrans4x4 = Nothing
    Set m_xService = Nothing
    Set pProfiles = Nothing
    
Exit Function

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    If Not oErrors Is Nothing Then
        oErrors.Add Err.Number, METHOD, Err.Description
    End If
    Err.Raise Err.Number
End Function

Public Function CreateSolidbyPlanes(pConnection As IUnknown, LineString As IJLineString, Height As Double) As IJElements
Const METHOD = "CreateSolidbyPlanes"
 
    On Error GoTo ErrorHandler
    Set CreateSolidbyPlanes = New JObjectCollection 'IMSElements.DynElements
    If LineString Is Nothing Then
         Exit Function
    End If
    
    If Not LineString.IsClosed Then
         Exit Function
    End If
    
    Dim pts() As Double
    Dim NumPts As Long
    LineString.GetPoints NumPts, pts
    
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    Dim plane As IngrGeom3D.Plane3d
    Dim i As Integer, Ttlpts As Integer
    Dim Points() As Double
    Ttlpts = ((NumPts - 1) * 3 - 1)
    
    ReDim Points(Ttlpts) As Double
    
    For i = 0 To Ttlpts
         Points(i) = pts(i)  ' Exclude the last point as it is a closed string first and last are same points
    Next i
    
    Set plane = GeomFactory.Planes3d.CreateByPoints(pConnection, NumPts - 1, Points)
    CreateSolidbyPlanes.Add plane
    
    
    For i = 2 To Ttlpts Step 3
         Points(i) = Points(i) - Height
    Next i
    
    Dim TmpPoints() As Double
    ReDim TmpPoints(Ttlpts) As Double
    
    Dim k As Integer
    Dim n As Integer
    Dim Counter As Integer
    Counter = 0
    k = Ttlpts - 2
    n = k
    For i = 0 To Ttlpts
         TmpPoints(i) = Points(k)
         Counter = Counter + 1
         k = k + 1
         If Counter = 3 Then
              Counter = 0
              k = n - 3
              n = k
         End If
    Next i
    
    Set plane = GeomFactory.Planes3d.CreateByPoints(pConnection, NumPts - 1, TmpPoints)
    CreateSolidbyPlanes.Add plane
     
    ReDim TmpPoints(Ttlpts) As Double
    ReDim Points(Ttlpts) As Double
    
    For i = 0 To Ttlpts
         Points(i) = pts(i)  ' Exclude the last point as it is a closed string first and last are same points
    Next i
    
    
    Counter = 0
    For i = 0 To NumPts - 3
         
         TmpPoints(0) = Points(Counter)
         TmpPoints(1) = Points(Counter + 1)
         TmpPoints(2) = Points(Counter + 2)
         
         TmpPoints(3) = Points(Counter)
         TmpPoints(4) = Points(Counter + 1)
         TmpPoints(5) = Points(Counter + 2) - Height
         
         TmpPoints(6) = Points(Counter + 3)
         TmpPoints(7) = Points(Counter + 4)
         TmpPoints(8) = Points(Counter + 5) - Height
         
         TmpPoints(9) = Points(Counter + 3)
         TmpPoints(10) = Points(Counter + 4)
         TmpPoints(11) = Points(Counter + 5)
      
         Set plane = GeomFactory.Planes3d.CreateByPoints(pConnection, 4, TmpPoints)
         CreateSolidbyPlanes.Add plane
          
         Counter = Counter + 3
    Next i
    
    ReDim Points(Ttlpts) As Double
    For i = 0 To Ttlpts
         Points(i) = pts(i)  ' Exclude the last point as it is a closed string first and last are same points
    Next i
    
    'Plane between first and last point
    ReDim TmpPoints(Ttlpts) As Double
    TmpPoints(0) = Points(Ttlpts - 2)
    TmpPoints(1) = Points(Ttlpts - 1)
    TmpPoints(2) = Points(Ttlpts)
    
    TmpPoints(3) = Points(Ttlpts - 2)
    TmpPoints(4) = Points(Ttlpts - 1)
    TmpPoints(5) = Points(Ttlpts) - Height
    
    TmpPoints(6) = Points(0)
    TmpPoints(7) = Points(1)
    TmpPoints(8) = Points(2) - Height
    
    TmpPoints(9) = Points(0)
    TmpPoints(10) = Points(1)
    TmpPoints(11) = Points(2)
    
    Set plane = GeomFactory.Planes3d.CreateByPoints(pConnection, 4, TmpPoints)
    CreateSolidbyPlanes.Add plane
    
    Exit Function

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    If Not oErrors Is Nothing Then
        oErrors.Add Err.Number, METHOD, Err.Description
    End If
    Err.Raise Err.Number
End Function
'*************************************************************************
'Function
'GetCrossSection
'
'Abstract
'Use this method to get the cross section object
'Arguments:
'Section Standard, Section Name
'Return
'Cross section object
'
'Exceptions
'
'***************************************************************************
Public Function GetCrossSection(ByVal SectionStandard As String, _
                                  ByVal SectionName As String) As Object
Const METHOD = "GetCrossSection"
                     
    On Error GoTo ErrorHandler
    Dim obj As Object
    Dim oStructServices As New RefDataMiddleServices.StructCrossSectionServices
    oStructServices.GetStructureCrossSectionDefinition GetCatalogResourceManager, SectionStandard, "", SectionName, obj
    
    Set GetCrossSection = obj
        
    Exit Function

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    If Not oErrors Is Nothing Then
        oErrors.Add Err.Number, METHOD, Err.Description
    End If
    Err.Raise Err.Number
End Function
'*************************************************************************
'Function
'GetCSAttribData
'
'Abstract
'Use this method to get any data on the attribute of the cross section, given
'its CrossSectionName, CrossSection Standard, the Interface to which the Attribute
'belongs and the Attribute Name. This method returns the value stored in the
'Attribute.
'
'Return
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions
'
'***************************************************************************
Public Function GetCSAttribData(SectionName As String, SectionStandard As String, strInterfaceName As String, strAttributeName As String) As Variant

    Const METHOD = "GetCSAttribData"
    On Error GoTo ErrorHandler
                         
    Dim oAttrs As IJDAttributes
    Dim obj As Object
    Dim oStructServices As New RefDataMiddleServices.StructCrossSectionServices
    oStructServices.GetStructureCrossSectionDefinition GetCatalogResourceManager, SectionStandard, "", SectionName, obj
    
    Set oAttrs = obj
    If Not oAttrs Is Nothing Then
        Dim oCollectionProxy As CollectionProxy
        On Error Resume Next
        Set oCollectionProxy = oAttrs.CollectionOfAttributes(strInterfaceName)
        On Error GoTo ErrorHandler
        If Not oCollectionProxy Is Nothing Then
            Dim oAttributeProxy As AttributeProxy
            Set oAttributeProxy = oCollectionProxy.Item(strAttributeName)
            If Not oAttributeProxy Is Nothing Then
                GetCSAttribData = oAttributeProxy.Value
            End If
        End If
    End If
    
    Set obj = Nothing
    
    Exit Function
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    If Not oErrors Is Nothing Then
        oErrors.Add Err.Number, METHOD, Err.Description
    End If
    Err.Raise Err.Number
End Function

Public Function DoesSymbolFileExist(oCrossSection As Object, sSymName As String) As Boolean
    Const METHOD = "DoesSymbolFileExist"

    On Error GoTo ErrorHandler
    Dim oIJDCrossSectHelper As IJDSymbolDefHelper
    Dim oContext As IJContext
    Dim vPath As Variant
    Dim sSymbolName As String
    Dim nErrNo As Long
    
    DoesSymbolFileExist = False
    
    Set oIJDCrossSectHelper = oCrossSection
    sSymName = oIJDCrossSectHelper.definitionParameters
    If InStr(1, sSymName, "\") = 0 Then
        sSymName = "CrossSections\" & sSymName
    End If
    Set oContext = GetJContext
    vPath = oContext.GetVariable("OLE_SERVER")
    sSymName = vPath & "\" & sSymName

    On Error Resume Next
    sSymbolName = Dir(sSymName)
    nErrNo = Err.Number
    On Error GoTo ErrorHandler
    
    If Len(sSymbolName) > 0 Then
        DoesSymbolFileExist = True
    End If
    
    'Added following error check if symbols directory is accessed from local machine using UNC path
    'TR-CP·40130  Application error if 'Symbols' directory is not share during members placement.
    If nErrNo = 52 Then
        DoesSymbolFileExist = False
    End If
    
    Exit Function

ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    If Not oErrors Is Nothing Then
        oErrors.Add Err.Number, METHOD, Err.Description
    End If
End Function

Public Function DoesXSectionAndAccessToSymFileExists(strDefinitionProgID As String, strErrorContextPrefix As String, strSectionName As String, strSectionStd As String) As Boolean
    On Error GoTo ErrorHandler
    
    DoesXSectionAndAccessToSymFileExists = True
    
    Dim oXSectiionService As SP3DStructGenericTools.CrossSectionServices
    Dim strSectype As String
    Dim strSymName As String
    Dim oCatalogDef  As Object
    Dim oErrorColl As IJEditErrors
    
    Set oXSectiionService = New SP3DStructGenericTools.CrossSectionServices
    Set oErrorColl = GetJContext().GetService("Errors")
    
    On Error Resume Next
    oXSectiionService.GetStructureCrossSectionDefinition GetCatalogResourceManager, strSectionStd, strSectype, strSectionName, oCatalogDef
    On Error GoTo ErrorHandler
         
    If Not oCatalogDef Is Nothing Then
        If DoesSymbolFileExist(oCatalogDef, strSymName) Then
            ' User does have access to sym file, so no problem
        Else
            'report it in the middle tier error log. Client command looks for it, with the same context and error
            'returning S_FALSE as error, so that it will be logged as warning
            If Not oErrorColl Is Nothing Then
                oErrorColl.Add S_FALSE, strDefinitionProgID, strSymName + " is not accessible ", strErrorContextPrefix + ":SymFileNotAccessible"
            End If
            DoesXSectionAndAccessToSymFileExists = False
        End If
    Else
        'report it in the middle tier error log. Client command looks for it, with the same context and error
        'returning S_FALSE, so that it will be logged as warning
        If Not oErrorColl Is Nothing Then
            oErrorColl.Add S_FALSE, strDefinitionProgID, strSectionStd + "-" + strSectionName + " is not available in Catalog ", strErrorContextPrefix + ":XSectioninCatlogDefinitionMissing"
        End If
        DoesXSectionAndAccessToSymFileExists = False
    End If
        
    Set oCatalogDef = Nothing
    
    Exit Function
ErrorHandler:
    'Even if something is wrong in writting to error log, we can ignore.
    Err.Clear
End Function
