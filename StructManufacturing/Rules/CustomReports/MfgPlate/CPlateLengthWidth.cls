VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlateLengthWidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CPlateReport.cls
''
''  Description : Populates a text file with PinJig related data
''
''
''  Author      : Intergraph Development.
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************
Implements IJDCustomReport

Private Const MODULE = "CustomPlateReports.CPlateReport"
Private Const TOLERANCE = 2

'- MfgPlate ----------------------------------------------------'
Private Const IJMfgPlatePart = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"
Private Declare Function GetTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long
         
Private m_iCount As Long

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
            
    If pElements.Count > 0 Then
        Dim objPinJig As Object
        Dim oSelectedObj As Object
        
        'Open log file
        Dim oStream As TextStream
        Dim oFSO As FileSystemObject
        Set oFSO = New FileSystemObject
        Set oStream = oFSO.OpenTextFile(strFileName, ForAppending, True)
        Set oFSO = Nothing
        
        oStream.WriteLine "All Dimensions are in meters  "
        oStream.WriteLine "Part Name, Assembly, Planar Length, Planar Width, Planar Accurate Length, Planar Accurate Width, Model Length, Model Width, Model Height, Model Range: Max X  ,Model Range:Max Y, Model Range:Max Z, Model Range: Min X, Model Range:Min Y, Model Range:Min Z, Time(Sec), SAT File"

        For Each oSelectedObj In pElements
            If TypeOf oSelectedObj Is IJMfgPlatePart Then
                Dim oMfgPlate           As IJMfgPlatePart
                Dim oPlatePart          As IJPlatePart
                
                Set oMfgPlate = oSelectedObj
                oMfgPlate.GetDetailedPart oPlatePart
                
                ReportPlateLengthWidthInformation oPlatePart, oStream
            ElseIf TypeOf oSelectedObj Is IJPlatePart Then
                'This is plate. Get Mfg Plate Part from plate
                ReportPlateLengthWidthInformation oSelectedObj, oStream
            ElseIf TypeOf oSelectedObj Is IJAssembly Then
                Dim oSelectedPlateObj As Object
                Dim oAllPlates As IJElements
                Set oAllPlates = GetAllPlatePartsFromAssembly(oSelectedObj)
                For Each oSelectedPlateObj In oAllPlates
                    If TypeOf oSelectedPlateObj Is IJPlatePart Then
                      ReportPlateLengthWidthInformation oSelectedPlateObj, oStream
                    End If
                Next
                Set oAllPlates = Nothing
            End If
        Next
        
        oStream.Close
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub
Private Sub ReportPlateLengthWidthInformation(oObject As Object, oStream As TextStream)
 Const METHOD = "ReportMfgPlatePartObjectInformation"
    On Error GoTo ErrorHandler

    Dim oMfgPlatePart As IJMfgPlatePart
    Dim oMfgNestData As IJMfgNestData
    Dim oNameditem As IJNamedItem
    Dim strPartName As String
    
    Set oNameditem = oObject
    strPartName = oNameditem.Name
    
    Dim dPlateWidth As Double
    Dim dPlateLength As Double
    
    Dim oPlnIntHelper As IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper
    
    Dim lStartTimer As Long
    Dim lEndTimer As Long
    
    lStartTimer = Timer
    
    Dim oPlatePart As IJPlatePart
    Set oPlatePart = oObject

    dPlateLength = oPlatePart.PlateLength
    dPlateWidth = oPlatePart.PlateWidth

    Dim strPath As String
    Dim dPlAccurateWidth As Double, dPlAccurateLength As Double
    Dim oBox    As GSCADPlnIntHelper.GBox
    
    'oPlnIntHelper.GetPlateWidthAndLength oObject, dPlateWidth, dPlateLength
    If dPlateLength <= 0 Or dPlateWidth <= 0 Then
        oPlnIntHelper.GetPlateDimension oObject, False, False, dPlateLength, dPlateWidth, oBox
    End If
    
    oPlnIntHelper.GetPlateDimension oObject, True, True, dPlAccurateLength, dPlAccurateWidth, oBox
    
    lEndTimer = Timer
    If (lEndTimer - lStartTimer) > TOLERANCE Or dPlateLength <= 0 Or dPlateWidth <= 0 Then
        m_iCount = m_iCount + 1
        strPath = GetTempFolder & Format(Date, "yyyy-mm-dd") & "_SurfaceGeometry_" & m_iCount & ".sat"
        DumpSurfaceSatFile oObject, strPath
    End If

    Dim sAssemblyPath As String
    sAssemblyPath = GetAssemblyType(oObject)
    
    oStream.WriteLine strPartName & "," & sAssemblyPath & "," & dPlateLength & "," & dPlateWidth & "," & dPlAccurateLength & "," & dPlAccurateWidth & "," & _
                        (oBox.m_high.x - oBox.m_low.x) & "," & (oBox.m_high.y - oBox.m_low.y) & "," & (oBox.m_high.z - oBox.m_low.z) & "," & _
                        oBox.m_high.x & "," & oBox.m_high.y & "," & oBox.m_high.z & "," & oBox.m_low.x & "," & oBox.m_low.y & "," & oBox.m_low.z & "," & _
                        CStr(lEndTimer - lStartTimer) & "," & strPath
                        
    Set oMfgPlatePart = Nothing
    Set oMfgNestData = Nothing
    Set oNameditem = Nothing
    
    'oStream.Close

    Exit Sub
ErrorHandler:
    'Instead of Erroring out, dump the surfaces which are causing error with ERROR_ as suffix
    m_iCount = m_iCount + 1
    strPath = GetTempFolder & "ERROR_" & Format(Date, "yyyy-mm-dd") & "_SurfaceGeometry_" & m_iCount & ".sat"
    DumpSurfaceSatFile oObject, strPath
    'Err.Raise Err.Number, , Err.Description
End Sub


' ***********************************************************************************
' Public Function GetAssemblyType
'
' Description:  Method will give the Type of Assembly
'
' ***********************************************************************************
Private Function GetAssemblyType(oPart As Object) As String
    Const METHOD = "GetAssemblyType"
    On Error GoTo ErrorHandler

    Dim bstrAssemblyPath As String
    
    ' Get Assembly hierarchy
    Dim oAssemblyChild As IJAssemblyChild
    Set oAssemblyChild = oPart
StartAgain:
    If Not oAssemblyChild Is Nothing Then
        Dim oAssemblyParentUnk As IUnknown
        Set oAssemblyParentUnk = oAssemblyChild.Parent
        If Not oAssemblyParentUnk Is Nothing Then
        If TypeOf oAssemblyParentUnk Is IJAssembly Then
            Dim oNamed As IJNamedItem
            Set oNamed = oAssemblyParentUnk
            If Len(bstrAssemblyPath) = 0 Then
                bstrAssemblyPath = oNamed.Name
            Else
                bstrAssemblyPath = oNamed.Name & " | " & bstrAssemblyPath
            End If
            If TypeOf oAssemblyParentUnk Is IJAssemblyChild Then
                Set oAssemblyChild = oAssemblyParentUnk
                GoTo StartAgain
            End If
            Set oNamed = Nothing
        End If
        End If
        Set oAssemblyParentUnk = Nothing
    End If

    GetAssemblyType = bstrAssemblyPath
    
    Exit Function
    
ErrorHandler:
   Err.Raise Err.Description
End Function

Private Sub DumpSurfaceSatFile(oPlatePart As Object, strPath As String)
    Const METHOD = "DumpSurfaceSatFile"
    On Error GoTo ErrorHandler
    
    Dim oPartSupp As IJPartSupport
    Dim oPlatePartSupp As IJPlatePartSupport
    Set oPartSupp = New PlatePartSupport
    Set oPlatePartSupp = oPartSupp
    Set oPartSupp.Part = oPlatePart

    Dim pSurface As IJSurfaceBody
    oPlatePartSupp.GetSurface PlateBaseSide, pSurface
    
    Dim oNamed As IJNamedItem
    Set oNamed = oPlatePart
    
    Dim oMB1 As IJDModelBody
    Set oMB1 = pSurface
    oMB1.DebugToSATFile strPath
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function GetAllPlatePartsFromAssembly(ByVal oElem As GSCADAsmHlpers.IJAssembly) As IJElements
    Const METHOD = "GetAllPlatePartsFromAssembly"
    
    On Error GoTo ErrorHandler
    
    Set GetAllPlatePartsFromAssembly = New JObjectCollection
    
    'Get children of assembly
    Dim oAssembly As IJAssembly
    Set oAssembly = oElem
    
    Dim oChildren As IJDTargetObjectCol
    Set oChildren = oAssembly.GetChildren
    
    'Only consider assemblies with assembly children
    If oChildren.Count > 1 Then
        Dim Index As Long
        Dim oItem As Object
                    
        'Get each assembly child and add child to elements list
        For Index = 1 To oChildren.Count
            'Get next item
            Set oItem = oChildren.Item(Index)
                                           
            'Add item to elements list if type of item
            'is IJPlatePart and not MfgPlatePart
            If TypeOf oItem Is IJPlatePart Then
                GetAllPlatePartsFromAssembly.Add oItem
            ElseIf TypeOf oItem Is IJAssembly Then 'Assembly or block is child
                'get parts in this assembly
                On Error Resume Next
                GetAllPlatePartsFromAssembly.AddElements GetAllPlatePartsFromAssembly(oItem)
            End If
        Next
    End If

Cleanup:
    'Clean up
    Set oAssembly = Nothing
    Set oChildren = Nothing
    Set oItem = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    GoTo Cleanup
End Function

Private Function GetTempFolder() As String

    Const METHOD = "GetTempFolder"
    
    On Error GoTo ErrorHandler
   ' Returns the path to the user's Temp folder. To boot, Windows
   ' requires that a temporary folder exist, so this should always
   ' safely return a path to one. Just in case, though, check the
   ' return value of GetTempPath.
   Dim strTempPath As String
   Dim lngTempPath As Long
   
   ' Fill string with null characters.
   strTempPath = String(144, vbNullChar)
   ' Get length of string.
   lngTempPath = Len(strTempPath)
   ' Call GetTempPath, passing in string length and string.
   If (GetTempPath(lngTempPath, strTempPath) > 0) Then
      ' GetTempPath returns path into string.
      ' Truncate string at first null character.
      GetTempFolder = Left(strTempPath, _
         InStr(1, strTempPath, vbNullChar) - 1)
   Else
      GetTempFolder = ""
   End If
   
   Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

