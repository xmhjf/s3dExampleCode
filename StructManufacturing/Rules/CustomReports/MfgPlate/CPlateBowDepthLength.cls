VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBowDepthLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2017, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgPlateCustomReports
''  File        : CPlateBowDepthLength.cls
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

Private Const MODULE = "MfgPlateCustomReports.CPlateBowDepthLength"
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
        
        oStream.WriteLine "All Dimensions are in milli-meters  "
        oStream.WriteLine "Part Name, Thickness, Bow Depth, Bow Length, Curvature"

        For Each oSelectedObj In pElements
            If TypeOf oSelectedObj Is IJMfgPlatePart Then
                Dim oMfgPlate           As IJMfgPlatePart
                Dim oPlatePart          As IJPlatePart
                
                Set oMfgPlate = oSelectedObj
                oMfgPlate.GetDetailedPart oPlatePart
                
                ReportPlateBowDepthLengthInformation oPlatePart, oStream
            ElseIf TypeOf oSelectedObj Is IJPlatePart Then
                'This is plate. Get Mfg Plate Part from plate
                ReportPlateBowDepthLengthInformation oSelectedObj, oStream
            ElseIf TypeOf oSelectedObj Is IJAssembly Then
                Dim oSelectedPlateObj As Object
                Dim oAllPlates As IJElements
                Set oAllPlates = GetAllPlatePartsFromAssembly(oSelectedObj)
                For Each oSelectedPlateObj In oAllPlates
                    If TypeOf oSelectedPlateObj Is IJPlatePart Then
                      ReportPlateBowDepthLengthInformation oSelectedPlateObj, oStream
                    End If
                Next
                Set oAllPlates = Nothing
            End If
        Next
        
        oStream.Close
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub
Public Sub ReportPlateBowDepthLengthInformation(oPlateObj As Object, oStream As TextStream)

    On Error GoTo Cleanup
    
    Dim oNameditem As IJNamedItem
    Dim strPartName As String
    
    Set oNameditem = oPlateObj
    strPartName = oNameditem.Name
    
    ' Before doing anything, get curvature type and set as candidate return value.
    Dim bIsPlanar  As Boolean
    Dim dPlateThickness As Double
    Dim ePlateCurvature As PlateCurvature
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Dim oSDPlatePartSupport As GSCADSDPartSupport.IJPlatePartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PlatePartSupport
    Set oSDPartSupport.Part = oPlateObj
    
    Set oSDPlatePartSupport = oSDPartSupport
    
    oSDPlatePartSupport.GetThickness dPlateThickness
    oSDPlatePartSupport.IsPlanar 0.001, bIsPlanar
    
    ePlateCurvature = PLATE_CURVATURE_Curved
    
    If bIsPlanar Then
        oStream.WriteLine strPartName & "," & CStr(Round(dPlateThickness * 1000#, 2)) & "," & 0# & "," & 0# & ", Flat"
    Else
        Dim oMfgUnfold As New MfgUnfold
        Dim dBowDepth As Double
        Dim dBowLength As Double
        
        oMfgUnfold.EvaluatePlateBowDepthAndLength oPlateObj, dBowDepth, dBowLength
        
        Dim strQuery As String
        strQuery = "SELECT BowStringDepth From JUAMfgPlateCurvatureLim WHERE ( ( " & CStr(Round(dPlateThickness, 5)) & _
                   " > (ThicknessSizeMin-0.00001) and " & CStr(Round(dPlateThickness, 5)) & _
                   " < (ThicknessSizeMax+0.00001) ) and ( " & CStr(Round(dBowLength, 5)) & _
                   " > (LengthOfCurvatureMin-0.00001) and " & CStr(Round(dBowLength, 5)) & _
                   " < (LengthOfCurvatureMax+0.00001) )  )"
    
        Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
        Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
    
        Dim oQueryOutputValues() As Variant
        oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
                
        If (UBound(oQueryOutputValues) >= LBound(oQueryOutputValues)) Then
            Dim dMaxBowStringDepth As Double
            dMaxBowStringDepth = CDbl(oQueryOutputValues(LBound(oQueryOutputValues)))
    
            If dBowDepth < dMaxBowStringDepth Then
                ePlateCurvature = PLATE_CURVATURE_Flat
            End If
        End If
        
        If ePlateCurvature = PLATE_CURVATURE_Curved Then
            oStream.WriteLine strPartName & "," & CStr(Round(dPlateThickness * 1000#, 2)) & "," & CStr(Round(dBowDepth * 1000#, 2)) & "," & CStr(Round(dBowLength * 1000#, 2)) & ", Curved"
        Else
            oStream.WriteLine strPartName & "," & CStr(Round(dPlateThickness * 1000#, 2)) & "," & CStr(Round(dBowDepth * 1000#, 2)) & "," & CStr(Round(dBowLength * 1000#, 2)) & ", Flat"
        End If
     End If
     
Cleanup:
    Set oSDPartSupport = Nothing
    Set oSDPlatePartSupport = Nothing
    Set oMfgUnfold = Nothing
    Set oMfgCatalogQueryHelper = Nothing
    Set oNameditem = Nothing
    
    Exit Sub
ErrorHandler:
    'Instead of Erroring out, dump the surfaces which are causing error with ERROR_ as suffix
    m_iCount = m_iCount + 1
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



