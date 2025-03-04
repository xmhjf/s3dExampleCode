VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlateMfgWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CPlateMfgWeight.cls
''
''  Description : Populates a csv file with plate detailing weight and mfg weight
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

Private Const MODULE = "CustomPlateReports.CPlateMfgWeight"
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
        oStream.WriteLine "UNIT,kg"
        oStream.WriteLine "Margin is applied (if any)."
        oStream.WriteLine "Part Name,Assembly,Dry Weight,Raw Mfg Weight,Bevels Applied,Chamfers Applied,Bevels + Chamfers Applied,Holes Applied,Non Trimmed Margin Applied,Final Mfg Weight"

        For Each oSelectedObj In pElements
            If TypeOf oSelectedObj Is IJMfgPlatePart Then
                Dim oMfgPlate           As IJMfgPlatePart
                Dim oPlatePart          As IJPlatePart
                
                ReportWeightInformation oSelectedObj, oStream
                
            ElseIf TypeOf oSelectedObj Is IJPlatePart Then
                'This is plate. Get Mfg Plate Part from plate
                
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMfgPlatePart)
                
                If Not oColl Is Nothing Then
                    If oColl.Count > 0 Then
                        Dim oMfgPlatePart As IJMfgPlatePart
                        Set oMfgPlatePart = oColl.Item(1) 'count will be one
                        ReportWeightInformation oMfgPlatePart, oStream
                    End If
                End If
                
            ElseIf TypeOf oSelectedObj Is IJAssembly Then
                Dim oSelectedPlateObj As Object
                Dim oAllPlates As IJElements
                Set oAllPlates = GetAllPlatePartsFromAssembly(oSelectedObj)
                For Each oSelectedPlateObj In oAllPlates
                
                    If TypeOf oSelectedPlateObj Is IJPlatePart Then
                        Dim oColl2 As IJElements
                        Dim oStructMfgGlobals2 As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                        Set oColl2 = oStructMfgGlobals2.GetMfgPart(oSelectedPlateObj, IJMfgPlatePart)
                        
                        If Not oColl2 Is Nothing Then
                            If oColl2.Count > 0 Then
                                Dim oMfgPlatePart2 As IJMfgPlatePart
                                Set oMfgPlatePart2 = oColl2.Item(1) 'count will be one
                                
                                ReportWeightInformation oMfgPlatePart2, oStream
                            End If
                        End If
                    End If
                    
                Next
                Set oAllPlates = Nothing
            End If
        Next
        
        oStream.Close
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub
Private Sub ReportWeightInformation(oMfgPlatePart As IJMfgPlatePart, oStream As TextStream)
 Const METHOD = "ReportMfgPlatePartObjectInformation"
    On Error GoTo ErrorHandler

    Dim oMfgNestData As IJMfgNestData
    Dim oPlatePart As IJPlatePart
    Dim oNameditem As IJNamedItem
    Dim strPartName As String
    Dim dPartTotalMfgWeight As Double, dPartNoBevelsMfgWeight As Double, dPartNoHolesMfgWeight As Double
    Dim dPartAllMfgWeight As Double, dPartNoChamfersMfgWeight As Double, dPartNoBevelsChamfersMfgWeight As Double
    Dim dPartMarginMfgWeight As Double
    
    Set oNameditem = oMfgPlatePart
    strPartName = oNameditem.Name
    
    oMfgPlatePart.GetDetailedPart oPlatePart
    
    Dim oIJPlate As IJPlate
    Set oIJPlate = oPlatePart
    
    
    Dim oEntityHelper As New GSCADStrMfgUtilities.MfgEntityHelper
    
    ' Without applying bevels, chamfers and hole
    dPartAllMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 0, 0, oIJPlate.thickness)
    
    ' With applying bevels, chamfers and hole
    dPartTotalMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 31, 0, oIJPlate.thickness)
    
    ' Apply bevels
    dPartNoBevelsMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 1, 0, oIJPlate.thickness)
     
    ' Apply chamfers
    dPartNoChamfersMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 2, 0, oIJPlate.thickness)
    
    ' Apply bevels and chamfers
    dPartNoBevelsChamfersMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 3, 0, oIJPlate.thickness)
    
    ' Apply hole
    dPartNoHolesMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 4, 0, oIJPlate.thickness)
    
    ' Apply margin
    dPartMarginMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgPlatePart, 16, 0, oIJPlate.thickness)
    
    Dim oWightCG As IJWeightCG
    Set oWightCG = oPlatePart
    
    Dim sAssemblyPath As String
    sAssemblyPath = GetAssemblyType(oPlatePart)
    
    'Get Nesting info of Mfg Plate Part
    oStream.WriteLine strPartName & "," & sAssemblyPath & "," _
                                & oWightCG.DryWeight & "," _
                                & dPartAllMfgWeight & "," _
                                & dPartNoBevelsMfgWeight & "," _
                                & dPartNoChamfersMfgWeight & "," _
                                & dPartNoBevelsChamfersMfgWeight & "," _
                                & dPartNoHolesMfgWeight & "," _
                                & dPartMarginMfgWeight & "," _
                                & oMfgPlatePart.MfgWeight(15)
                                
    
    Set oMfgPlatePart = Nothing
    Set oNameditem = Nothing
    
    'oStream.Close

    Exit Sub
ErrorHandler:
    'Instead of Erroring out, dump the surfaces which are causing error with ERROR_ as suffix
'    m_iCount = m_iCount + 1
'    strPath = GetTempFolder & "ERROR_" & Format(Date, "yyyy-mm-dd") & "_SurfaceGeometry_" & m_iCount & ".sat"
'    DumpSurfaceSatFile oObject, strPath
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



