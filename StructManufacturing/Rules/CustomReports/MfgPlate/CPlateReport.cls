VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlateReport"
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

'- MfgPlate ----------------------------------------------------'
Private Const IJMfgPlatePart = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
            
    If pElements.Count > 0 Then
        Dim objPinJig As Object
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
            If Not TypeOf oSelectedObj Is IJMfgPlatePart Then
                'This is plate. Get Mfg Plate Part from plate
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMfgPlatePart)
                
                Dim oMfgPlatePart As Object
                Set oMfgPlatePart = oColl.Item(1) 'count will be one
                ReportMfgPlatePartObjectInformation oMfgPlatePart, strFileName
            Else
                ReportMfgPlatePartObjectInformation oSelectedObj, strFileName
            End If
        Next
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub
Private Sub ReportMfgPlatePartObjectInformation(oObject As Object, strLogFile As String)
 Const METHOD = "ReportMfgPlatePartObjectInformation"
    On Error GoTo ErrorHandler

    Dim oMfgPlatePart As IJMfgPlatePart
    Set oMfgPlatePart = oObject
                        
    Dim bFileExists As Boolean
    
    'Open log file
    Dim oStream As TextStream
    Dim oFSO As FileSystemObject
    Set oFSO = New FileSystemObject
    Set oStream = oFSO.OpenTextFile(strLogFile, ForWriting, True)
    Set oFSO = Nothing
    
    oStream.WriteLine "Reporting Manufacturing Plate information: "
    
    Dim oMfgNestData As IJMfgNestData
    Set oMfgNestData = oMfgPlatePart
    
    Dim oNameditem As IJNamedItem
    Dim strPartName As String
    Set oNameditem = oMfgPlatePart
    strPartName = oNameditem.Name
    
    'Get Nesting info of Mfg Plate Part
    oStream.WriteLine " "
    oStream.WriteLine "PartName : " & strPartName & " | " _
                    & "LotNumber : " & oMfgNestData.LotNumber & " | " _
                    & "OrderNumber : " & oMfgNestData.OrderNumber & " | " _
                    & "RegistrationTimestamp : " & oMfgNestData.RegistrationTimestamp & " | " _
                    & "NestingTimestamp : " & oMfgNestData.NestingTimestamp & " | " _
                    & "LotMaterialType : " & oMfgNestData.LotMaterialType & " | " _
                    & "LotMaterialGrade : " & oMfgNestData.LotMaterialGrade & " | " _
                    & "LotLength : " & oMfgNestData.LotLength & " | " _
                    & "LotWidth : " & oMfgNestData.LotWidth & " | " _
                    & "LotThickness : " & oMfgNestData.LotThickness
             
    oStream.WriteLine "============================= End of report ==================================="
     
     Set oMfgPlatePart = Nothing
     Set oMfgNestData = Nothing
     Set oNameditem = Nothing
    
    oStream.Close

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub
