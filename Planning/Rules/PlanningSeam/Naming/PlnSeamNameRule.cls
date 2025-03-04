VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlnSeamNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
' Copyright (C) 2004-2009, Intergraph Corporation.  All rights reserved.
'
' Project:
'   plnNameRules.plnSeamNameRule
'
' Abstract:
'   The file contains the VB-namerule for Seams created through manage block intersection command
'
' Author:
'  N.Apparao Created for  CR-CP-124697  Merge Planning and Design Splits
' History:
'******************************************************************

Option Explicit

Implements IJDMoldedFormNamingSolver
                                                        
Private Const Module = "PlnSeamNameRule: "
Private Const MODELDATABASE = "Model"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.
Private Const strCountFormat = "0000"


Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

Private Function IJDMoldedFormNamingSolver_ComputeName(ByVal pMoldedFormObject As Object) As String
Const METHOD = "IJDMoldedFormNamingSolver_ComputeName"
On Error GoTo ErrorHandler

    
    Dim oName As IJNamedItem
    Dim strBlock1Name As String
    Dim strBlock2Name As String
    Dim oStructCustomGeom As IJDStructCustomGeometry
    Dim oParentCol As IJElements
    Dim strProgID As String
    Dim oSplitter As Object
    Dim oPlaneSurface As GSCADSurfaces.IJPlaneSurface

    Dim strBaseName As String
    Set oStructCustomGeom = pMoldedFormObject
    Set oParentCol = New JObjectCollection

    oStructCustomGeom.GetCustomGeometry strProgID, oParentCol
    
    If oParentCol.Count >= 2 Then
    
        Set oSplitter = oParentCol.Item(2)
        
        If TypeOf oSplitter Is IJPlaneSurface Then
            Set oPlaneSurface = oSplitter 'oIntersection.IntersectingSurface
        End If
        
        If Not oPlaneSurface Is Nothing Then
            Dim oCuttingSurface As GSCADSurfaces.IJCuttingSurface
            Set oCuttingSurface = oPlaneSurface.CuttingSurface
            If Not oCuttingSurface Is Nothing Then
                Dim oBlockConn As IJBlockConnections
                Set oBlockConn = oCuttingSurface
                
                Dim oFace1 As IJPlaneSurface
                Dim oFace2 As IJPlaneSurface
                
                oBlockConn.FacesFromOverlap oPlaneSurface, oFace1, oFace2
                
                Dim oVolume As IJVolume
                Set oVolume = oFace2.Volume
                Set oName = Nothing
                Set oName = oVolume.Block
                
                ' Save selected Block Name for Intersection
                strBlock1Name = oName.Name
                Set oVolume = oFace1.Volume
                Set oName = Nothing
                Set oName = oVolume.Block
                strBlock2Name = oName.Name
                
                If TypeOf pMoldedFormObject Is IJSeam Then
                    strBaseName = strBlock1Name & "-" & strBlock2Name & "-" & "PSeam"
                ElseIf TypeOf pMoldedFormObject Is IJSeamPoint Then
                    strBaseName = strBlock1Name & "-" & strBlock2Name & "-" & "PSeamPoint"
                End If
            End If
        End If
        If oParentCol.Count >= 1 Then
            Dim JContext As IJContext
            Dim oDBTypeConfig As IJDBTypeConfiguration
            Dim oConnectMiddle As IJDAccessMiddle
            Dim strModelDBID As String
            Dim oModelResourceMgr As IUnknown
            Dim oNameCounter As IJNameCounter
            Dim strLocation As String
            Dim nCount As Long
            Dim oPart As Object
            Dim oChildNamedItem             As IJNamedItem
            Dim strChildName                As String
            'Get the connection to the model database
            Set JContext = GetJContext()
            
            Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
            Set oConnectMiddle = JContext.GetService("ConnectMiddle")
            
            strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
            Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
            
            Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
            
            Set oPart = oParentCol.Item(1)
            Set oChildNamedItem = oPart
        
            'Get the parent's name
            strChildName = oChildNamedItem.Name
            
            strLocation = vbNullString
            'GetCountEx:Returns the number of occurrence of a string in addtion to the LocationID
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName, strLocation)
            
            If strLocation <> vbNullString Then
                strChildName = "-" & strLocation & "-" & Format(nCount, strCountFormat)
            Else
                strChildName = "-" & Format(nCount, strCountFormat)
            End If
            
            
        End If
        
    End If
   
    IJDMoldedFormNamingSolver_ComputeName = strBaseName & strChildName
    
    Set oName = Nothing
    Set oVolume = Nothing
            
    
    
Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function

Private Sub IJDMoldedFormNamingSolver_ReleaseNamingResponsibility(ByVal pMoldedFormObject As Object)

End Sub

Private Sub IJDMoldedFormNamingSolver_TakeNamingResponsibility(ByVal pMoldedFormObject As Object)

End Sub
