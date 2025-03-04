Attribute VB_Name = "Common"
' ***************************************************************************
' Copyright (c) 2009, Intergraph Corporation.  All rights reserved.
'
' Project
'   GSCAD Planning Task - Sequence Rule
'
' File
'   Common.cls
'
' Author
'
'
' History
'
'
' ***************************************************************************

Public Const PROJECT_ID As String = "SMSequenceRules"

Option Explicit

Private Const Module As String = "Common"
Public Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Public m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Public m_oError As IJEditError         ' The error to raise.

Public Const GET_DIRECTION_FILTER_NONE As Long = 0
Public Const GET_DIRECTION_FILTER_CONTINUOUS As Long = 1

Public Const SEQUENCE_FILTER_ALL As Long = &HFFFFFFFF
Public Const SEQUENCE_FILTER_BASEPLATE_NON_HULL = &H1
Public Const SEQUENCE_FILTER_BASEPLATE_HULL As Long = &H2
Public Const SEQUENCE_FILTER_STIFFENING_PROFILE As Long = &H4
Public Const SEQUENCE_FILTER_CONNECTED_PLATE_NON_HULL As Long = &H8
Public Const SEQUENCE_FILTER_CONNECTED_PLATE_HULL As Long = &H10
Public Const SEQUENCE_FILTER_CONNECTED_PROFILE As Long = &H20

Public Type PART_DISTANCE_INFO
   oPart As Object
   dDistToPlateCenter As Double
End Type

Public Type PART_INFO
   oPart As Object
   bIsSystem As Boolean
   oLeafSystem As Object
   oRootSystem As Object
   dDistToPlateCenter As Double
   oDirection As IJDVector
   bAlongU As Boolean
   bAlongV As Boolean
   bMainDirection As Boolean
   bPerpendicularToMain As Boolean
End Type

Public Type PLATE_INFO
   oPlate As Object
   bIsSystem As Boolean
     
   oPlateParts() As PART_INFO
   oStiffeningProfileParts() As PART_INFO
   oConnectedPlateParts() As PART_INFO
   oConnectedHullParts() As PART_INFO
   oConnectedProfileParts() As PART_INFO
   
   nPlatePartCount As Long
   nStiffeningProfilePartCount As Long
   nConnectedPlatePartCount As Long
   nConnectedHullPartCount As Long
   nConnectedProfilePartCount As Long
End Type

Private Const m_bLogToFile As Boolean = False



' ***************************************************************************
'
' Function
'   GetAssemblyChildren()
'
' Abstract
'   Returns an elements collection containing the children of the given
'   assembly.
'
' ***************************************************************************

Public Function GetAssemblyChildren( _
   ByVal oAssembly As GSCADAsmHlpers.IJAssembly) As IJElements

   Const Method As String = "GetAssemblyChildren"
   On Error GoTo ErrorHandler
     
   Dim oPlnInitHelper As IJDPlnIntHelper
   
   Set oPlnInitHelper = New CPlnIntHelper
   Set GetAssemblyChildren = oPlnInitHelper.GetHierarchyAssemblyChildren(oAssembly, False)

   Set oPlnInitHelper = Nothing
    
   Exit Function
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & Method)
   m_oError.Raise
    
End Function

' ***************************************************************************
'
' Function
'   IsChildSequenced()
'
' Abstract
'   Returns true if child sequenced
'
'
' ***************************************************************************

Public Function IsChildSequenced(oObject As Object, oAssemblyChild As IJAssemblyChild) As Boolean
   Const Method As String = ".IsChildSequenced"
   On Error GoTo ErrorHandler
   
   Dim oAssemblySequence As IJAssemblySequence
   
   Set oAssemblySequence = oObject
   If oAssemblySequence.ChildIndex(oAssemblyChild) <> -1 Then
      IsChildSequenced = True
   Else
      IsChildSequenced = False
   End If

   Exit Function
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & Method)
   m_oError.Raise
    
End Function


' ***************************************************************************
'
' Function
'   AssignAnIndex()
'
' Abstract
'   Assigns a sequence index to a child of an assembly
'
' ***************************************************************************

Public Sub AssignAnIndex(ByVal index As Integer, Element As Object)

   Const Method As String = ".AssignAnIndex"
   On Error GoTo ErrorHandler

   Dim oAssemblySequence As IJAssemblySequence
   
   Set oAssemblySequence = Element
   If TypeOf Element Is IJPlnUnprocessedParts Then
      Exit Sub
   End If
    
   If index > 0 Then
      ' Set index number of element
      oAssemblySequence.ChildIndex(Element) = index
   End If

   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & Method)
   m_oError.Raise
End Sub
' ***************************************************************************
'
' Function
'   MaxSeqIndex()
'
' Abstract
'   Returns the heighst index of sequenced children
'
'
' ***************************************************************************
Public Function MaxSeqIndex(oObject As Object) As Long

   Const Method As String = ".MaxSeqIndex"
   On Error GoTo ErrorHandler

   Dim oElements As IJElements
   Dim oAssemblySequence As IJAssemblySequence
   
   Set oElements = GetAssemblyChildren(oObject)
   Set oAssemblySequence = oObject
   
   Dim index As Integer
   Dim MaxIndex As Long
   Dim oElement As Object
   
   MaxIndex = -1
   For index = 1 To oElements.Count
      Set oElement = oElements.Item(index)
      If Not TypeOf oElement Is IJPlnUnprocessedParts Then
         If MaxIndex < oAssemblySequence.ChildIndex(oElement) Then
            MaxIndex = oAssemblySequence.ChildIndex(oElement)
         End If
      End If
   Next index
   MaxSeqIndex = MaxIndex
    
   Set oElements = Nothing
   Set oElement = Nothing
   Set oAssemblySequence = Nothing
   
   Exit Function
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & Method)
   m_oError.Raise
 End Function
 
' Method:
'   GetBasePlates
' Descriptions:
'   Get all the plates that have the same normal as assembly base plate normal
'
Public Sub GetBasePlates(ByVal oAssembly As IJAssembly, _
                         ByVal oPlateParts As IJElements, _
                         ByRef nPlateCount As Long, _
                         ByRef arrayBasePlateInfo() As PLATE_INFO, _
                         ByRef oViewMatrix As IJDT4x4, _
                         Optional ByVal nBasePlateOption As Long = SEQUENCE_FILTER_BASEPLATE_NON_HULL)
   Const sMethod As String = "GetBasePlates"
   On Error GoTo ErrorHandler

   ' Get base plate normal
   Dim oAssemblyOrientation As IJAssemblyOrientation
   Dim oBasePlateNormal As IJDVector

   Set oAssemblyOrientation = oAssembly
   Set oViewMatrix = oAssemblyOrientation.ViewMatrix
   Set oBasePlateNormal = New DVector
   oBasePlateNormal.Set oViewMatrix.IndexValue(8), oViewMatrix.IndexValue(9), oViewMatrix.IndexValue(10)
   
   Dim bIncludeNonHull As Boolean
   Dim bIncludeHull As Boolean
   
   If (nBasePlateOption And SEQUENCE_FILTER_BASEPLATE_NON_HULL) > 0 Then
      bIncludeNonHull = True
   Else
      bIncludeNonHull = False
   End If
   
   If (nBasePlateOption And SEQUENCE_FILTER_BASEPLATE_HULL) > 0 Then
      bIncludeHull = True
   Else
      bIncludeHull = False
   End If
      
   Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim oSDHelper As New StructDetailHelper
   Dim oTopolocate As IJTopologyLocate
   Dim oStructGraph As IJStructGraph
   Dim oPlateRootSystem As IJSystem
   Dim nPlateIndex As Long
   Dim bExists As Boolean
   Dim sName As String
   Dim oPlateBasePort As IJPort
   Dim oPlatePortGeom As Object
   Dim oPlateNormal As IJDVector
   Dim oDummyPoint As IJDPosition
   Dim dDot As Double
   Dim bProcess As Boolean
   Dim nPartIndex As Long
   Dim oPart As Object
   Dim oPlate As IJPlate

   Set oTopolocate = New GSCADStructGeomUtilities.TopologyLocate
   
   For nPartIndex = 1 To oPlateParts.Count
      Set oPart = oPlateParts.Item(nPartIndex)
      Set oPlate = oPart
      bProcess = True
      
      If oPlate.plateType = Hull Then
         If bIncludeHull = True Then
            bProcess = True
         Else
            bProcess = False
         End If
      Else
         If bIncludeNonHull = True Then
            bProcess = True
         Else
            bProcess = False
         End If
      End If
            
      If bProcess = True Then
         Set oPlatePartWrapper.Object = oPart
         Set oPlateBasePort = oPlatePartWrapper.BasePort(BPT_Base)
         Set oPlatePortGeom = oPlateBasePort.Geometry
         Set oPlateBasePort = Nothing
         oTopolocate.FindApproxCenterAndNormal oPlatePortGeom, oDummyPoint, oPlateNormal
         Set oPlatePortGeom = Nothing
         Set oDummyPoint = Nothing
         
         dDot = oBasePlateNormal.Dot(oPlateNormal)
         
         If Abs(Abs(dDot) - 1) < 0.01 Then ' 0.01 is selected arbitrarily
            ' Part normal is close to base plate normal
            Set oStructGraph = oPart
            oSDHelper.IsPartDerivedFromSystem oStructGraph, oPlateRootSystem, True
      
            ' Check if part
            bExists = False
            For nPlateIndex = 0 To nPlateCount - 1
               Dim stPlateInfo As PLATE_INFO
               
               stPlateInfo = arrayBasePlateInfo(nPlateIndex)
               If oPlateRootSystem Is Nothing Then
                  ' Stand alone part
                  If oPart Is stPlateInfo.oPlate Then
                     bExists = True
                     Exit For
                  End If
               Else
                  ' System derived part
                  If oPlateRootSystem Is stPlateInfo.oPlate Then
                     bExists = True
                     Exit For
                  End If
               End If
            Next
            
            If bExists = False Then
               Dim stNewPlateInfo As PLATE_INFO
               
               stNewPlateInfo.nConnectedHullPartCount = 0
               stNewPlateInfo.nConnectedPlatePartCount = 0
               stNewPlateInfo.nConnectedProfilePartCount = 0
               stNewPlateInfo.nStiffeningProfilePartCount = 0
               stNewPlateInfo.nPlatePartCount = 0
               If oPlateRootSystem Is Nothing Then
                  stNewPlateInfo.bIsSystem = False
                  Set stNewPlateInfo.oPlate = oPart
               Else
                  stNewPlateInfo.bIsSystem = True
                  Set stNewPlateInfo.oPlate = oPlateRootSystem
               End If
               
               stNewPlateInfo.nPlatePartCount = stNewPlateInfo.nPlatePartCount + 1
               ReDim stNewPlateInfo.oPlateParts(stNewPlateInfo.nPlatePartCount)
               Set stNewPlateInfo.oPlateParts(stNewPlateInfo.nPlatePartCount - 1).oPart = oPart
               Set stNewPlateInfo.oPlateParts(stNewPlateInfo.nPlatePartCount - 1).oDirection = oPlateNormal
               nPlateCount = nPlateCount + 1
               ReDim Preserve arrayBasePlateInfo(nPlateCount)
               arrayBasePlateInfo(nPlateCount - 1) = stNewPlateInfo
            Else
               arrayBasePlateInfo(nPlateIndex).nPlatePartCount = arrayBasePlateInfo(nPlateIndex).nPlatePartCount + 1
               stPlateInfo = arrayBasePlateInfo(nPlateIndex)
               ReDim Preserve arrayBasePlateInfo(nPlateIndex).oPlateParts(stPlateInfo.nPlatePartCount)
               Set arrayBasePlateInfo(nPlateIndex).oPlateParts(stPlateInfo.nPlatePartCount - 1).oPart = oPart
               Set arrayBasePlateInfo(nPlateIndex).oPlateParts(stPlateInfo.nPlatePartCount - 1).oDirection = oPlateNormal
            End If
         Else
            '
         End If
         Set oPlateNormal = Nothing
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
End Sub

' Method:
'   GetStiffeningProfileParts
' Descriptions:
'   For each base plate, get all the stiffening profile parts
'
Public Sub GetStiffeningProfileParts( _
        ByVal oAssembly As IJAssembly, _
        ByVal oProfileParts As IJElements, _
        ByVal nPlateCount As Long, _
        ByRef arrayBasePlateInfo() As PLATE_INFO)
   Const sMethod As String = "GetStiffeningProfileParts"
   On Error GoTo ErrorHandler

   If oProfileParts Is Nothing Then
      Exit Sub
   End If
   
   ' Get all stiffening profiles on each deck
   Dim oProfileUtil As IJProfileAttributes
   Dim nStiffeningProfileCount As Long
   Dim nPartIndex As Long
   Dim oPart As Object
   Dim oStiffener As IJStiffener
   Dim oStiffenedPlate As IJPlate
   
   Dim nPlateIndex As Long
   Dim bIsStiffeningProfile As Boolean
   Dim oSDHelper As New StructDetailHelper
   Dim oStructGraph As IJStructGraph
   Dim oLeafSystem As IJSystem
   Dim oRootSystem As IJSystem
   Dim oStructEntityUtil As IJStructEntityUtils
   
   Set oStructEntityUtil = New GSCADCreateModifyUtilities.StructEntityUtils
   Set oProfileUtil = New ProfileUtils
   nStiffeningProfileCount = 0
   For nPartIndex = 1 To oProfileParts.Count
      Set oPart = oProfileParts.Item(nPartIndex)
      If TypeOf oPart Is IJStiffener Then
         Set oStiffener = oPart
         Set oStiffenedPlate = oStiffener.PlateSystem
         If TypeOf oStiffenedPlate Is IJSystem Then
            Set oRootSystem = oStructEntityUtil.RootSystem(oStiffenedPlate)
            If TypeOf oRootSystem Is IJPlate Then
               Set oStiffenedPlate = oRootSystem
            End If
         End If

         bIsStiffeningProfile = False
         For nPlateIndex = 0 To nPlateCount - 1
            If oStiffenedPlate Is arrayBasePlateInfo(nPlateIndex).oPlate Then
               Dim oAssemblyChild As IJAssemblyChild
               
               Set oAssemblyChild = oPart
               If Not oAssemblyChild Is Nothing Then
                  If oAssemblyChild.Parent Is oAssembly Then
                     bIsStiffeningProfile = True
                     Exit For
                  End If
               End If
            End If
         Next
         
         If bIsStiffeningProfile = True Then
            Set oStructGraph = oPart
            oSDHelper.IsPartDerivedFromSystem oStructGraph, oLeafSystem, False
            
            arrayBasePlateInfo(nPlateIndex).nStiffeningProfilePartCount = arrayBasePlateInfo(nPlateIndex).nStiffeningProfilePartCount + 1
            nStiffeningProfileCount = arrayBasePlateInfo(nPlateIndex).nStiffeningProfilePartCount
            
            ReDim Preserve arrayBasePlateInfo(nPlateIndex).oStiffeningProfileParts(nStiffeningProfileCount)
            Set arrayBasePlateInfo(nPlateIndex).oStiffeningProfileParts(nStiffeningProfileCount - 1).oPart = oPart
            If Not oLeafSystem Is Nothing Then
               Set arrayBasePlateInfo(nPlateIndex).oStiffeningProfileParts(nStiffeningProfileCount - 1).oLeafSystem = oLeafSystem
               arrayBasePlateInfo(nPlateIndex).oStiffeningProfileParts(nStiffeningProfileCount - 1).bIsSystem = True
            Else
               arrayBasePlateInfo(nPlateIndex).oStiffeningProfileParts(nStiffeningProfileCount - 1).bIsSystem = False
            End If
         End If
      End If
   Next
   
   Exit Sub

ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   GetConnectedPlateParts
' Descriptions:
'   For each base plate, get connected plate parts,including hull part is required
'
Public Sub GetConnectedPlateParts( _
        ByVal oObject As IJAssembly, _
        ByVal oPlateParts As IJElements, _
        ByVal nPlateCount As Long, _
        ByVal nPlateOption As Long, _
        ByRef arrayBasePlateInfo() As PLATE_INFO)
   Const sMethod As String = "GetConnectedPlateParts"
   On Error GoTo ErrorHandler
   
   Dim nPlateIndex As Long
   Dim stPlateInfo As PLATE_INFO
   Dim nPartIndex As Long
   Dim oPartToCheck As Object
   Dim bCheckIfConnected As Boolean
   
   Dim bIncludeHull As Boolean
   Dim bIncludeNonHull As Boolean
      
   If (nPlateOption And SEQUENCE_FILTER_CONNECTED_PLATE_NON_HULL) > 0 Then
      bIncludeNonHull = True
   Else
      bIncludeNonHull = False
   End If
   
   If (nPlateOption And SEQUENCE_FILTER_CONNECTED_PLATE_HULL) > 0 Then
      bIncludeHull = True
   Else
      bIncludeHull = False
   End If
   
   Dim bIsHullPart As Boolean
   Dim nIndex As Long
   Dim oPlate As IJPlate
   
   Dim oStructGraph As IJStructGraph
   Dim oSDHelper As New StructDetailHelper
   Dim oConnectable As IJConnectable
   
   Dim oPartToCheckRootSystem As IJSystem
   Dim bIsConnected As Boolean
   Dim oConnections As IJElements
   Dim nConnectedCount As Long
   Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim oAssemblyChild As IJAssemblyChild
   
   For nPlateIndex = 0 To nPlateCount - 1
      stPlateInfo = arrayBasePlateInfo(nPlateIndex)

      ' Get all plate parts that are not from current plate
      For nPartIndex = 1 To oPlateParts.Count
         Set oPartToCheck = oPlateParts.Item(nPartIndex)
         
         bCheckIfConnected = True
         bIsHullPart = False
         
         Set oAssemblyChild = oPartToCheck
         If Not oAssemblyChild Is Nothing Then
            If Not oObject Is oAssemblyChild.Parent Then
               bCheckIfConnected = False
            End If
         End If

         If bCheckIfConnected = True Then
            ' Check if part is from current plate
            For nIndex = 0 To stPlateInfo.nPlatePartCount - 1
               If oPartToCheck Is stPlateInfo.oPlateParts(nIndex).oPart Then
                  bCheckIfConnected = False
                  Exit For
               End If
            Next
   
            ' Check if it's hull part
            Set oPlate = oPartToCheck
            If oPlate.plateType = Hull Then
               bIsHullPart = True
               If bIncludeHull = False Then
                  bCheckIfConnected = False
               End If
            End If
         End If
         Set oPlate = Nothing

         bIsConnected = False
         If bCheckIfConnected = True Then
            Set oStructGraph = oPartToCheck
            oSDHelper.IsPartDerivedFromSystem oStructGraph, oPartToCheckRootSystem, True
            
            If stPlateInfo.bIsSystem = True Then
               ' Base plate is a plate system, check if it's connected to oParttoCheck's system
               If Not oPartToCheckRootSystem Is Nothing Then
                  Set oConnectable = stPlateInfo.oPlate
                  oConnectable.isConnectedTo oPartToCheckRootSystem, bIsConnected, oConnections
               Else
                  'oPartToCheck is a standalone part, check if it's bounded by base plate
                  Dim oBoundaries As Collection
                  Dim stBoundaryData As BoundaryData
                  Dim nBoundaryIndex As Long
               
                  Set oPlatePartWrapper.Object = oPartToCheck
                  Set oBoundaries = oPlatePartWrapper.PlateBoundaries()
                  For nBoundaryIndex = 1 To oBoundaries.Count
                     stBoundaryData = oBoundaries.Item(nBoundaryIndex)
                     
                     ' Check base plate's each part
                     For nIndex = 0 To stPlateInfo.nPlatePartCount - 1
                        If stPlateInfo.oPlateParts(nIndex).oPart Is stBoundaryData.Boundary Then
                           bIsConnected = True
                           Exit For
                        End If
                     Next
                     
                     If bIsConnected = True Then
                        Exit For
                     End If
                  Next
               End If
            Else
               ' Base plate is a stand alone plate
               If oPartToCheckRootSystem Is Nothing Then
                  'oPartToCheck is also a stand alone part
                  Set oConnectable = stPlateInfo.oPlate
                  oConnectable.isConnectedTo oPartToCheck, bIsConnected, oConnections
               Else
                  ' oPartToCheck is a system derived part. Check if based plate is bounded by oPartToCheck's system
                  Set oPlatePartWrapper.Object = stPlateInfo.oPlate
                  Set oBoundaries = oPlatePartWrapper.PlateBoundaries()
                  For nBoundaryIndex = 1 To oBoundaries.Count
                     stBoundaryData = oBoundaries.Item(nBoundaryIndex)
                     If oPartToCheck Is stBoundaryData.Boundary Then
                        bIsConnected = True
                        Exit For
                     End If
                  Next
               End If
            End If
         End If
         
         If bIsConnected = True Then
            If bIsHullPart = False Then
               If bIncludeNonHull = True Then
                  arrayBasePlateInfo(nPlateIndex).nConnectedPlatePartCount = arrayBasePlateInfo(nPlateIndex).nConnectedPlatePartCount + 1
                  nConnectedCount = arrayBasePlateInfo(nPlateIndex).nConnectedPlatePartCount
   
                  ReDim Preserve arrayBasePlateInfo(nPlateIndex).oConnectedPlateParts(nConnectedCount)
                  Set arrayBasePlateInfo(nPlateIndex).oConnectedPlateParts(nConnectedCount - 1).oPart = oPartToCheck
                  Set arrayBasePlateInfo(nPlateIndex).oConnectedPlateParts(nConnectedCount - 1).oRootSystem = oPartToCheckRootSystem
               End If
            Else
               If bIncludeHull = True Then
                  arrayBasePlateInfo(nPlateIndex).nConnectedHullPartCount = arrayBasePlateInfo(nPlateIndex).nConnectedHullPartCount + 1
                  nConnectedCount = arrayBasePlateInfo(nPlateIndex).nConnectedHullPartCount
   
                  ReDim Preserve arrayBasePlateInfo(nPlateIndex).oConnectedHullParts(nConnectedCount)
                  Set arrayBasePlateInfo(nPlateIndex).oConnectedHullParts(nConnectedCount - 1).oPart = oPartToCheck
                  Set arrayBasePlateInfo(nPlateIndex).oConnectedHullParts(nConnectedCount - 1).oRootSystem = oPartToCheckRootSystem
               End If
            End If
         End If
      Next
   Next

   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   GetConnectedProfileParts
' Descriptions:
'   For each base plate, get connected profile parts
'
Public Sub GetConnectedProfileParts( _
        ByVal oObject As IJAssembly, _
        ByVal oProfileParts As IJElements, _
        ByVal nPlateCount As Long, _
        ByRef arrayBasePlateInfo() As PLATE_INFO)
   Const sMethod As String = "GetConnectedProfileParts"
   On Error GoTo ErrorHandler
   
   Dim nPlateIndex As Long
   Dim stPlateInfo As PLATE_INFO
   Dim nPartIndex As Long
   Dim oPart As Object
   Dim bCheckIfConnected As Boolean
   Dim nIndex As Long
   Dim oStructGraph As IJStructGraph
   Dim oSDHelper As New StructDetailHelper
   Dim oRootSystem As IJSystem
   Dim oLeafSystem As IJSystem
   
   Dim oConnectable As IJConnectable
   Dim oStructConnectable As IJStructConnectable
   
   Dim bIsConnected As Boolean
   Dim oConnections As IJElements
   Dim nConnectedCount As Long
   Dim oAssemblyChild As IJAssemblyChild
   Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
      
   For nPlateIndex = 0 To nPlateCount - 1
      stPlateInfo = arrayBasePlateInfo(nPlateIndex)
      
      ' Get all profile parts but stiffening part that are connected to current plate
      For nPartIndex = 1 To oProfileParts.Count
         Set oPart = oProfileParts.Item(nPartIndex)
         
         bCheckIfConnected = True
         Set oAssemblyChild = oPart
         If Not oAssemblyChild Is Nothing Then
            If Not oObject Is oAssemblyChild.Parent Then
               bCheckIfConnected = False
            End If
         End If
         
         ' Check if the part is stiffening the plate
         If bCheckIfConnected = True Then
            For nIndex = 0 To stPlateInfo.nStiffeningProfilePartCount - 1
               If oPart Is stPlateInfo.oStiffeningProfileParts(nIndex).oPart Then
                  bCheckIfConnected = False
               End If
            Next
         End If
         
         If bCheckIfConnected = True Then
            ' Check if part is connected to plate
            bIsConnected = False
            If stPlateInfo.bIsSystem = True Then
               Set oStructGraph = oPart
               oSDHelper.IsPartDerivedFromSystem oStructGraph, oRootSystem, True
               If Not oRootSystem Is Nothing Then
                  ' Check at root system level
                  Set oConnectable = oRootSystem
                  oConnectable.isConnectedTo stPlateInfo.oPlate, bIsConnected, oConnections
               Else
                  Dim nPlatePartIndex As Long
                  
                  ' Check with each plate part
                  Set oConnectable = oPart
                  For nPlatePartIndex = 0 To stPlateInfo.nPlatePartCount - 1
                     oConnectable.isConnectedTo stPlateInfo.oPlateParts(nPlatePartIndex).oPart, bIsConnected, oConnections
                  Next
                  
                  ' Check if profile part is bounded by plate
                  Dim oBoundaries As Collection
                  Dim stBoundaryData As BoundaryData
                  Dim nBoundaryIndex As Long
                  
                  Set oProfilePartWrapper.Object = oPart
                  Set oBoundaries = oProfilePartWrapper.ProfileBoundaries()
                  For nBoundaryIndex = 1 To oBoundaries.Count
                     stBoundaryData = oBoundaries.Item(nBoundaryIndex)
                     If stPlateInfo.oPlate Is stBoundaryData.Boundary Then
                        bIsConnected = True
                        Exit For
                     End If
                  Next
               End If
            Else
               ' For stand alone plate, only stand alone profile can be bounded by it
               Set oConnectable = oPart
               oConnectable.isConnectedTo stPlateInfo.oPlate, bIsConnected, oConnections
            End If
            
            If bIsConnected = True Then
               nIndex = arrayBasePlateInfo(nPlateIndex).nConnectedProfilePartCount
               arrayBasePlateInfo(nPlateIndex).nConnectedProfilePartCount = arrayBasePlateInfo(nPlateIndex).nConnectedProfilePartCount + 1
               nConnectedCount = arrayBasePlateInfo(nPlateIndex).nConnectedProfilePartCount

               ReDim Preserve arrayBasePlateInfo(nPlateIndex).oConnectedProfileParts(nConnectedCount)
               Set arrayBasePlateInfo(nPlateIndex).oConnectedProfileParts(nConnectedCount - 1).oPart = oPart
               If Not oRootSystem Is Nothing Then
                  arrayBasePlateInfo(nPlateIndex).oConnectedProfileParts(nConnectedCount - 1).bIsSystem = True
                  oSDHelper.IsPartDerivedFromSystem oStructGraph, oLeafSystem, False
                  Set arrayBasePlateInfo(nPlateIndex).oConnectedProfileParts(nConnectedCount - 1).oRootSystem = oRootSystem
                  Set arrayBasePlateInfo(nPlateIndex).oConnectedProfileParts(nConnectedCount - 1).oLeafSystem = oLeafSystem
               Else
                  arrayBasePlateInfo(nPlateIndex).oConnectedProfileParts(nConnectedCount - 1).bIsSystem = False
               End If
            End If
         End If
      Next
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
End Sub

' Method:
'   GetAllTheParts
' Descriptions:
'   Get all the plate and profile parts in an assembly
'
Public Sub GetAllTheParts( _
               ByVal oObject, _
               ByVal oElements As IJElements, _
               ByRef oPlateParts As IJElements, _
               ByRef oProfileParts As IJElements)
   Const sMethod As String = "GetAllTheParts"
   On Error GoTo ErrorHandler

   If oObject Is Nothing And oElements Is Nothing Then
      Exit Sub
   End If
   
   Dim oPartsToProcess As IJElements
   
   If Not oElements Is Nothing Then
      Set oPartsToProcess = oElements
      If oPartsToProcess.Count = 0 Then
         Set oPartsToProcess = Nothing
      End If
   End If
   
   If oPartsToProcess Is Nothing Then
      If Not TypeOf oObject Is IJAssembly Then
         MsgBox "An assembly should be selected."
         Exit Sub
      End If
      
      Dim oAssembly As IJAssembly
      
      Set oAssembly = oObject
      Set oPartsToProcess = GetAssemblyChildren(oAssembly)
   Else
      Dim oAssemblyChild As IJAssemblyChild
      
      Set oAssemblyChild = oPartsToProcess.Item(1)
      Set oAssembly = oAssemblyChild.Parent
   End If

   If oAssembly Is Nothing Then
      MsgBox "No assembly or parts were selected!"
      Exit Sub
   End If
      
   ' Get all the plate parts and profile parts in the assembly
   Dim nPartIndex As Long
   Dim oPart As Object
   
   Set oPlateParts = New JObjectCollection
   Set oProfileParts = New JObjectCollection
   For nPartIndex = 1 To oPartsToProcess.Count
      Set oPart = oPartsToProcess.Item(nPartIndex)
      If TypeOf oPart Is IJPlate Then
         oPlateParts.Add oPart
      ElseIf TypeOf oPart Is IJProfilePart Then
         oProfileParts.Add oPart
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   GetDistanceToPlateCenterForAllParts
' Descriptions:
'   Calculate distance to base plate center
'
Public Sub GetDistanceToPlateCenterForAllParts(ByRef stPlateInfo As PLATE_INFO)
   Const sMethod As String = "GetDistanceToPlateCenterForAllParts"
   On Error GoTo ErrorHandler
   
   ' Get center of deck
   Dim oGraphicRange As IJRangeAlias
   Dim dLowX As Double
   Dim dLowY As Double
   Dim dLowZ As Double
   
   Dim dHiX As Double
   Dim dHiY As Double
   Dim dHiZ As Double
   Dim nIndex As Long
   Dim oBox  As GBox
   
   If stPlateInfo.bIsSystem = True Then
      Dim dMinX As Double
      Dim dMinY As Double
      Dim dMinZ As Double
      
      Dim dMaxX As Double
      Dim dMaxY As Double
      Dim dMaxZ As Double
      
      dMinX = 1000
      dMinY = 1000
      dMinZ = 1000
      
      dMaxX = -1000
      dMaxY = -1000
      dMaxZ = -1000
      
      For nIndex = 0 To stPlateInfo.nPlatePartCount - 1
         Set oGraphicRange = stPlateInfo.oPlateParts(nIndex).oPart
         oBox = oGraphicRange.GetRange
        
         dLowX = oBox.m_low.x
         dLowY = oBox.m_low.y
         dLowZ = oBox.m_low.z
         dHiX = oBox.m_high.x
         dHiY = oBox.m_high.y
         dHiZ = oBox.m_high.z
         
         If dLowX < dMinX Then
            dMinX = dLowX
         End If
         
         If dLowY < dMinY Then
            dMinY = dLowY
         End If
         
         If dLowZ < dMinZ Then
            dMinZ = dLowZ
         End If
         
         If dHiX > dMaxX Then
            dMaxX = dHiX
         End If
         
         If dHiY > dMaxY Then
            dMaxY = dHiY
         End If
            
         If dHiZ > dMaxZ Then
            dMaxZ = dHiZ
         End If
      Next
      
      dLowX = dMinX
      dLowY = dMinY
      dLowZ = dMinZ
      
      dHiX = dMaxX
      dHiY = dMaxY
      dHiZ = dMaxZ
   Else
   
      Set oGraphicRange = stPlateInfo.oPlate
      oBox = oGraphicRange.GetRange
        
      dLowX = oBox.m_low.x
      dLowY = oBox.m_low.y
      dLowZ = oBox.m_low.z
      dHiX = oBox.m_high.x
      dHiY = oBox.m_high.y
      dHiZ = oBox.m_high.z
      
   End If
   
   Dim oSGOMBUtil As New SGOModelBodyUtilities
   Dim oCenterPoint As IJDPosition
   Dim oClosestPoint As IJDPosition
   Dim dDistance As Double
   Dim oPartMB As IJDModelBody
   
   Set oCenterPoint = New DPosition
   oCenterPoint.Set (dLowX + dHiX) / 2, (dLowY + dHiY) / 2, (dLowZ + dHiZ) / 2
   
   For nIndex = 0 To stPlateInfo.nPlatePartCount - 1
      Set oPartMB = stPlateInfo.oPlateParts(nIndex).oPart
      oSGOMBUtil.GetClosestPointOnBody oPartMB, oCenterPoint, oClosestPoint, dDistance
      stPlateInfo.oPlateParts(nIndex).dDistToPlateCenter = dDistance
   Next
   
   For nIndex = 0 To stPlateInfo.nConnectedHullPartCount - 1
      Set oPartMB = stPlateInfo.oConnectedHullParts(nIndex).oPart
      oSGOMBUtil.GetClosestPointOnBody oPartMB, oCenterPoint, oClosestPoint, dDistance
      stPlateInfo.oConnectedHullParts(nIndex).dDistToPlateCenter = dDistance
   Next

   For nIndex = 0 To stPlateInfo.nConnectedPlatePartCount - 1
      Set oPartMB = stPlateInfo.oConnectedPlateParts(nIndex).oPart
      oSGOMBUtil.GetClosestPointOnBody oPartMB, oCenterPoint, oClosestPoint, dDistance
      stPlateInfo.oConnectedPlateParts(nIndex).dDistToPlateCenter = dDistance
   Next

   For nIndex = 0 To stPlateInfo.nConnectedProfilePartCount - 1
      Set oPartMB = stPlateInfo.oConnectedProfileParts(nIndex).oPart
      oSGOMBUtil.GetClosestPointOnBody oPartMB, oCenterPoint, oClosestPoint, dDistance
      stPlateInfo.oConnectedProfileParts(nIndex).dDistToPlateCenter = dDistance
   Next

   For nIndex = 0 To stPlateInfo.nStiffeningProfilePartCount - 1
      Set oPartMB = stPlateInfo.oStiffeningProfileParts(nIndex).oPart
      oSGOMBUtil.GetClosestPointOnBody oPartMB, oCenterPoint, oClosestPoint, dDistance
      stPlateInfo.oStiffeningProfileParts(nIndex).dDistToPlateCenter = dDistance
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   GetDirectionForAllParts
' Descriptions:
'   Get plate part's normal,profile part's direction(From start to end)
'
Public Sub GetDirectionForAllParts(ByRef stPlateInfo As PLATE_INFO, _
                                   ByVal oViewMatrix As IJDT4x4)
   Const sMethod As String = "GetDirectionForAllParts"
   On Error GoTo ErrorHandler
                                      
   Dim nPartIndex As Long
   Dim oWireBody As IJWireBody
   Dim oStartPos As IJDPosition
   Dim oEndPos As IJDPosition
   Dim oDirection As IJDVector
   Dim stPartInfo As PART_INFO
   Dim oProfileUtil As IJProfileAttributes
   
   Set oProfileUtil = New ProfileUtils
   For nPartIndex = 0 To stPlateInfo.nStiffeningProfilePartCount - 1
      stPartInfo = stPlateInfo.oStiffeningProfileParts(nPartIndex)
      If stPartInfo.bIsSystem Then
         Set oWireBody = stPlateInfo.oStiffeningProfileParts(nPartIndex).oLeafSystem
      Else
         oProfileUtil.GetLandingCurveFromProfile stPartInfo.oPart, oWireBody
      End If
      
      oWireBody.GetEndPoints oStartPos, oEndPos
      Set oDirection = oEndPos.Subtract(oStartPos)
      oDirection.Length = 1
      Set stPlateInfo.oStiffeningProfileParts(nPartIndex).oDirection = oDirection
   Next
   
   Set oWireBody = Nothing
   Set oStartPos = Nothing
   Set oEndPos = Nothing
   Set oDirection = Nothing
   
   Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
   Dim oPlateBasePort As IJPort
   Dim oPlatePortGeom As Object
   Dim oTopolocate As IJTopologyLocate
   Dim oDummyPoint As IJDPosition
   Dim oPlateNormal As IJDVector
   Dim oBasePlateNormal As IJDVector
   
   Set oBasePlateNormal = New DVector
   oBasePlateNormal.Set oViewMatrix.IndexValue(8), oViewMatrix.IndexValue(9), oViewMatrix.IndexValue(10)
   Set oTopolocate = New GSCADStructGeomUtilities.TopologyLocate
   For nPartIndex = 0 To stPlateInfo.nConnectedPlatePartCount - 1
      Set oPlatePartWrapper.Object = stPlateInfo.oConnectedPlateParts(nPartIndex).oPart
      Set oPlateBasePort = oPlatePartWrapper.BasePort(BPT_Base)
      Set oPlatePortGeom = oPlateBasePort.Geometry
      Set oPlateBasePort = Nothing
      oTopolocate.FindApproxCenterAndNormal oPlatePortGeom, oDummyPoint, oPlateNormal
      Set stPlateInfo.oConnectedPlateParts(nPartIndex).oDirection = oBasePlateNormal.Cross(oPlateNormal)
      Set oPlatePortGeom = Nothing
      Set oDummyPoint = Nothing
   Next
   
   For nPartIndex = 0 To stPlateInfo.nConnectedHullPartCount - 1
      Set oPlatePartWrapper.Object = stPlateInfo.oConnectedHullParts(nPartIndex).oPart
      Set oPlateBasePort = oPlatePartWrapper.BasePort(BPT_Base)
      Set oPlatePortGeom = oPlateBasePort.Geometry
      Set oPlateBasePort = Nothing
      oTopolocate.FindApproxCenterAndNormal oPlatePortGeom, oDummyPoint, oPlateNormal
      Set stPlateInfo.oConnectedHullParts(nPartIndex).oDirection = oBasePlateNormal.Cross(oPlateNormal)
      Set oPlatePortGeom = Nothing
      Set oDummyPoint = Nothing
   Next
   
   Set oBasePlateNormal = Nothing
   Set oPlatePartWrapper = Nothing
   Set oTopolocate = Nothing
      
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   UnsequenceBasePlateAndItsStiffeners
' Descriptions:
'   Unsequence all the parts related to a base plate
'
Public Sub UnsequenceBasePlateAndItsStiffeners( _
                     ByVal oObject As Object, _
                     ByRef stPlateInfo As PLATE_INFO)
   Const sMethod As String = "UnsequenceBasePlateAndItsStiffeners"
   On Error GoTo ErrorHandler
      
   Dim oAssemblySequence As IJAssemblySequence
   Dim nIndex As Long
   
   Set oAssemblySequence = oObject

   For nIndex = 0 To stPlateInfo.nPlatePartCount - 1
      oAssemblySequence.UnsequenceChild stPlateInfo.oPlateParts(nIndex).oPart
   Next
   
   For nIndex = 0 To stPlateInfo.nConnectedHullPartCount - 1
      oAssemblySequence.UnsequenceChild stPlateInfo.oConnectedHullParts(nIndex).oPart
   Next

   For nIndex = 0 To stPlateInfo.nConnectedPlatePartCount - 1
      oAssemblySequence.UnsequenceChild stPlateInfo.oConnectedPlateParts(nIndex).oPart
   Next

   For nIndex = 0 To stPlateInfo.nConnectedProfilePartCount - 1
      oAssemblySequence.UnsequenceChild stPlateInfo.oConnectedProfileParts(nIndex).oPart
   Next

   For nIndex = 0 To stPlateInfo.nStiffeningProfilePartCount - 1
      oAssemblySequence.UnsequenceChild stPlateInfo.oStiffeningProfileParts(nIndex).oPart
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise

End Sub

' Method:
'   AssignSequenceNumber
' Descriptions:
'   Assign sequence number to parts
'
' Note: nNextSequence is the sequence number to start with
'       When the method returns,it's updated to a number for next call to use
'
Public Sub AssignSequenceNumber( _
                        ByVal oAssemblySequence As IJAssemblySequence, _
                        ByVal nPartCount As Long, _
                        ByRef oArrayOfPart() As PART_DISTANCE_INFO, _
                        ByRef nNextSequence As Long)
   Const sMethod As String = "AssignSequenceNumber"
   On Error GoTo ErrorHandler
   
   Dim nIndex As Long
   Dim oPart As Object
   Dim oNI As IJNamedItem
   
   Dim sMsg As String
   Dim strFileName As String
   Dim nFileNumber As Integer
   
   If m_bLogToFile = True Then
      strFileName = "c:\temp\Sequence.txt"
      nFileNumber = FreeFile
      Open strFileName For Append As nFileNumber
   End If
   
   For nIndex = 0 To nPartCount - 1
      Set oPart = oArrayOfPart(nIndex).oPart
      If nNextSequence = -1 Then
         nNextSequence = 1
      Else
         nNextSequence = nNextSequence + 1
      End If
      
      If m_bLogToFile = True Then
         Set oNI = oPart
         If Not oNI Is Nothing Then
            sMsg = "Index: " & nNextSequence & " -- PartName: " & oNI.Name
            
            Print #nFileNumber, sMsg
         Else
            MsgBox "oNI is nothing"
         End If
         Set oNI = Nothing
      End If
      
      oAssemblySequence.ChildIndex(oPart) = nNextSequence
      Set oPart = Nothing
   Next
   
   If m_bLogToFile = True Then
      Print #nFileNumber, vbNewLine
      Close #nFileNumber
   End If
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
End Sub

' Method:
'   SortByDistance
' Descriptions:
'   Sort part by it's distance to base plate center
'
Public Sub SortByDistance(arrayPartToSequence() As PART_DISTANCE_INFO)
   Const sMethod As String = "SortByDistance"
   On Error GoTo ErrorHandler
   
   Dim nCount As Long
   Dim nOuterIndex As Long
   Dim nInnerIndex As Long
   Dim nToSwapIndex As Long
   Dim stTempPartInfo As PART_DISTANCE_INFO
   Dim dMinDistance As Double
   
   nCount = UBound(arrayPartToSequence)
   For nOuterIndex = 0 To nCount - 1
      nToSwapIndex = nOuterIndex
      dMinDistance = arrayPartToSequence(nOuterIndex).dDistToPlateCenter
      For nInnerIndex = nOuterIndex + 1 To nCount - 1
         If dMinDistance > arrayPartToSequence(nInnerIndex).dDistToPlateCenter Then
            dMinDistance = arrayPartToSequence(nInnerIndex).dDistToPlateCenter
            nToSwapIndex = nInnerIndex
         End If
      Next
      
      If nToSwapIndex <> nOuterIndex Then
         ' Need to swap
         ' Move nOuterIndex content to temporary storage
         Set stTempPartInfo.oPart = arrayPartToSequence(nOuterIndex).oPart
         stTempPartInfo.dDistToPlateCenter = arrayPartToSequence(nOuterIndex).dDistToPlateCenter
         
         ' Move nToSwapIndex content to nOuterIndex
         Set arrayPartToSequence(nOuterIndex).oPart = arrayPartToSequence(nToSwapIndex).oPart
         arrayPartToSequence(nOuterIndex).dDistToPlateCenter = arrayPartToSequence(nToSwapIndex).dDistToPlateCenter
         
         ' Move stored content to nToSwapIndex
         Set arrayPartToSequence(nToSwapIndex).oPart = stTempPartInfo.oPart
         arrayPartToSequence(nToSwapIndex).dDistToPlateCenter = stTempPartInfo.dDistToPlateCenter
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise

End Sub

' Method:
'   OutputBasePlateInfo
' Descriptions:
'   Output base plate info to a file for debugging purpose
'
Public Sub OutputBasePlateInfo( _
                     nPlateCount As Long, _
                     arrayBasePlateInfo() As PLATE_INFO)
   Const sMethod As String = "OutputBasePlateInfo"
   On Error GoTo ErrorHandler
      
   Dim strFileName As String
   Dim nFileNumber As Integer
   Dim nPlateIndex As Long
   Dim nIndex As Long
   
   strFileName = "c:\temp\BasePlateInfo.txt"
   nFileNumber = FreeFile
   Open strFileName For Append As nFileNumber

   Print #nFileNumber, "BasePlateCount = " & nPlateCount
   If nPlateCount = 0 Then
      Close #nFileNumber
      Exit Sub
   End If
   
   nPlateCount = UBound(arrayBasePlateInfo)
   For nPlateIndex = 0 To nPlateCount - 1
      Dim stPlateInfo As PLATE_INFO
      Dim oNI As IJNamedItem
      Dim stPartInfo As PART_INFO
      
      stPlateInfo = arrayBasePlateInfo(nPlateIndex)
      Set oNI = stPlateInfo.oPlate

      Print #nFileNumber, "BasePlate: " & oNI.Name
      
      Print #nFileNumber, "   PlatePartCount: " & stPlateInfo.nPlatePartCount
      For nIndex = 0 To stPlateInfo.nPlatePartCount - 1
         stPartInfo = stPlateInfo.oPlateParts(nIndex)
         
         Set oNI = stPartInfo.oPart
         Print #nFileNumber, "      PlatePart: " & oNI.Name
         
         If Not stPartInfo.oDirection Is Nothing Then
            Print #nFileNumber, "         Direction:(" & stPartInfo.oDirection.X & ", " & _
                                                         stPartInfo.oDirection.Y & ", " & _
                                                         stPartInfo.oDirection.Z & ")"
            
         Else
            Print #nFileNumber, "         Direction is nothing"
         End If
         Print #nFileNumber, "         Distance: " & stPartInfo.dDistToPlateCenter
      Next
      Print #nFileNumber, vbNewLine
      
      Print #nFileNumber, "   StiffeningCount: " & stPlateInfo.nStiffeningProfilePartCount
      For nIndex = 0 To stPlateInfo.nStiffeningProfilePartCount - 1
         Set oNI = stPlateInfo.oStiffeningProfileParts(nIndex).oPart
         Print #nFileNumber, "      StiffeningPart: " & oNI.Name
            
         stPartInfo = stPlateInfo.oStiffeningProfileParts(nIndex)
         If Not stPartInfo.oDirection Is Nothing Then
            Print #nFileNumber, "         Direction:(" & stPartInfo.oDirection.X & ", " & _
                                                         stPartInfo.oDirection.Y & ", " & _
                                                         stPartInfo.oDirection.Z & ")"
         Else
            Print #nFileNumber, "         Direction is nothing"
         End If
         Print #nFileNumber, "         Distance: " & stPartInfo.dDistToPlateCenter
      Next
      Print #nFileNumber, vbNewLine
      
      Print #nFileNumber, "   ConnectedProfilePartCout: " & stPlateInfo.nConnectedProfilePartCount
      For nIndex = 0 To stPlateInfo.nConnectedProfilePartCount - 1
         Set oNI = stPlateInfo.oConnectedProfileParts(nIndex).oPart
         Print #nFileNumber, "      ConnectedProfilePart: " & oNI.Name
         
         stPartInfo = stPlateInfo.oConnectedProfileParts(nIndex)
         If Not stPartInfo.oDirection Is Nothing Then
            Print #nFileNumber, "         Direction:(" & stPartInfo.oDirection.X & ", " & _
                                                         stPartInfo.oDirection.Y & ", " & _
                                                         stPartInfo.oDirection.Z & ")"
         Else
            Print #nFileNumber, "         Direction: is nothing"
         End If
         Print #nFileNumber, "         Distance: " & stPartInfo.dDistToPlateCenter
      Next
      Print #nFileNumber, vbNewLine
      
      Print #nFileNumber, "   ConnectedPlatePartCount: " & stPlateInfo.nConnectedPlatePartCount
      For nIndex = 0 To stPlateInfo.nConnectedPlatePartCount - 1
         Set oNI = stPlateInfo.oConnectedPlateParts(nIndex).oPart
         Print #nFileNumber, "      ConnectedPlatePart: " & oNI.Name
         
         stPartInfo = stPlateInfo.oConnectedPlateParts(nIndex)
         If Not stPartInfo.oDirection Is Nothing Then
            Print #nFileNumber, "         Direction:(" & stPartInfo.oDirection.X & ", " & _
                                                         stPartInfo.oDirection.Y & ", " & _
                                                         stPartInfo.oDirection.Z & ")"
         Else
            Print #nFileNumber, "         Direction: is nothing"
         End If
         Print #nFileNumber, "         Distance: " & stPartInfo.dDistToPlateCenter
      Next
      Print #nFileNumber, vbNewLine

      Print #nFileNumber, "   ConnectedHullPartCount: " & stPlateInfo.nConnectedHullPartCount
      For nIndex = 0 To stPlateInfo.nConnectedHullPartCount - 1
         Set oNI = stPlateInfo.oConnectedHullParts(nIndex).oPart
         Print #nFileNumber, "      ConnectedHullPart: " & oNI.Name
         
         stPartInfo = stPlateInfo.oConnectedHullParts(nIndex)
         If Not stPartInfo.oDirection Is Nothing Then
         
            Print #nFileNumber, "         Direction:(" & stPartInfo.oDirection.X & ", " & _
                                                         stPartInfo.oDirection.Y & ", " & _
                                                         stPartInfo.oDirection.Z & ")"
         Else
            Print #nFileNumber, "         Direction: is nothing"
         End If
         Print #nFileNumber, "         Distance: " & stPartInfo.dDistToPlateCenter
                                                            
      Next
      Print #nFileNumber, vbNewLine
      
      Print #nFileNumber, vbNewLine
   
      Set oNI = Nothing
   Next

   Close #nFileNumber
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   GetPartsDirectionInfo
' Descriptions:
'   Determine if part is along base plate U direction or V direction
'   Return the number of parts for each direction
'
Public Sub GetPartsDirectionInfo( _
                     ByVal nPartCount As Long, _
                     ByRef oArrayPartInfo() As PART_INFO, _
                     ByVal oBasePlateU As IJDVector, _
                     ByVal oBasePlateV As IJDVector, _
                     ByRef nUDirectionCount As Long, _
                     ByRef nVDirectionCount As Long, _
                     ByVal nFilter As Long)
   Const sMethod As String = "GetPartsDirectionInfo"
   On Error GoTo ErrorHandler

   Dim nPartIndex As Long
   Dim oStructContinuity As IJStructContinuity
   Dim stPartInfo As PART_INFO
   Dim bCheckDirection As Boolean
   
   For nPartIndex = 0 To nPartCount - 1
      stPartInfo = oArrayPartInfo(nPartIndex)
      If nFilter = GET_DIRECTION_FILTER_NONE Then
         bCheckDirection = True
         
      ElseIf nFilter = GET_DIRECTION_FILTER_CONTINUOUS Then
         If Not stPartInfo.oLeafSystem Is Nothing Then
            Set oStructContinuity = stPartInfo.oLeafSystem
         ElseIf Not stPartInfo.oRootSystem Is Nothing Then
            Set oStructContinuity = stPartInfo.oRootSystem
         End If
         
         If Not oStructContinuity Is Nothing Then
            If oStructContinuity.ContinuityType = ContinuousType Then
               bCheckDirection = True
            Else
               bCheckDirection = False
            End If
         Else
            ' Does not support IJStructContinuity,can not be intercoastal,treat it as continuous
            bCheckDirection = True
         End If
      Else
         ' Any other filter,set it to true for now
         bCheckDirection = True
      End If
      
      If bCheckDirection = True Then
         If Abs(oBasePlateU.Dot(stPartInfo.oDirection)) > Abs(oBasePlateV.Dot(stPartInfo.oDirection)) Then
            oArrayPartInfo(nPartIndex).bAlongU = True
            oArrayPartInfo(nPartIndex).bAlongV = False
            oArrayPartInfo(nPartIndex).bMainDirection = False
            oArrayPartInfo(nPartIndex).bPerpendicularToMain = False
            
            nUDirectionCount = nUDirectionCount + 1
         Else
            oArrayPartInfo(nPartIndex).bAlongU = False
            oArrayPartInfo(nPartIndex).bAlongV = True
            oArrayPartInfo(nPartIndex).bMainDirection = False
            oArrayPartInfo(nPartIndex).bPerpendicularToMain = False
            
            nVDirectionCount = nVDirectionCount + 1
         End If
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   SetPartsDirectionFlags
' Descriptions:
'   Determine if part is along main direction or perpendicular to main direction
'   Return the number of parts for each direction
'
Public Sub SetPartsDirectionFlags( _
                              ByVal nPartCount As Long, _
                              ByRef oArrayPartInfo() As PART_INFO, _
                              ByVal nUContinuousCount As Long, _
                              ByVal nVContinuousCount As Long, _
                              ByVal oBasePlateU As IJDVector, _
                              ByVal oBasePlateV As IJDVector, _
                              ByRef nMainDirectionCount As Long, _
                              ByRef nPerpendicularCount As Long, _
                              ByRef nOtherCount As Long)
   Const sMethod As String = "SetPartsDirectionFlags"
   On Error GoTo ErrorHandler
   
   Dim nPartIndex As Long
   Dim stPartInfo As PART_INFO
   
   For nPartIndex = 0 To nPartCount - 1
      stPartInfo = oArrayPartInfo(nPartIndex)
      If nUContinuousCount >= nVContinuousCount Then
         If stPartInfo.bAlongU = True Then
            oArrayPartInfo(nPartIndex).bMainDirection = True
            nMainDirectionCount = nMainDirectionCount + 1
         ElseIf stPartInfo.bAlongV = True Then
            oArrayPartInfo(nPartIndex).bMainDirection = False
            If (Abs(oBasePlateV.Dot(stPartInfo.oDirection) - 1)) < 0.01 Then
               oArrayPartInfo(nPartIndex).bPerpendicularToMain = True
               nPerpendicularCount = nPerpendicularCount + 1
            Else
               oArrayPartInfo(nPartIndex).bPerpendicularToMain = False
               nOtherCount = nOtherCount + 1
            End If
         Else
            nOtherCount = nOtherCount + 1
         End If
      Else
         If stPartInfo.bAlongU = True Then
            oArrayPartInfo(nPartIndex).bMainDirection = False
            If (Abs(oBasePlateU.Dot(stPartInfo.oDirection) - 1)) < 0.01 Then
               oArrayPartInfo(nPartIndex).bPerpendicularToMain = True
               nPerpendicularCount = nPerpendicularCount + 1
            Else
               oArrayPartInfo(nPartIndex).bPerpendicularToMain = False
               nOtherCount = nOtherCount + 1
            End If
         ElseIf stPartInfo.bAlongV = True Then
            oArrayPartInfo(nPartIndex).bMainDirection = True
            nMainDirectionCount = nMainDirectionCount + 1
         Else
            nOtherCount = nOtherCount + 1
         End If
      End If
   Next

   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   GetGroupedPartsInfo
' Descriptions:
'   Get parts that are along main direction or perpendicular to main direction
'   Return the number of each group
'
Public Sub GetGroupedPartsInfo( _
                       ByRef nPartCount As Long, _
                       ByRef oArrayPartToGroup() As PART_INFO, _
                       ByRef oArrayMainDirectionPart() As PART_DISTANCE_INFO, _
                       ByRef oArrayPerpendicularPart() As PART_DISTANCE_INFO, _
                       ByRef oArrayOtherPart() As PART_DISTANCE_INFO, _
                       ByRef nMainDirectionCount As Long, _
                       ByRef nPerpendicularCount As Long, _
                       ByRef nOtherCount As Long)
   Const sMethod As String = "GetGroupedPartsInfo"
   On Error GoTo ErrorHandler

   Dim nPartIndex As Long
   Dim stPartInfo As PART_INFO
   
   For nPartIndex = 0 To nPartCount - 1
      stPartInfo = oArrayPartToGroup(nPartIndex)
      If stPartInfo.bMainDirection = True Then
         Set oArrayMainDirectionPart(nMainDirectionCount).oPart = stPartInfo.oPart
         oArrayMainDirectionPart(nMainDirectionCount).dDistToPlateCenter = stPartInfo.dDistToPlateCenter
         nMainDirectionCount = nMainDirectionCount + 1
         
      ElseIf stPartInfo.bPerpendicularToMain = True Then
         Set oArrayPerpendicularPart(nPerpendicularCount).oPart = stPartInfo.oPart
         oArrayPerpendicularPart(nPerpendicularCount).dDistToPlateCenter = stPartInfo.dDistToPlateCenter
         nPerpendicularCount = nPerpendicularCount + 1
         
      Else
         Set oArrayOtherPart(nOtherCount).oPart = stPartInfo.oPart
         oArrayOtherPart(nOtherCount).dDistToPlateCenter = stPartInfo.dDistToPlateCenter
         nOtherCount = nOtherCount + 1
      End If
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise
   
End Sub

' Method:
'   CollectBasePlateInfo
' Descriptions:
'   Get all the base plates
'   For each base palte,get any of following if required:
'     Collect base plate parts
'     Stiffening profile parts
'     Connected non-hull plate parts
'     Connected hull plate parts
'     Connected profile parts
'  Unsequence parts for each base plate
'
Public Sub CollectBasePlateInfo( _
                    ByVal oObject As Object, _
                    ByVal oElements As IJElements, _
                    ByVal nSequenceOption As Long, _
                    ByRef arrayBasePlateInfo() As PLATE_INFO, _
                    ByRef nBasePlateCount As Long, _
                    ByRef oViewMatrix As IJDT4x4, _
                    ByRef nUnrelatedPartCount As Long, _
                    ByRef oUnrelatedPart() As Object, _
                    Optional oElem As Object)
   Const sMethod As String = "CollectBasePlateInfo"
   On Error GoTo ErrorHandler

   Dim bIncludeHullInBasePlate As Boolean
      
   Dim oAllPlateParts As IJElements
   Dim oAllProfileParts As IJElements
   
   ' Get all the parts in the assembly
   GetAllTheParts oObject, _
                  oElements, _
                  oAllPlateParts, _
                  oAllProfileParts

   ' Get base plates
   Dim nBasePlateOption As Long
   
   nBasePlateOption = (nSequenceOption And SEQUENCE_FILTER_BASEPLATE_NON_HULL) Or _
                      (nSequenceOption And SEQUENCE_FILTER_BASEPLATE_HULL)
   If Not nBasePlateOption > 0 Then
      nBasePlateOption = SEQUENCE_FILTER_BASEPLATE_NON_HULL
   End If
   GetBasePlates oObject, _
                 oAllPlateParts, _
                 nBasePlateCount, _
                 arrayBasePlateInfo, _
                 oViewMatrix, _
                 nBasePlateOption

   If nBasePlateCount > 0 Then
   
      ' Get all stiffening profiles on each base plate
      If (nSequenceOption And SEQUENCE_FILTER_STIFFENING_PROFILE) > 0 Then
         GetStiffeningProfileParts oObject, _
                                   oAllProfileParts, _
                                   nBasePlateCount, _
                                   arrayBasePlateInfo
      End If
      
      ' Get connected parts for each base plate
      ' Connected Plates
      Dim nConnectedPlateOption As Long
      
      nConnectedPlateOption = (nSequenceOption And SEQUENCE_FILTER_CONNECTED_PLATE_NON_HULL) Or _
                              (nSequenceOption And SEQUENCE_FILTER_CONNECTED_PLATE_HULL)
      If nConnectedPlateOption > 0 Then
         GetConnectedPlateParts oObject, _
                                oAllPlateParts, _
                                nBasePlateCount, _
                                nConnectedPlateOption, _
                                arrayBasePlateInfo
      End If
      
      ' Connected profiles
      If (nSequenceOption And SEQUENCE_FILTER_CONNECTED_PROFILE) > 0 Then
         GetConnectedProfileParts oObject, _
                                  oAllProfileParts, _
                                  nBasePlateCount, _
                                  arrayBasePlateInfo
      End If
   End If
   
   ' For each part related to plate, find it's distance to the center of the plate and its direction
   Dim nBasePlateIndex As Long
   
   For nBasePlateIndex = 0 To nBasePlateCount - 1
      GetDistanceToPlateCenterForAllParts arrayBasePlateInfo(nBasePlateIndex)
      GetDirectionForAllParts arrayBasePlateInfo(nBasePlateIndex), oViewMatrix
   Next
   
   ' Output plates info into file
   If m_bLogToFile = True Then
      OutputBasePlateInfo nBasePlateCount, arrayBasePlateInfo
   End If

   GetUnrelatedParts oAllPlateParts, _
                     oAllProfileParts, _
                     nBasePlateCount, _
                     arrayBasePlateInfo, _
                     nUnrelatedPartCount, _
                     oUnrelatedPart
   
   ' Unsequence all the parts
   '    Parts related to base plate
   For nBasePlateIndex = 0 To nBasePlateCount - 1
      UnsequenceBasePlateAndItsStiffeners oObject, arrayBasePlateInfo(nBasePlateIndex)
   Next
   
   '    All other parts
   Dim nUnrelatedPartIndex As Long
   Dim oAssemblySequence As IJAssemblySequence
   
   Set oAssemblySequence = oObject
   For nUnrelatedPartIndex = 0 To nUnrelatedPartCount - 1
      oAssemblySequence.UnsequenceChild oUnrelatedPart(nUnrelatedPartIndex)
   Next
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise

End Sub

' Method:
'   GetUnrelatedParts
' Descriptions:
'   Get all the parts that are not related to any base plate
'
Public Sub GetUnrelatedParts( _
                     oAllPlateParts As IJElements, _
                     oAllProfileParts As IJElements, _
                     nBasePlateCount As Long, _
                     arrayBasePlateInfo() As PLATE_INFO, _
                     nUnrelatedPartCount As Long, _
                     oUnrelatedPart() As Object)
                     
   Const sMethod As String = "GetUnrelatedParts"
   On Error GoTo ErrorHandler

   Dim nOuterIndex As Long
   Dim nBasePlateIndex As Long
   Dim stPlateInfo As PLATE_INFO
   Dim nPartIndex As Long
   Dim bFound As Boolean
   Dim oPart As Object
   Dim nPlateIndex As Long
   
   If Not oAllPlateParts Is Nothing Then
      For nOuterIndex = 1 To oAllPlateParts.Count
         bFound = False
         Set oPart = oAllPlateParts.Item(nOuterIndex)
         For nPlateIndex = 0 To nBasePlateCount - 1
            stPlateInfo = arrayBasePlateInfo(nPlateIndex)
            
            For nPartIndex = 0 To stPlateInfo.nPlatePartCount - 1
               If oPart Is stPlateInfo.oPlateParts(nPartIndex).oPart Then
                  bFound = True
                  Exit For
               End If
            Next
            If bFound = True Then
               Exit For
            End If
            
            For nPartIndex = 0 To stPlateInfo.nConnectedHullPartCount - 1
               If oPart Is stPlateInfo.oConnectedHullParts(nPartIndex).oPart Then
                  bFound = True
               End If
            Next
            If bFound = True Then
               Exit For
            End If
         
            For nPartIndex = 0 To stPlateInfo.nConnectedPlatePartCount - 1
               If oPart Is stPlateInfo.oConnectedPlateParts(nPartIndex).oPart Then
                  bFound = True
               End If
            Next
            If bFound = True Then
               Exit For
            End If
         Next
         
         If bFound = False Then
            nUnrelatedPartCount = nUnrelatedPartCount + 1
            ReDim Preserve oUnrelatedPart(nUnrelatedPartCount)
            Set oUnrelatedPart(nUnrelatedPartCount - 1) = oPart
         End If
      Next
   End If
   
   If Not oAllProfileParts Is Nothing Then
      For nOuterIndex = 1 To oAllProfileParts.Count
         bFound = False
         Set oPart = oAllProfileParts.Item(nOuterIndex)
         
         For nPlateIndex = 0 To nBasePlateCount - 1
            stPlateInfo = arrayBasePlateInfo(nPlateIndex)
         
            For nPartIndex = 0 To stPlateInfo.nConnectedProfilePartCount - 1
               If oPart Is stPlateInfo.oConnectedProfileParts(nPartIndex).oPart Then
                  bFound = True
                  Exit For
               End If
            Next
            If bFound = True Then
               Exit For
            End If
            
            For nPartIndex = 0 To stPlateInfo.nStiffeningProfilePartCount - 1
               If oPart Is stPlateInfo.oStiffeningProfileParts(nPartIndex).oPart Then
                  bFound = True
                  Exit For
               End If
            Next
            If bFound = True Then
               Exit For
            End If
         Next
         
         If bFound = False Then
            nUnrelatedPartCount = nUnrelatedPartCount + 1
            ReDim Preserve oUnrelatedPart(nUnrelatedPartCount)
            Set oUnrelatedPart(nUnrelatedPartCount - 1) = oPart
         End If
      Next
   End If
   
   Exit Sub
   
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & sMethod)
   m_oError.Raise

End Sub


