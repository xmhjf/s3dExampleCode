VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeqBasePlate_Hull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ***************************************************************************
' Copyright (c) 2009, Intergraph Corporation.  All rights reserved.
'
' Project
'   SMSequenceRules
'
' File
'   SeqBasePlate_Hull.cls
'
' Example rule for base plates and connected parts

' ***************************************************************************

Option Explicit

Implements IJPlnSequenceRule
Private Const Module As String = "SeqBasePlate_Hull"

Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

'
' When this method is called, oObject is assembly object,
' oElements is objects selected in grid,could be empty
Private Sub IJPlnSequenceRule_DefSeqRule( _
                  ByVal oObject As Object, _
                  ByVal oElements As GSCADPlnIntHelper.IJElements, _
                  Optional ByVal oElem As Object = Nothing)

   Dim arrayBasePlateInfo() As PLATE_INFO
   Dim nSequenceFilter As Long
   Dim nBasePlateCount As Long
   Dim oViewMatrix As IJDT4x4
   Dim nUnrelatedPartCount As Long
   Dim oUnrelatedPart() As Object

   ' Specify parts that will be sequenced
   nSequenceFilter = SEQUENCE_FILTER_BASEPLATE_NON_HULL Or _
                     SEQUENCE_FILTER_BASEPLATE_HULL Or _
                     SEQUENCE_FILTER_STIFFENING_PROFILE Or _
                     SEQUENCE_FILTER_CONNECTED_PLATE_NON_HULL Or _
                     SEQUENCE_FILTER_CONNECTED_PLATE_HULL Or _
                     SEQUENCE_FILTER_CONNECTED_PROFILE
                     
   nUnrelatedPartCount = 0
   CollectBasePlateInfo _
                       oObject, _
                       oElements, _
                       nSequenceFilter, _
                       arrayBasePlateInfo, _
                       nBasePlateCount, _
                       oViewMatrix, _
                       nUnrelatedPartCount, _
                       oUnrelatedPart, _
                       oElem
                       
   ' Sequence all plate parts and profile parts for each deck
   Dim nBasePlateIndex As Long
   
   For nBasePlateIndex = 0 To nBasePlateCount - 1
      SequenceInnerBottomIncludingHull _
                      oObject, _
                      arrayBasePlateInfo(nBasePlateIndex), _
                      oViewMatrix
   Next
   
   ' Handle parts that are not related to any base plate here!
   '
End Sub

' Sequence order:
'  BasePlate parts
'    Connected parts
'      Continuous parts(Connected plate part, stiffening profile part, hull part)
'      Intersocal parts
'      Other parts
'
Private Sub SequenceInnerBottomIncludingHull( _
                                 oObject As Object, _
                                 stPlateInfo As PLATE_INFO, _
                                 oViewMatrix As IJDT4x4)
                                 
   If oObject Is Nothing Or oViewMatrix Is Nothing Then
      Exit Sub
   End If
   
   ' Find continuous direction
   Dim oBasePlateU As IJDVector
   Dim oBasePlateV As IJDVector
   
   Set oBasePlateU = New DVector
   Set oBasePlateV = New DVector
   oBasePlateU.Set oViewMatrix.IndexValue(0), oViewMatrix.IndexValue(1), oViewMatrix.IndexValue(2)
   oBasePlateV.Set oViewMatrix.IndexValue(4), oViewMatrix.IndexValue(5), oViewMatrix.IndexValue(6)

   Dim stPartInfo As PART_INFO
   Dim nUContinuousCount As Long
   Dim nVContinuousCount As Long
   Dim nPartIndex As Long
   Dim oStructContinuity As IJStructContinuity
   
   ' Count number of parts along base plate U and V direction for:
   '   Stiffening profile part
   '   Connected non-hull plate part
   '   Connected hull part
   
   nUContinuousCount = 0
   nVContinuousCount = 0
   
   GetPartsDirectionInfo _
                     stPlateInfo.nStiffeningProfilePartCount, _
                     stPlateInfo.oStiffeningProfileParts, _
                     oBasePlateU, _
                     oBasePlateV, _
                     nUContinuousCount, _
                     nVContinuousCount, _
                     GET_DIRECTION_FILTER_CONTINUOUS
   
   GetPartsDirectionInfo _
                     stPlateInfo.nConnectedPlatePartCount, _
                     stPlateInfo.oConnectedPlateParts, _
                     oBasePlateU, _
                     oBasePlateV, _
                     nUContinuousCount, _
                     nVContinuousCount, _
                     GET_DIRECTION_FILTER_CONTINUOUS
   
   GetPartsDirectionInfo _
                     stPlateInfo.nConnectedHullPartCount, _
                     stPlateInfo.oConnectedHullParts, _
                     oBasePlateU, _
                     oBasePlateV, _
                     nUContinuousCount, _
                     nVContinuousCount, _
                     GET_DIRECTION_FILTER_CONTINUOUS
   
   ' Set main/perpendicular flag for parts
   Dim nMainDirectionCount As Long
   Dim nPerpendicularCount As Long
   Dim nOtherCount As Long
   
   nMainDirectionCount = 0
   nPerpendicularCount = 0
   nOtherCount = 0
       
   SetPartsDirectionFlags stPlateInfo.nConnectedPlatePartCount, _
                          stPlateInfo.oConnectedPlateParts, _
                          nUContinuousCount, _
                          nVContinuousCount, _
                          oBasePlateU, _
                          oBasePlateV, _
                          nMainDirectionCount, _
                          nPerpendicularCount, _
                          nOtherCount

   SetPartsDirectionFlags stPlateInfo.nConnectedHullPartCount, _
                          stPlateInfo.oConnectedHullParts, _
                          nUContinuousCount, _
                          nVContinuousCount, _
                          oBasePlateU, _
                          oBasePlateV, _
                          nMainDirectionCount, _
                          nPerpendicularCount, _
                          nOtherCount
             
   SetPartsDirectionFlags stPlateInfo.nStiffeningProfilePartCount, _
                          stPlateInfo.oStiffeningProfileParts, _
                          nUContinuousCount, _
                          nVContinuousCount, _
                          oBasePlateU, _
                          oBasePlateV, _
                          nMainDirectionCount, _
                          nPerpendicularCount, _
                          nOtherCount
   
   SetPartsDirectionFlags stPlateInfo.nConnectedProfilePartCount, _
                          stPlateInfo.oConnectedProfileParts, _
                          nUContinuousCount, _
                          nVContinuousCount, _
                          oBasePlateU, _
                          oBasePlateV, _
                          nMainDirectionCount, _
                          nPerpendicularCount, _
                          nOtherCount
   
   Dim oArrayPlatePart() As PART_DISTANCE_INFO
   Dim oArrayMainDirectionPart() As PART_DISTANCE_INFO
   Dim oArrayPerpendicularPart() As PART_DISTANCE_INFO
   Dim oArrayOtherPart() As PART_DISTANCE_INFO
   
   ReDim oArrayPlatePart(stPlateInfo.nPlatePartCount)
   ReDim oArrayMainDirectionPart(nMainDirectionCount)
   ReDim oArrayPerpendicularPart(nPerpendicularCount)
   ReDim oArrayOtherPart(nOtherCount)
   
   nMainDirectionCount = 0
   nPerpendicularCount = 0
   nOtherCount = 0
   
   For nPartIndex = 0 To stPlateInfo.nPlatePartCount - 1
      stPartInfo = stPlateInfo.oPlateParts(nPartIndex)
      Set oArrayPlatePart(nPartIndex).oPart = stPartInfo.oPart
      oArrayPlatePart(nPartIndex).dDistToPlateCenter = stPartInfo.dDistToPlateCenter
   Next
   
   GetGroupedPartsInfo stPlateInfo.nConnectedPlatePartCount, _
                       stPlateInfo.oConnectedPlateParts, _
                       oArrayMainDirectionPart, _
                       oArrayPerpendicularPart, _
                       oArrayOtherPart, _
                       nMainDirectionCount, _
                       nPerpendicularCount, _
                       nOtherCount
      
   GetGroupedPartsInfo stPlateInfo.nConnectedHullPartCount, _
                       stPlateInfo.oConnectedHullParts, _
                       oArrayMainDirectionPart, _
                       oArrayPerpendicularPart, _
                       oArrayOtherPart, _
                       nMainDirectionCount, _
                       nPerpendicularCount, _
                       nOtherCount
   
   GetGroupedPartsInfo stPlateInfo.nStiffeningProfilePartCount, _
                       stPlateInfo.oStiffeningProfileParts, _
                       oArrayMainDirectionPart, _
                       oArrayPerpendicularPart, _
                       oArrayOtherPart, _
                       nMainDirectionCount, _
                       nPerpendicularCount, _
                       nOtherCount
   
   GetGroupedPartsInfo stPlateInfo.nConnectedProfilePartCount, _
                       stPlateInfo.oConnectedProfileParts, _
                       oArrayMainDirectionPart, _
                       oArrayPerpendicularPart, _
                       oArrayOtherPart, _
                       nMainDirectionCount, _
                       nPerpendicularCount, _
                       nOtherCount
   
   If stPlateInfo.nPlatePartCount > 0 Then
      SortByDistance oArrayPlatePart
   End If
   
   If nMainDirectionCount > 0 Then
      SortByDistance oArrayMainDirectionPart
   End If
   
   If nPerpendicularCount > 0 Then
      SortByDistance oArrayPerpendicularPart
   End If
   
   If nOtherCount > 0 Then
      SortByDistance oArrayOtherPart
   End If
   
   Dim nMaxIndex As Long
   Dim oAssemblySequence As IJAssemblySequence
   
   nMaxIndex = 0
   nMaxIndex = MaxSeqIndex(oObject)
      
   Set oAssemblySequence = oObject
   If stPlateInfo.nPlatePartCount > 0 Then
      AssignSequenceNumber oAssemblySequence, _
                           stPlateInfo.nPlatePartCount, _
                           oArrayPlatePart, _
                           nMaxIndex
   End If
   
   If nMainDirectionCount > 0 Then
      AssignSequenceNumber oAssemblySequence, _
                           nMainDirectionCount, _
                           oArrayMainDirectionPart, _
                           nMaxIndex
   End If
   
   If nPerpendicularCount > 0 Then
      AssignSequenceNumber oAssemblySequence, _
                           nPerpendicularCount, _
                           oArrayPerpendicularPart, _
                           nMaxIndex
   End If
   If nOtherCount > 0 Then
      AssignSequenceNumber oAssemblySequence, _
                           nOtherCount, _
                           oArrayOtherPart, _
                           nMaxIndex
   End If
   
End Sub


