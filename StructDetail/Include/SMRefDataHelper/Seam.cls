VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Seam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULE As String = "S:\StructDetail\Data\Include\SMRefDataHelper\Seam.cls"

'Method IfSeamsExistWithInDistance --- 'This Function Return the distance between the seam and pLocation, if there is more than one seam,
'it will return the distance which is  between the furthest seam and pLocation

'Inputs: 1) root object of a plate or a profile is to be passed
'        2) location from which seams are looked for
'        3) Specified distance

'Outputs: 1) Distance of seam(that lies with in distance) from the location(pLocation)
'         2)Direction of the vector that joins the location and its projected
            'point on seam landing curve

'oDirec: pass in the profile secondery oritation
'bolNeedSameSide: True: means you want the seam which is on the Web Right Side
'                 False: Means you want the seam which is on the Web Left Side

Public Function IfSeamsExistWithInDistance(pLocation As IJDPosition, oRootObject As Object, dWithInDistance As Double, _
                                                dSeamDistance As Double, Optional oDirec As IJDVector, _
                                                Optional bolNeedSameSide As Boolean = False, _
                                                Optional bSeam_Feature As Boolean) As Boolean
    IfSeamsExistWithInDistance = False
    bSeam_Feature = False
    
    Dim oJSeam As IJSeam
    Dim oWireBody As IUnknown
    Dim oLandingCurve As IJLandCurve
    Dim oSeamAttributes As IJSeamAttributes
    Dim oSeamUtils As SeamUtils
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Dim oVector As IJDVector
    Dim oPntOnWire As IJDPosition
    Dim dDistance As Double
    Dim oStructDetailHelper As New StructDetailObjects.Helper
    Dim oListOfAllSeams As Collection
    
    Dim bolDetermineSide As Boolean
    bolDetermineSide = False
    If Not oDirec Is Nothing Then
        bolDetermineSide = True
    End If
    
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    Set oSeamUtils = New SeamUtils
    Set oSeamAttributes = oSeamUtils
    
'    oStructDetailHelper.GetSeamsOnObject oRootObject, oListOfAllSeams
    Set oListOfAllSeams = New Collection
    GetSeamsOnPlate oRootObject, oListOfAllSeams
    
    dSeamDistance = 0
    
    Dim iIndex As Long
    Dim nSeams As Long
    nSeams = oListOfAllSeams.Count
    For iIndex = 1 To nSeams
        ' For Each Seam
        ' Retrieve it's Geometry (WireBody)
        Set oJSeam = oListOfAllSeams.Item(iIndex)
        oSeamAttributes.GetLandingCurveFromSeam oJSeam, oLandingCurve
         
        ' Retrieve the nearest Point on the Seam from the Penetration Point
        ' by Projecting the Penetration Point onto the Seam's WireBody
        Set oWireBody = oLandingCurve
        oTopologyLocate.GetProjectedPointOnModelBody oWireBody, pLocation, _
                                              oPntOnWire, oVector
                                              
        Dim oTempVector As IJDVector
        Set oTempVector = pLocation.Subtract(oPntOnWire)
                                           
        'Determine the seam is on which side of the location. if its a profile
        'it will be decided if it is on web right or web left.
        Dim bolSameSide As Boolean
        bolSameSide = False
        If bolDetermineSide = True Then
            Dim dFlag As Double
            dFlag = oDirec.x * oTempVector.x + oDirec.y * oTempVector.y + oDirec.z * oTempVector.z
            
            If dFlag > 0 Then
               bolSameSide = False
            Else
               bolSameSide = True
            End If
        End If
                                                                            
        If (Not oVector Is Nothing) And (bolNeedSameSide = bolSameSide) Then
            If oVector.Length > 0.0001 Then
                ' Calculate distance between the Point on Seam and Penetration Point
                dDistance = oPntOnWire.DistPt(pLocation)
                If dDistance <= dWithInDistance Then
                    IfSeamsExistWithInDistance = True
                    If dDistance > dSeamDistance Then
                        dSeamDistance = dDistance
                    End If
                End If
                If dDistance >= dWithInDistance And dDistance <= (dWithInDistance + 0.015) Then
                    bSeam_Feature = True
                    dSeamDistance = dDistance
                End If
            End If
        End If
        
        Set oVector = Nothing
        Set oPntOnWire = Nothing
        
        Set oJSeam = Nothing
        Set oWireBody = Nothing
        Set oLandingCurve = Nothing
    Next iIndex
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, "Seam.cls", "IfSeamsExistWithInDistance").Number
End Function


'GetSeamsOnPlate -- used to fetch all the seams on a given object.
'wrapper GetSeamsOnObject necessitates passing root object whereas this method can take
'input as any object. be it a root parent or a leaf plate.

Public Sub GetSeamsOnPlate(oSystemUn As IUnknown, oCollection As Collection)
    On Error Resume Next
    Dim oParent As IJSystem
    
    Dim lCount As Long
    Dim sChildren As String
    Dim oChildren As IJDTargetObjectCol
    
    If TypeOf oSystemUn Is IJProfile Then
        Exit Sub
    End If
    
    Set oParent = oSystemUn
    Set oChildren = oParent.GetChildren
    
    lCount = oChildren.Count
    Dim i As Integer
    For i = 1 To lCount
        Dim oObject As Object
        Set oObject = oChildren.Item(i)
        If TypeOf oObject Is IJSeam Then
            oCollection.Add oObject
        Else
            GetSeamsOnPlate oObject, oCollection
        End If
    Next

End Sub
