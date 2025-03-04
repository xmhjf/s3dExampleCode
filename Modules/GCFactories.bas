Attribute VB_Name = "GCFactories"
Option Explicit
Private m_pGCGeomFactory As IJGCGeomFactory
Private m_pGCGeomFactory2 As IJGCGeomFactory2
Public Function GetGCGeomFactory() As IJGCGeomFactory
    If m_pGCGeomFactory Is Nothing Then
        Set m_pGCGeomFactory = CreateObject("GCCMNSTRDefinitions.GCGeomFactory")
    End If
    Set GetGCGeomFactory = m_pGCGeomFactory
End Function
Public Function GetGCGeomFactory2() As IJGCGeomFactory2
    If m_pGCGeomFactory2 Is Nothing Then
        Set m_pGCGeomFactory2 = CreateObject("GCCMNSTRDefinitions.GCGeomFactory")
    End If
    Set GetGCGeomFactory2 = m_pGCGeomFactory2
End Function

