Attribute VB_Name = "WefFrameGlobal"
Option Explicit

Private m_WebFrameLocation As Long ' to be passed to PointAtMinDistEx ( 1 inside, 2 outside )

Property Get WebFrameLocation() As Long
    If m_WebFrameLocation <> 1 And m_WebFrameLocation <> 2 Then m_WebFrameLocation = 1
    WebFrameLocation = m_WebFrameLocation
End Property

Property Let WebFrameLocation(oValue As Long)
    m_WebFrameLocation = oValue
End Property

