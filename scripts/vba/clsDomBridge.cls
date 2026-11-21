Option Explicit

' Bridge object handed to a page: it becomes the onclick handler of the page's hidden #ihCmd
' element. The browser invokes the default member (DISPID 0) when the page "clicks" that element.

Public Page As Long

Public Function Fire(Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant) As Variant
Attribute Fire.VB_UserMemId = 0
    On Error Resume Next
    InlineBridge Page
    Fire = False
End Function
