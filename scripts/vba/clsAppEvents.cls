Option Explicit

' Application events: keep the inline HTML in step with the slide show.

Public WithEvents App As Application

Private Sub App_SlideShowBegin(ByVal Wn As SlideShowWindow)
    On Error Resume Next
    If Wn.Presentation.Tags(TAG_DECK) = "1" Then InlineShowBegan
End Sub

Private Sub App_SlideShowNextSlide(ByVal Wn As SlideShowWindow)
    On Error Resume Next
    If Wn.Presentation.Tags(TAG_DECK) = "1" Then InlineSync "next"
End Sub

Private Sub App_SlideShowEnd(ByVal Pres As Presentation)
    On Error Resume Next
    If Pres.Tags(TAG_DECK) = "1" Then InlineShowEnded "event"
End Sub

Private Sub App_PresentationClose(ByVal Pres As Presentation)
    On Error Resume Next
    If Pres.Tags(TAG_DECK) = "1" Then InlineShowEnded "close"
End Sub
