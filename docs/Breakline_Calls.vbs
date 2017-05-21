' Applying Breakline _ for Wide func/sub calls
Sub ManyArgs(A, B, C, D, E)
  MsgBox A&B&C&D&E
End Sub


Call ManyArgs( "111111111 sadaskd", _
    "22222222 sadaskd",_
    _
    "33333333 sadaskd",_
    _
    "44444444 sadaskd",_
    _
    "55555555 sadaskd")

Call ManyArgs(  "111111111 sadaskd",_
                "22222222 sadaskd",_
                "33333333 sadaskd",_
                "44444444 sadaskd",_
                "55555555 sadaskd" )
