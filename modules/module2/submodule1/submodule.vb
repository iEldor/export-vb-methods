Private Sub Button1_Click(...)
 
  Dim oFoo As New IDLexFooClass()
   
  Try
  oFoo.CreateObject(0, 0, 0)
  Catch ex As Exception
  Debug.WriteLine(oFoo.GetLastError())
  Return
  End Try
   
  ' use object here... 
   
End Sub

Private Sub Button1_Click(...) 'checks blah blah
 
  Dim oFoo As New IDLexFooClass
   
  Dim parmStr As String = "I am a string parameter"
  Dim parmVal As Int32 = 24
  Dim parmArr As Int32(,) = {{10, 11, 12}, {20, 21, 22}}
   
  Dim argc As Int32 = 3
  Dim argval As Object() = {parmStr, parmVal, parmArr}
  Dim argpal As Int32() = {PARMFLAG_CONST, PARMFLAG_CONST, (PARMFLAG_CONST + PARMFLAG_CONV_MAJORITY)}
   
  Try
  oFoo.CreateObject(argc, argval, argpal)
  Catch ex As Exception
  Debug.WriteLine(oFoo.GetLastError())
  Return
  End Try
   
  ' use object here...
 
End Sub