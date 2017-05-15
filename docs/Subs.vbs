' VBScript::Subroutines '
' 1) - Parameter passing '
' 2) - Scoping
'------------------------'


'------------------------------------------------------------------------------'

'' 1) - Parameter Passing
' - It is possible to pass by Value and by Reference via call syntax
' - Use 'Call' SubName or no surrounding parenthesis to call by reference
' - Use surrounding parenthesis with NO Call prefix to pass by value
' - Put byVal in Subroutine argument list to prevent by reference calls
Sub doubleIt(a)
	a=a*2
End Sub

Sub doubleItVal(byVal a) ' Forces by Value only calls'
	a=a*2
End Sub

Dim value : value = 4

MsgBox "[1]: Initial Value::" & value

' Sub Call with () without the Call keyword passes byValue '
doubleIt(value)
MsgBox "ByVal Value::" & value
' Value:4 - Unchanged'

' naked Sub Call pass byReference '
doubleIt value
MsgBox "ByRef Value::" & value
' Value:8 '

' Using Call passes by reference '
Call doubleIt(value)
MsgBox "Call ByRef Value::" & value
' Value:16 '

Call doubleItVal(value)
doubleItVal value
MsgBox "byVal defined Call ByRef Value::" & value
' Value:16 - Unchanged'

'------------------------------------------------------------------------------'

'' 2) - Scoping
' - All variables dimmed in global scope can be accessed/changed within subs '
' - Variables dimmed in sub context can be accessed/changed within subs where they are dimmed only '
' - Shadowing is allowed in subs to dim a variable that already exists to prevent interfering w. globally scopped var with same name

Dim gVar
		gVar = 5
Sub countA()
	For gVar = 4 to 1 Step -1
		' do 4 3 2 1
	Next
	MsgBox "end-countA Value::" & gVar
End Sub

Sub countB()
	Dim gVar	' shadow scoping; global gVar is now unaccessable '
	For gVar = 4 to 1 Step -1
		' do 4 3 2 1
	Next
	MsgBox "end-countB Value::" & gVar
End Sub

MsgBox "[2]: Initial Global Value::" & gVar

countA()
MsgBox "Global Value::" & gVar
' 0'
		gVar = 5
		MsgBox "ResetGlobal Value::" & gVar
countB()
MsgBox "Global Value::" & gVar
' 5'
