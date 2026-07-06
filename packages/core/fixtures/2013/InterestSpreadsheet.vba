Function Future_Value(Present_Value, Interest_Rate, Years)
Future_Value = Present_Value*POWER(1+Interest_Rate, Years)
End Function

Sub DescribeFunction(FuncName As String, FuncDesc As String, Category As String, ArgDesc As Variant)
Application.MacroOptions _
Macro:=FuncName, _
Description:=FuncDesc, _
Category:=Category, _
ArgumentDescriptions:=ArgDesc
End Sub         
Function Power(Var1 , Var2)
Power = Var1 ^ Var2
End Function         
Function Pi()
Pi = 4 * Atn(1)
End Function         
Function inlineIf(statement, iftrue, Optional iffalse)
 
If (statement) Then 
inlineIf = iftrue 
ElseIf isMissing(iffalse) Then 
inlineIf = False 
Else: inlineIf = iffalse
Endif

End Function 

Function inlineAnd(b1, b2) 
inlineAnd = b1 And b2 
End Function 
Function ABSFun(num) 



ABSFun = Abs(num) 

End Function 

Function SQRTFun(num) 

SQRTFun = Sqr(num) 
End Function 

Function MIN(num, num2, Optional num3, Optional num4, Optional num5, Optional num6, Optional num7) 

MIN = num 

If (num2 < num) Then 
MIN = num2 
End If 


If (IsMissing(num3) = False) Then 
If (num3 < num) Then 
MIN = num3 
End If 
End If 

If (IsMissing(num4) = False) Then 
If (num4 < num) Then 
MIN = num4 
End If 
End If 

If (IsMissing(num5) = False) Then 
If (num5 < num) Then 
MIN = num5 
End If 
End If 

If (IsMissing(num6) = False) Then 
If (num6 < num) Then 
MIN = num6 
End If 
End If 

If (IsMissing(num7) = False) Then 
If (num7 < num) Then 
MIN = num7 
End If 
End If 

End Function              
Sub DefineFunctions()
Dim ArgDesc1(1 To 3) As String
ArgDesc1(1) = " (), "
ArgDesc1(2) = " (), "
ArgDesc1(3) = " (), "
DescribeFunction "Future_Value", "Future Value", "1", ArgDesc1

End Sub
