public bool MethodFuture_Value(double Present_Value, double Interest_Rate, double Years, double Future_Value)
{
//Variable: Present_Value = 1000 - Sheet1!E7 
//Variable: Interest_Rate = 0.05 - Sheet1!E8 
//Variable: Years = 20 - Sheet1!E9 
//RETURN: Future_Value = 2653.297705 - Future Value
return Present_Value*POWER(1+Interest_Rate, Years);

}

public bool test_MethodFuture_Value()
{
 return MethodFuture_Value(1000, 0.05, 20, 2653.297705);
}

