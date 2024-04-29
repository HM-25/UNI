using System;
class testiranje
{
static void Main(string[] args) 
    {
    	Console.Write(" Suma brojeva: ");
	Console.Write(" Unesi koliko brojeva zelis sabrati : ");
	int n = Convert.ToInt32(Console.ReadLine());    
    Console.Write(" Suma prvih {0} brojeva je : {1}\n\n", n,SumOfTen(1,n));
    }

static int SumOfTen(int min, int max) 
    {
    return CalcuSum(min, max);
    }

static int CalcuSum(int min, int val) 
    {
    if (val == min)
        return val;
    return val + CalcuSum(min, val - 1);
    }
}