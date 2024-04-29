using System;

class brojeviunazad
{
static int ispisiCijele(int var1, int var2)
{
    if (var1 <= 1)
    {
        Console.Write(" {0} ",var1);
        return var1;
    }

    Console.Write(" {0} ",var1);
    var1--;
    return ispisiCijele(var1, var2);
}


    static void Main()
    {
        //isprintaj niz brojeva unazad do 1

        Console.Write(" Od kojeg broja zelite krenuti : ");
        int var1= Convert.ToInt32(Console.ReadLine());
        // poziv rekurzivne funkcije sa 2 parametra.	
        ispisiCijele(var1,1);
        Console.Write("\n\n");
    }
}
