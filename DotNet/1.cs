using System;
using System.Linq;

class vjezba
{        
    static void Main()
    {
        //array-niz
        int[] n1 = new int[10] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        
  

        // query
        var nQuery =
            from VrNum in n1
            where (VrNum % 2) == 0
            select VrNum;

        // The third part is Query execution.
        
        Console.Write("\nbrojevi koji imaju ostatak 0 nakon sto su podijeljeni sa 2 su : \n");
        foreach (int VrNum in nQuery)
        {
            Console.Write("{0} ", VrNum);
        }
         Console.Write("\n\n");
    }
}