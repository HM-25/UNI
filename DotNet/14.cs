using System;

class rekurzivnavjezba
{
    static int printNatural(int vrijednost1, int vrijednost2)
    {
	if (vrijednost2 < 1)
	{
	    return vrijednost1;
	}
	vrijednost2--;
	Console.Write(" {0} ",vrijednost1);
	return printNatural(vrijednost1 + 1, vrijednost2);
    }
    static void Main()
    {
//ispisi brojeve u nizu redosljedom
	Console.Write(" Koliko brojeva zelite isprintati : ");
	int vrijednost2= Convert.ToInt32(Console.ReadLine());
	// pozivanje rekurzivne funkcije sa 2 parametra	
	printNatural(1, vrijednost2);
	Console.Write("\n\n");
    }
}