using System;
//ispisi slova u suprotnom redosljedu
public class Vjezba20
{
  public static void Main()
  {  
    char slovo1,slovo2,slovo3;
  
    Console.Write("Unesi prvo slovo: ");
    slovo1 = Convert.ToChar(Console.ReadLine());
 
    Console.Write("Unesi drugo slovo: ");
    slovo2 = Convert.ToChar(Console.ReadLine());       
 
    Console.Write("Unesi trece slovo: ");
    slovo3 = Convert.ToChar(Console.ReadLine());
     
    Console.WriteLine("{0} {1} {2}",slovo3,slovo2,slovo1);
   }
}