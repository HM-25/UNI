using System;  
public class vjezba
{  
    public static void Main()
{
   int i,j,r;
   
//nacrtaj dijamant     
   
   Console.Write("unesi broj redova za dijamant :");
   r = Convert.ToInt32(Console.ReadLine());   
   for(i=0;i<=r;i++)
   {
     for(j=1;j<=r-i;j++)
     Console.Write(" ");
     for(j=1;j<=2*i-1;j++)
     Console.Write("*");
     Console.Write("\n");
   }
 
   for(i=r-1;i>=1;i--)
   {
     for(j=1;j<=r-i;j++)
     Console.Write(" ");
     for(j=1;j<=2*i-1;j++)
       Console.Write("*");
     Console.Write("\n");
   }
  }
}