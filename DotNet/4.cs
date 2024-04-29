using System;  
public class Exercise26  
{  
    public static void Main()
{
  int n,i,sum=0;
  int t=1;
    Console.Write("izracunaj sumu serije brojeva 1+11+111+1111.....:\n");
  
  
  Console.Write("unesi duzinu niza : ");
   n= Convert.ToInt32(Console.ReadLine());  
  for(i=1;i<=n;i++)
  {
     Console.Write("{0} + ",t);
     sum=sum+t;
     t=(t*10)+1;
  }
  Console.Write("\nsuma je : {0}\n",sum);
   }
}