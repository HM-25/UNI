using System;
public class test
{
public static void Main()
{
int num,i,ctr=0;
  Console.Write("\n\n");
Console.Write("Provjeri da li je broj prost:\n");
Console.Write("\n\n");  	

Console.Write("Unesi broj: ");
num= Convert.ToInt32(Console.ReadLine());

for(i=2;i<=num/2;i++){
    if(num % i==0){
        ctr++;
        break; 
    }
}
if(ctr==0 && num!= 1) 
Console.Write("{0} je prost broj.\n",num);
else
Console.Write("{0} nije prost broj\n",num);
}
}