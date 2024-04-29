using System;  
public class Exercise25  
{  
    public static void Main() 
   {
  int num1,num2,opt;

           Console.Write("\n\n");
           Console.Write("kalkulator (komplikovani):\n");
           Console.Write("\n\n");


  Console.Write("unesi prvi broj :");
  num1 = Convert.ToInt32(Console.ReadLine());
  Console.Write("unesi drugi broj :");
  num2 = Convert.ToInt32(Console.ReadLine());

  
    Console.Write("\nocpije :\n");
    Console.Write("1-zbir.\n2-oduzimanje.\n3.mnozenje.\ndijeljenje\n");
    Console.Write("\nInput your choice :");
    opt = Convert.ToInt32(Console.ReadLine());

    switch(opt) {
      case 1:
        Console.Write("zbir  {0} i {1} je: {2}\n",num1,num2,num1+num2);
        break;
        
      case 2:
        Console.Write("razlika {0} i {1} je: {2}\n",num1,num2,num1-num2);
        break;
        
      case 3:
        Console.Write("proizvod {0} i {1} je: {2}\n",num1,num2,num1*num2);
        break;  
      
      case 4:
        if(num2==0) {
          Console.Write("drugi broj je 0, dijeljenje sa nulom nije moguce.\n");
        } else {
          Console.Write("razlomak {0} i {1} je: {2}\n",num1,num2,num1/num2);
        }  
        break;
        
      case 5: 
        break; 
        
      default:
        Console.Write("GRESKA\n");
        break; 
    }
  }
}
