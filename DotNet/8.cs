using System;  
public class test
{  
    public static void Main()
{
   int monno;


           Console.Write("\n\n");
 Console.Write("provjeri naziv mjeseca po broju");
           Console.Write("\n\n");


         Console.Write("unesi broj mjeseca : ");
         monno = Convert.ToInt32(Console.ReadLine());

   switch(monno)
   {
	case 1:
	       Console.Write("Januar\n");
	       break;
	case 2:
	       Console.Write("Februar\n");
	       break;
	case 3:
	       Console.Write("Mart\n");
	       break;
	case 4:
	       Console.Write("April\n");
	       break;
	case 5:
	       Console.Write("Maj\n");
	       break;
	case 6:
	       Console.Write("Juni\n");
	       break;
	case 7:
	       Console.Write("Juli\n");
	       break;
	case 8:
	       Console.Write("August\n");
	       break;
	case 9:
	       Console.Write("Septembar\n");
	       break;
	case 10:
	       Console.Write("Oktobar\n");
	       break;
	case 11:
	       Console.Write("Novembar\n");
	       break;
	case 12:
	       Console.Write("Decembar\n");
	       break;
	default:
	       Console.Write("netacan broj, pokusajte opet ....\n");
	       break;
      }
   }
}
