using System;

class Kalkulator
{
    public static double saberi(double a, double b)
    {
        return a + b;
    }

    public static double oduzmi(double a, double b)
    {
        return a - b;
    }

    public static double pomnozi(double a, double b)
    {
        return a * b;
    }

    public static double podijeli(double a, double b)
    {
        if (b == 0)
        {
            throw new DivideByZeroException();
        }
        return a / b;
    }
}
class Program
{
    static void Main(string[] args)
    {
        double num1, num2;
        char operation;

        Console.WriteLine("Unesi prvi broj: ");
        num1 = double.Parse(Console.ReadLine());

        Console.WriteLine("unesi operandu (+,-,*,/) ");
        operation = char.Parse(Console.ReadLine());

        Console.WriteLine("Unesi drugi broj: ");
        num2 = double.Parse(Console.ReadLine());

        switch (operation)
         {
            case '+':
                Console.WriteLine("Rezultat: " + Kalkulator.saberi(num1, num2));
                break;
            case '-':
                Console.WriteLine("Rezultat: " + Kalkulator.oduzmi(num1, num2));
                break;
            case '*':
                Console.WriteLine("Rezultat: " + Kalkulator.pomnozi(num1, num2));
                break;
            case '/':
                Console.WriteLine("Rezultat: " + Kalkulator.podijeli(num1, num2));
                break;
                
                default:
                Console.WriteLine("Greska");
                break;
        }

        Console.ReadLine();
    }
}
