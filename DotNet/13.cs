using System;

class Program
{
    static void Main(string[] args)
    {
        int minutes;

        Console.Write("Unesi broj minuta: ");
        minutes = int.Parse(Console.ReadLine());

        int seconds = minutes * 60;

        Console.WriteLine("{0} minuta je {1} sekundi", minutes, seconds);
        Console.ReadLine();
    }
}
