using System;
//provjeri da li string pocinje sa C#
namespace Vjezba18
{
class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(test("C# Sharp"));
            Console.WriteLine(test("C#"));
            Console.WriteLine(test("C++"));
            Console.WriteLine(test("C# test dodatni"));
            Console.ReadLine();
        }
       public static string test(string str)
    {
        return (str.Length < 3 && str.Equals("C#")) || (str.StartsWith("C#") && str[2] == ' ') ? "Pocinje sa C#" : "Ne pocinje sa C#";

        }
    }
}