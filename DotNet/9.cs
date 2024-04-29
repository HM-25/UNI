using System;
using System.Linq;
namespace exercises
{
class Program
{
static void Main(string[] args)
{
Console.WriteLine("unesi string");
string str1 = Console.ReadLine();
string sortedString = test(str1);
Console.WriteLine("string po abecednom redu: " + sortedString);
}
public static string test(string str1)
{
return new string(str1.OrderBy(x => x).ToArray());
}
}
}