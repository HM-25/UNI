using System;

class ChristmasEveChecker
{
static bool IsChristmasEve(int year, int month, int day)
{
if (month == 11 && day == 24)
{
return true;
}
else
{
return false;
}
}

static void Main()
{
    Console.Write("unesi godinu: ");
    int year = int.Parse(Console.ReadLine());

    Console.Write("unesi mjesec: ");
    int month = int.Parse(Console.ReadLine());

    Console.Write("unesi dan: ");
    int day = int.Parse(Console.ReadLine());

    if (IsChristmasEve(year, month, day))
    {
        Console.WriteLine("Bozicno vece je. vrijeme je za kolacice i mlijeko.");
    }
    else
    {
        Console.WriteLine("jos nije Bozic. pricekaj jos malo.");
    }
}
  }