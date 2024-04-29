using System;

class AnimalFarm
{
static int TotalLegs(int chickenCount, int cowCount, int pigCount)
{
int chickenLegs = chickenCount * 2;
int cowLegs = cowCount * 4;
int pigLegs = pigCount * 4;
      return chickenLegs + cowLegs + pigLegs;
}

static void Main()
{
    Console.Write("unesi broj kokosi: ");
    int chickenCount = int.Parse(Console.ReadLine());

    Console.Write("unei broj krava: ");
    int cowCount = int.Parse(Console.ReadLine());

    Console.Write("unesi broj svinja: ");
    int pigCount = int.Parse(Console.ReadLine());

    int totalLegs = TotalLegs(chickenCount, cowCount, pigCount);

    Console.WriteLine("ukupno nogu: " + totalLegs);
}
}