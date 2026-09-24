// Exercise 4: Inventory with file saving
//
// Console app for a small shop inventory. Products are stored in a
// Dictionary and saved to / loaded from a CSV text file.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

class Product
{
    public string Code;
    public string Name;
    public int Quantity;
    public decimal Price;

    public decimal TotalValue { get { return Quantity * Price; } }

    public string ToCsv()
    {
        return string.Join(";", Code, Name, Quantity, Price.ToString(CultureInfo.InvariantCulture));
    }

    public static Product FromCsv(string line)
    {
        string[] p = line.Split(';');
        return new Product
        {
            Code = p[0],
            Name = p[1],
            Quantity = int.Parse(p[2]),
            Price = decimal.Parse(p[3], CultureInfo.InvariantCulture)
        };
    }
}

class Inventory
{
    private readonly Dictionary<string, Product> products = new Dictionary<string, Product>();
    private readonly string file;

    public Inventory(string file)
    {
        this.file = file;
        if (File.Exists(file))
        {
            foreach (string line in File.ReadAllLines(file))
            {
                if (line.Trim().Length == 0) continue;
                Product p = Product.FromCsv(line);
                products[p.Code] = p;
            }
        }
    }

    public void Add(Product p)
    {
        if (products.ContainsKey(p.Code))
            products[p.Code].Quantity += p.Quantity;
        else
            products[p.Code] = p;
    }

    public void Sell(string code, int amount)
    {
        Product p;
        if (!products.TryGetValue(code, out p)) throw new KeyNotFoundException("Unknown product code");
        if (amount > p.Quantity) throw new InvalidOperationException("Not enough items in stock");
        p.Quantity -= amount;
    }

    public void Print()
    {
        decimal total = 0;
        Console.WriteLine("{0,-6} {1,-15} {2,5} {3,8}", "Code", "Name", "Qty", "Price");
        foreach (Product p in products.Values)
        {
            Console.WriteLine("{0,-6} {1,-15} {2,5} {3,8:F2}{4}", p.Code, p.Name, p.Quantity, p.Price, p.Quantity < 5 ? "  (low stock)" : "");
            total += p.TotalValue;
        }
        Console.WriteLine("Total inventory value: {0:F2}", total);
    }

    public void Save()
    {
        List<string> lines = new List<string>();
        foreach (Product p in products.Values) lines.Add(p.ToCsv());
        File.WriteAllLines(file, lines);
    }
}

class Program
{
    static void Main()
    {
        Inventory inv = new Inventory("inventory.csv");

        while (true)
        {
            Console.WriteLine("\n1-list  2-add  3-sell  0-save and exit");
            Console.Write("Choice: ");
            string choice = Console.ReadLine();
            if (choice == null) break;

            try
            {
                switch (choice.Trim())
                {
                    case "1":
                        inv.Print();
                        break;
                    case "2":
                        Product p = new Product();
                        Console.Write("Code: "); p.Code = Console.ReadLine();
                        Console.Write("Name: "); p.Name = Console.ReadLine();
                        Console.Write("Quantity: "); p.Quantity = int.Parse(Console.ReadLine());
                        Console.Write("Price: "); p.Price = decimal.Parse(Console.ReadLine(), CultureInfo.InvariantCulture);
                        inv.Add(p);
                        break;
                    case "3":
                        Console.Write("Code: "); string code = Console.ReadLine();
                        Console.Write("Amount: "); int amount = int.Parse(Console.ReadLine());
                        inv.Sell(code, amount);
                        Console.WriteLine("Sold.");
                        break;
                    case "0":
                        inv.Save();
                        Console.WriteLine("Saved to inventory.csv");
                        return;
                    default:
                        Console.WriteLine("Unknown option.");
                        break;
                }
            }
            catch (FormatException)
            {
                Console.WriteLine("Error: please enter a valid number.");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
        }
    }
}
