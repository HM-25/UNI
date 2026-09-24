// Exercise 1: Shapes (inheritance, abstract classes, polymorphism)
//
// An abstract Shape class with Area() and Perimeter(), and three
// concrete shapes. The program sorts a list of shapes by area.

using System;
using System.Collections.Generic;

abstract class Shape
{
    public string Name { get; protected set; }

    public abstract double Area();
    public abstract double Perimeter();

    public override string ToString()
    {
        return string.Format("{0,-10} area = {1,8:F2}   perimeter = {2,8:F2}", Name, Area(), Perimeter());
    }
}

class Circle : Shape
{
    private readonly double radius;

    public Circle(double radius)
    {
        if (radius <= 0) throw new ArgumentException("Radius must be positive");
        this.radius = radius;
        Name = "Circle";
    }

    public override double Area() { return Math.PI * radius * radius; }
    public override double Perimeter() { return 2 * Math.PI * radius; }
}

class Rectangle : Shape
{
    protected readonly double width;
    protected readonly double height;

    public Rectangle(double width, double height)
    {
        if (width <= 0 || height <= 0) throw new ArgumentException("Sides must be positive");
        this.width = width;
        this.height = height;
        Name = "Rectangle";
    }

    public override double Area() { return width * height; }
    public override double Perimeter() { return 2 * (width + height); }
}

class Square : Rectangle
{
    public Square(double side) : base(side, side)
    {
        Name = "Square";
    }
}

class Triangle : Shape
{
    private readonly double a, b, c;

    public Triangle(double a, double b, double c)
    {
        if (a + b <= c || a + c <= b || b + c <= a)
            throw new ArgumentException("These sides cannot form a triangle");
        this.a = a; this.b = b; this.c = c;
        Name = "Triangle";
    }

    public override double Perimeter() { return a + b + c; }

    public override double Area()
    {
        // Heron's formula
        double s = Perimeter() / 2;
        return Math.Sqrt(s * (s - a) * (s - b) * (s - c));
    }
}

class Program
{
    static void Main()
    {
        List<Shape> shapes = new List<Shape>
        {
            new Circle(2),
            new Rectangle(3, 5),
            new Square(4),
            new Triangle(3, 4, 5)
        };

        shapes.Sort((x, y) => x.Area().CompareTo(y.Area()));

        Console.WriteLine("Shapes sorted by area:");
        foreach (Shape s in shapes)
        {
            Console.WriteLine(s);
        }

        try
        {
            new Triangle(1, 2, 10);
        }
        catch (ArgumentException e)
        {
            Console.WriteLine("\nError: " + e.Message);
        }
    }
}
