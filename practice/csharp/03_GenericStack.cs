// Exercise 3: Generic stack and bracket checker
//
// Own implementation of a generic stack (array based, grows when full),
// then used to check if brackets in an expression are balanced.

using System;

class MyStack<T>
{
    private T[] items = new T[4];
    private int count = 0;

    public int Count { get { return count; } }
    public bool IsEmpty { get { return count == 0; } }

    public void Push(T item)
    {
        if (count == items.Length)
        {
            Array.Resize(ref items, items.Length * 2);
        }
        items[count++] = item;
    }

    public T Pop()
    {
        if (IsEmpty) throw new InvalidOperationException("Stack is empty");
        T item = items[--count];
        items[count] = default(T);
        return item;
    }

    public T Peek()
    {
        if (IsEmpty) throw new InvalidOperationException("Stack is empty");
        return items[count - 1];
    }
}

class Program
{
    static bool IsBalanced(string expression)
    {
        MyStack<char> stack = new MyStack<char>();

        foreach (char c in expression)
        {
            if (c == '(' || c == '[' || c == '{')
            {
                stack.Push(c);
            }
            else if (c == ')' || c == ']' || c == '}')
            {
                if (stack.IsEmpty) return false;
                char open = stack.Pop();
                if ((c == ')' && open != '(') || (c == ']' && open != '[') || (c == '}' && open != '{'))
                    return false;
            }
        }
        return stack.IsEmpty;
    }

    static void Main()
    {
        MyStack<int> numbers = new MyStack<int>();
        for (int i = 1; i <= 10; i++) numbers.Push(i * i);

        Console.Write("Popping squares: ");
        while (!numbers.IsEmpty) Console.Write(numbers.Pop() + " ");
        Console.WriteLine();

        string[] tests = { "(a + b) * [c - d]", "{[()]}", "(a + b", "([)]", "" };
        foreach (string t in tests)
        {
            Console.WriteLine("{0,-20} -> {1}", "\"" + t + "\"", IsBalanced(t) ? "balanced" : "NOT balanced");
        }
    }
}
