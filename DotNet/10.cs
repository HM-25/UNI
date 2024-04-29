using System;
using System.Collections.Generic;
using System.IO;

//nadji velicinu fajla

public class test{
 public static void Main() {
        FileInfo f = new FileInfo("main.exe"); //unesi path ovdje unutar apostrofa
        Console.WriteLine("\nvelicina fajla je: "+f.Length.ToString());
 }
}