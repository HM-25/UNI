<?php
// Zadatak 6: Repeat the first two characters of a string four times
function test($str) 
{
   return strlen($str) < 2 ? $str : substr($str, 0, 2).substr($str, 0, 2).substr($str, 0, 2).substr($str, 0, 2);
   
}

echo test("C Sharp")."\n";
echo test("JS")."\n";
echo test("a")."\n";
