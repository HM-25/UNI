<?php
// Zadatak 2: Difference from 51, tripled if the number is greater than 51
function test($n) 
{
     $x = 51;

     if ($n > $x)
     {
       return ($n - $x)*3;
     }
   return $x - $n;
}
echo test(53)."\n";
echo test(30)."\n";
echo test(51)."\n";
