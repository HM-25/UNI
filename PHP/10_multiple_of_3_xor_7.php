<?php
// Zadatak 10: True if a number is a multiple of 3 or 7, but not both
function test($n)
{
   return ($n % 3 == 0) ^ ($n % 7 == 0);
 }

var_dump(test(3));
var_dump(test(7));
var_dump(test(21));
