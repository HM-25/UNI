<?php
// Zadatak 4: True if a number is within 10 of 100 or 200
function test($x) 
{
   if(abs($x - 100) <= 10 || abs($x - 200) <= 10)
            return true;
     return false;
}

var_dump(test(103));
var_dump(test(90));
var_dump(test(89));
