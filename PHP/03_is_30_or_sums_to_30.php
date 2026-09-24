<?php
// Zadatak 3: True if either number is 30 or their sum is 30
function test($x, $y) 
{
    return ($x == 30) || ($y == 30) || ($x + $y == 30);
}

var_dump(test(30, 0));
var_dump( test(25, 5));
var_dump( test(20, 30));
var_dump(test(20, 25));
