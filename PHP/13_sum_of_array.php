<?php
// Zadatak 13: Sum all elements of an array
function variadicFunction($operands)
{
    $sum = 0;
    foreach($operands as $singleOperand) {
        $sum += $singleOperand;
    }
    return $sum;
}

var_dump(variadicFunction([1, 2]));
var_dump(variadicFunction([1, 2, 3, 4]));
  
?>
