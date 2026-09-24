<?php
// Zadatak 5: Add "if" to the start of a string unless it already starts with it
function test($s)
{
  if (strlen($s) >= 2 && substr($s,0, 2) == "if")
            {
                return $s;
            }
      return "if ".$s;
}
echo test("if else")."\n";
echo test("else")."\n";
echo test("if")."\n";
