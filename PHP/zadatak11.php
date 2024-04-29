<?php
function test($s)
{
    if ((substr($s,0,1)=="F") && (substr($s,strlen($s)-1,1) =="B"))
            {
                return "FizzBuzz";
            }
            else if (substr($s,0,1) == "F")
            {
                return "Fizz";
            }
            else if (substr($s,strlen($s)-1,1) =="B")
            {
                return "Buzz";
            }
            else
            {
                return $s;
            }
}

echo test("FizzBuzz")."\n";
echo test("Fizz")."\n";
echo test("Buzz")."\n";
echo test("Founder")."\n";