'check updates @ https://github.com/atuncuti/vbs
'ver 0.1
Dim Input
Dim substr 
Dim strlen
Dim even
Dim odd
Dim result
Dim ck

'Accept Input 
Input  = InputBox( "enter code" )
strlen = len(Input)

'exit if blanks
if Input="" then
	WScript.Quit
end if


even = 0
odd = 0

'Get each position and its number.
'Add the digits in the odd-numbered positions (first, third, fifth, etc.) together.
'Add the digits in the even-numbered positions (second, fourth, sixth, etc.).
For i = 1 to strlen
	substr = mid(Input,i,1)
	if i mod 2 = 0 then 
		even = even + substr
	else
		odd = odd + substr
	end if
next 

'Multiply the Odd total by three:
odd = odd * 3

'Add the two results together:
result = even + odd

'Now what single digit number makes the total a multiple of 10? That’s the check digit.
'So
'Get the last digit no
strlen = len(result)
substr = mid(result,strlen,1)
'Now Minus that with 10 to get your check digit number
if substr = 0 then
	ck = 0
else
	ck = 10 - substr
end if

WScript.Echo "Check Digit :    " & Input & "[" & ck & "]"
