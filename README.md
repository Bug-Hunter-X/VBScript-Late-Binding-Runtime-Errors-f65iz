# VBScript Late Binding Runtime Errors
This repository demonstrates a common runtime error in VBScript caused by late binding and provides a solution using early binding and error handling. Late binding, while flexible, can lead to unexpected errors if objects or methods don't exist.  Early binding and robust error handling help mitigate this risk.

## Bug Description
The bug showcases how using `CreateObject` with a non-existent object results in a runtime error only detected during script execution. This can be particularly problematic in larger, complex scripts where the error's source is not immediately apparent.

## Solution
The solution demonstrates how to use early binding (declaring object variables with explicit types) and error handling (`On Error Resume Next` or `On Error GoTo`) to make the script more robust and prevent unexpected crashes.