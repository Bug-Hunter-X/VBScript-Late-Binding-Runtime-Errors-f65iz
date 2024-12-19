Late Binding: VBScript's late binding can lead to runtime errors that are difficult to debug during development.  If an object or method doesn't exist, the error might not be caught until the script is executed.
Example:
```vbscript
Dim obj
Set obj = CreateObject("NonExistent.Object")
'Error occurs at runtime if NonExistent.Object doesn't exist
```