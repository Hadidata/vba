# Array Functions

This Module is a list of functions that I use to make working with arrays in
VBA easier. I will continue to add more functions to this Module. 

Tested in Windows Excel 2016, but should work with Excel 2007+. (Not fully tested in Mac yet)

- For Windows requires reference to "Microsoft Scripting Runtime"


# Funtions

list of Functions that are currently avaliable in the Module 
as I add more functions I will update this list.

- Transpose
- RemoveDuplicates

## Transpose

```
'===========================================================
' Transpose Public Function
' ----------------------------------------------------------
' Purpose: To transpose a 2 dimmensional array clockwise
'          is not limited by the Application.Transpose
'          ~64,000 row limit
'
'Parameter(s)
'-----------
'varArray: A 2 dimmensional Array
'-----------------------------------------------------------
'Returns: A 2 dimmensional Array tranposed counterclockwise
'         if an error occurs then a null value is returned
'-----------------------------------------------------------
'Revision History
'-----------------------------------------------------------
'20Nov19 HA: Initial Version
'===========================================================
```

### Examples
```
'Tranposing a 2 Dimmensional Array
Sub TranposeExample()

   Dim varResult As Variant
   Dim varExampleArray As Variant
   
   'creating a 2 dimmensional Array
   varExampleArray = [{1,2;3,4;5,6;7,8}]
   '[[1,2],
   ' [3,4],
   ' [5,6],
   ' [7,8]]
   
   varResult = ArrayFun.Transpose(varExampleArray)
   '[[1,3,5,7],
   '[2,4,6,8]],
   
End Sub

```
## RemoveDuplicates

```

'===========================================================
' RemoveDuplicates Public Function
' ----------------------------------------------------------
' Purpose: To remove Duplicates from a 1D or a 2D array
'
' Author: Hadi Ali November 2019
'
'Parameter(s)
'-----------
'varArray: A 2 dimmensional or 1 dimmensional Array
'-----------------------------------------------------------
'Returns: A 2 dimmensional or 1 dimmensional Array
'         without its duplicate if an error occurs
'         then a null value is returned
'-----------------------------------------------------------
'Revision History
'-----------------------------------------------------------
'20Nov19 HA: Initial Version
'===========================================================

```

### Examples

```
'Remove Duplicates of a 2 Dimmensional Array
Sub RemoveDuplicatesExample2D()

   Dim varResult As Variant
   Dim varExampleArray As Variant
   
   'creating a 2 dimmensional Array
   varExampleArray = [{1,2;1,2;5,6;7,8}]
   '[[1,2],
   ' [3,4],
   ' [5,6],
   ' [7,8]]
   
   varResult = ArrayFun.RemoveDuplicates(varExampleArray)
   '[[1,2],
   ' [5,6],
   ' [7,8]]
   
End Sub

```
```
'Remove Duplicates of a 1 Dimmensional Array
Sub RemoveDuplicatesExample1D()

   Dim varResult As Variant
   Dim varExampleArray As Variant
   
   'creating a 2 dimmensional Array
   varExampleArray = Array(1, 2, 3, 3, 4, 4, 5, 6)
   '[1,2,3,3,4,4,5,6]
   
   varResult = ArrayFun.RemoveDuplicates(varExampleArray)
   '[1,2,3,4,5,6]
   
End Sub
```
