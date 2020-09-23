<div align="center">

## Sort arrays


</div>

### Description

Vbscript function that sorts

an one-dimensional array

of number either ascending

or descending.

This function sorts using

the well known double

for..next principle.

Script gives an overview

of passing parameters

and especially passing

arrays of numbers

to a function.
 
### More Info
 
The function gets as

input an unsorted

array, plus 1 for

ascending sorting

or something else

for descending sorting

Using SQL sorting is not

a problem. However, for

those cases when there

is no SQL to sort with,

or an array of numbers

is manipulated requiring

a sort, this function

might come in handy.

Also, VBscript does

not have a build-in

sorting function.

The function returns

a sorted array of numbers.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rogier Doekes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rogier-doekes.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Sorting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sorting__4-24.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rogier-doekes-sort-arrays__4-6495/archive/master.zip)





### Source Code

```
'--------Begin Function----
Function fnSort(aSort, intAsc)
 Dim intTempStore
 Dim i, j
 For i = 0 To UBound(aSort) - 1
 For j = i To UBound(aSort)
 'Sort Ascending
 If intAsc = 1 Then
 If aSort(i) > aSort(j) Then
  intTempStore = aSort(i)
  aSort(i) = aSort(j)
  aSort(j) = intTempStore
 End If 'i > j
 'Sort Descending
 Else
 If aSort(i) < aSort(j) Then
  intTempStore = aSort(i)
  aSort(i) = aSort(j)
  aSort(j) = intTempStore
 End If 'i < j
 End If 'intAsc = 1
 Next 'j
 Next 'i
 fnSort = aSort
End Function 'fnSort
'-------------------------
Dim aUnSort(3), aSorted
aUnSort(0) = 4
aUnSort(1) = 2
aUnSort(2) = 6
aUnSort(3) = 20
'call the function
'second argument:
' * ascending sorted = 1
' * descending sorting = any other character
aSorted = fnSort(aUnSort, 1)
Erase aUnSort
```

