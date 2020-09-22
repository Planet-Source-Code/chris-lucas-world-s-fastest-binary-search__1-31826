<div align="center">

## World's Fastest Binary Search


</div>

### Description

This ain't your daddy's binary search tool! Mine goes about 30% faster than his. Why people insist on using the same tired and slow function I'll never know, but it suffers from two glaring weaknesses. First and foremost: < and > are about the slowest way to compare strings! Second: The tired old function assumes that you will find a match more often than not which is slow! My function doesn't suffer from either of these two "features".
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris\_Lucas ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-lucas.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 6\.0, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-lucas-world-s-fastest-binary-search__1-31826/archive/master.zip)





### Source Code

```
Public Function FastBinarySearch(ByRef arr() As String, ByRef search As String) As Long
  Dim first As Long
  Dim last As Long
  Dim middle As Long
  first = LBound(arr)
  last = UBound(arr)
  Do
    middle = (first + last) \ 2
    Select Case StrComp(arr(middle), search, vbBinaryCompare)
      Case -1: first = middle + 1
      Case 1: last = middle - 1
      Case 0
        FastBinarySearch = middle
        Exit Function
    End Select
  Loop Until first > last
End Function
```

