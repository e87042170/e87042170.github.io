---
layout: post
title:  "VBA - 使用 With 陳述式"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

**With** 陳述式可讓您指定的`物件`或`使用者定義型別`，能一次做整個系列的陳述式，而不需多次指定物件的名稱。 **With**陳述式讓您更快速地執行，並協助您避免重複輸入的程序。<!--more-->

下列範例會使用數字 30 填滿儲存格範圍、 將套用粗體格式設定，並儲存格的內景色彩設定為黃色。

```vb
Sub FormatRange() 
 With Worksheets("Sheet1").Range("A1:C10") 
 .Value = 30 
 .Font.Bold = True 
 .Interior.Color = RGB(255, 255, 0) 
 End With 
End Sub
```

您可以使用巢狀**With**陳述式的更大的效率。 下列範例會插入儲存格 A1 的公式和下例字型。

```vb
Sub MyInput() 
 With Workbooks("Book1").Worksheets("Sheet1").Cells(1, 1) 
 .Formula = "=SQRT(50)" 
 With .Font 
 .Name = "Arial" 
 .Bold = True 
 .Size = 8 
 End With 
 End With 
End Sub
```
