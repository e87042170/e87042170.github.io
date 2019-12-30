---
layout: post
title:  "VBA - 宣告陣列"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

**Arrays**的宣告方式與其他變數相同，方法是使用 **Dim**、**Static**、**Private** 或 **Public** 陳述式。 純量變數 (非陣列的變數) 與**陣列**變數之間的差異為，您通常必須指定**陣列**的大小。
指定大小的陣列是**固定陣列**。 您可以在執行程式時，變更其大小的陣列是**動態陣列**。<!--more-->

將陣列從 0 或 1 建立索引，取決於 **Option Base** 陳述式的設定。 如果未指定 **Option Base 1**，則所有的陣列索引會從**0**開始。

## 宣告固定陣列

在下列程式碼中，固定陣列宣告為具有 11 列和 11 行的 Integer 陣列：

```vb
Dim MyArray(10, 10) As Integer 
```

第一個引數代表列，第二個引數代表行。

如同任何其他變數宣告，除非您為陣列指定資料類型，否則宣告陣列中元素的資料類型為 **Variant**。 陣列的每個數值 **Variant** 元素會使用 16 個位元組。 每個字串 **Variant** 元素會使用 22 個位元組。 若要撰寫盡可能精簡的程式碼，請將您的陣列明確宣告為 **Variant** 以外的資料類型。

下列幾行程式碼會比較數個陣列的大小。

```vb
' Integer array uses 22 bytes (11 elements * 2 bytes). 
ReDim MyIntegerArray(10) As Integer 
 
' Double-precision array uses 88 bytes (11 elements * 8 bytes). 
ReDim MyDoubleArray(10) As Double 
 
' Variant array uses at least 176 bytes (11 elements * 16 bytes). 
ReDim MyVariantArray(10) 
 
' Integer array uses 100 * 100 * 2 bytes (20,000 bytes). 
ReDim MyIntegerArray (99, 99) As Integer 
 
' Double-precision array uses 100 * 100 * 8 bytes (80,000 bytes). 
ReDim MyDoubleArray (99, 99) As Double 
 
' Variant array uses at least 160,000 bytes (100 * 100 * 16 bytes). 
ReDim MyVariantArray(99, 99) 
```

陣列的大小上限，將根據您的作業系統與可用的記憶體而有所不同。 使用超出您的系統上可用 RAM 數量的陣列會比較緩慢，因為必須往返磁碟讀取及寫入資料。

## 宣告動態陣列

透過宣告動態陣列，您便可以在程式碼執行時調整陣列的大小。 使用 **Static**、**Dim**、**Private** 或 **Public** 陳述式來宣告陣列，將括號內保留空白，如下列範例所示。

```vb
Dim sngArray() As Single 
```

您可以使用 **ReDim** 陳述式隱含宣告程序內的陣列。 使用 **ReDim** 陳述式時請務必注意，不要將陣列的名稱輸入錯誤。 即使已在模組中包含 **Option Explicit** 陳述式，系統仍會建立第二個陣列。

在陣列的 **scope** 程序內，使用 **ReDim** 陳述式來變更維度的數量、定義元素數量，以及定義每個維度的上限和下限。 您可以使用 **ReDim** 陳述式，視需要變更動態陣列。 不過，每次您執行此動作時，陣列中現有的值將會遺失。 使用 **ReDim Preserve** 來擴大陣列，同時保留陣列中現有的值。

例如，下列陳述式會將陣列放大 10 個元素，而不會遺失原始元素目前的值。

```vb
ReDim Preserve varArray(UBound(varArray) + 10) 
```

使用**Preserve** 關鍵字搭配陣列時，您只可以變更最後一個維度的上限，但無法變更維度的數量。
