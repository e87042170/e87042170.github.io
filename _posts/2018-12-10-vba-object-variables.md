---
layout: post
title:  "VBA - 建立物件變數"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

您可以將物件變數完全相同視為它參照的物件。 您可以設定或傳回物件的屬性或使用下列任何一它的方法。

## 建立物件變數

1.宣告物件變數。
2.將物件變數指定給物件。<!--more-->

## 宣告物件變數

使用**Dim** 陳述式或其中一個其他宣告陳述式 (**Public**, **Private**, or **Static**) 來宣告物件變數。 參照到物件變數必須是物件的**Variant、物件** 或特定類型。 例如，下列宣告是有效的：

```vb
' Declare MyObject as Variant data type. 
Dim MyObject 
' Declare MyObject as Object data type. 
Dim MyObject As Object 
' Declare MyObject as Font type. 
Dim MyObject As Font 
```

如果您不需要先宣告使用的物件變數，該物件變數資料型別預設為Variant 。

直到執行的程序，才會知道特定的物件類型時，您可以與物件資料型別宣告物件變數。 若要建立的任何物件的一般參考使用物件的資料類型。

如果您知道的特定物件類型，您應該物件變數宣告為該物件類型。 例如，如果應用程式包含範例物件類型，您可以使用其中一個這些陳述式宣告該物件的物件變數：

```vb
Dim MyObject As Object ' Declared as generic object. 
Dim MyObject As Sample ' Declared only as Sample object. 
```

宣告特定物件類型提供更快的程式碼，並改善的可讀性檢查自動類型。

## 將物件變數指定給物件

若要指派給物件變數的物件使用**Set** 陳述式。 您可以指派物件運算式或**Nothing**。 例如，下列的物件變數工作分派是有效的。

```vb
Set MyObject = YourObject ' Assign object reference. 
Set MyObject = Nothing ' Discontinue association. 
```

您可以合併宣告物件變數使用**Set**陳述式中使用**New** 關鍵字物件指派給它。 例如：

```vb
Set MyObject = New Object ' Create and Assign 
```

物件變數設定為等於**Nothing**會停止與任何特定物件關聯的物件變數。 這會防止您不小心變更藉由變更變數的物件。 物件變數一律設為**Nothing**之後關閉相關聯的物件，讓您可以測試是否該物件變數指向有效的物件。 例如：

```vb
If Not MyObject Is Nothing Then 
 ' Variable refers to valid object. 
 . . . 
End If 
```

當然，這項測試可以永遠不會判斷絕對使用者已關閉包含的物件變數所參考的物件的應用程式。

## 參照物件的目前執行個體

使用**Me** 關鍵字來參照該物件的目前執行個體執行程式碼。 目前物件相關聯的所有程序具有存取權稱為 「我的物件。 使用**Me**是特別適合用來將物件的目前執行個體的相關資訊傳遞至另一個模組中的程序。 例如，假設您在模組中有下列程序：

```vb
Sub ChangeObjectColor(MyObjectName As Object) 
 MyObjectName.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256) 
End Sub
```

您可以呼叫程序，並使用下列陳述式做為引數傳遞物件的目前執行個體：

```vb
ChangeObjectColor Me 
```
