---
layout: post
title:  "VBA - 撰寫 Visual Basic 程序"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

## 撰寫 Sub 程序

**Sub**程序是一系列的 Visual Basic 陳述式前後加上**Sub**到**End Sub** 陳述式所執行的動作，但不會傳回一個值。 **Sub** 程序可能需要傳遞引數，例如常數、變數或運算式，會呼叫程序。 如果子程序沒有任何引數， **Sub**陳述式必須包含括號空集合。<!--more-->

## 撰寫屬性程序

屬性程序是一系列的 Visual Basic陳述式，可讓程式設計人員來建立和操作的自訂屬性。
*	屬性程序可以用來建立表單、標準模組和類別模組的唯讀屬性。
*	屬性程序應該用於而不是設定此屬性值時必須執行的程式碼中的Public變數。
*	與公用變數，不同屬性程序可以有指派給他們在物件瀏覽器中的說明字串。

當您建立屬性程序時，它會成為包含該程序的模組的屬性。 Visual Basic 提供下列三種類型的屬性程序。

程序	| 描述
---|---
Property Let | 此程序設定屬性的值。
Property Get | 傳回值屬性的程序。
Property Set | 設定物件的參照程序。

宣告屬性程序的語法如下所示。

\[**Public** \| **Private**\] \[**Static**\] Property \{**Get** \| **Let** \| **Set**\} propertyname \[（arguments 引數）\]\[**As** type\] statements陳述式 **End Property**

屬性程序通常是用於配對： **Property Get**與屬性設定並使用**Property Get**與**Property Let** 。 宣告**Property Get**程序單獨就像是宣告唯讀屬性。 一起使用所有三個屬性程序類型時才有用的**Variant**變數，因為只有一個**Variant**可以包含的物件或其他資料類型資訊。 **Property Set**被預定用於物件;**Property Let**則不。

下表顯示在屬性程序宣告中宣告的必要引數。

程序 | 宣告語法
**Property Get** | **Property Get** propname(1，...， n) **As** type
**Property Let** | **Property Let** propname(1，...、、 n， n + 1)
**Property Set** | **Property Set** propname(1，...、 n， n + 1)

透過最後一個引數 (1，...， n) 下一步] 的第一個引數必須共用相同的名稱和資料類型屬性的所有程序具有相同名稱。

**Property Get**程序宣告會比相關的**Property Let**和**Property Set**宣告一個較少引數。 **Property Get**程序的資料類型必須是相關的**Property Let**和**Property Set**宣告中的最後一個引數 (n + 1) 的資料類型相同。 例如，如果您要宣告下列的**Property Let**程序， **Property Get**宣告都必須以**Property Let**程序的引數具有相同名稱和資料類型使用引數。

```vb
Property Let Names(intX As Integer, intY As Integer, varZ As Variant) 
 ' Statement here. 
End Property 
 
Property Get Names(intX As Integer, intY As Integer) As Variant 
 ' Statement here. 
End Property 
```

**Property Set**的宣告中的最後一個引數的資料類型必須是一種物件類型或Variant。

## 撰寫函式程序
 
**函式**程序是一系列的**Function** 和**End Function** 陳述式所含括的 Visual Basic陳述式。 該函數程序會類似於**Sub** 程序，但函數也可以傳回一個值。

**函數**程序可能需要引數，例如常數、變數或運算式，會以傳入呼叫的程序。 如果該**函數**程序會不有任何引數，其**Function**陳述式必須包含括號空集合。 函式會傳回一個值指派值給它的程序的一個或多個陳述式中的名稱。

在下列範例中，**攝氏**函式會計算從華氏度攝氏度。 函式呼叫時從**Main**程序，包含引數值變數會傳遞至函數。 計算的結果是傳回呼叫程序，並顯示在訊息方塊中。

```vb
Sub Main() 
 temp = Application.InputBox(Prompt:= _ 
 "Please enter the temperature in degrees F.", Type:=1) 
 MsgBox "The temperature is " & Celsius(temp) & " degrees C." 
End Sub 
 
Function Celsius(fDegrees) 
 Celsius = (fDegrees - 32) * 5 / 9 
End Function
```