---
layout: post
title:  "VBA - 使用迴圈"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

## 使用 Do...Loop 陳述式

你可以使用 **Do...Loop** 陳述式，去重複執行一個區塊的陳述式無數次，直到預設條件變為 **True** 或某個條件變為 **True**。

## 重複陳述式，當條件為 True

在 **Do...Loop** 陳述式中，有2種使用**While**關鍵字去檢查條件的方式。一種是在進入迴圈前先檢查條件；另一種是先執行迴圈一次，結束迴圈時再檢查條件。<!--more-->

在以下的`ChkFirstWhile`程序中，在進入迴圈前先檢查條件。 如果`myNum`設為 9，而不是 20，將永遠不會執行迴圈會內的陳述式。 在 `ChkLastWhile`程序中，迴圈的陳述式先執行一次後條件會變為**False**。

```vb
Sub ChkFirstWhile() 
    counter = 0 
    myNum = 20 
    Do While myNum > 10 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub 
 
Sub ChkLastWhile() 
    counter = 0 
    myNum = 9 
    Do 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop While myNum > 10 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub
```

## 重複陳述式，直到條件變成 True

在 **Do...Loop** 陳述式中，有2種使用**Until**關鍵字去檢查條件的方式。一種是在進入迴圈前先檢查條件(如下列`ChkFirstUntil`程序所示)；另一種是先執行迴圈一次(如下列`ChkLastUntil`程序所示)，結束迴圈時再檢查條件，條件保持為 False時，會繼續執行迴圈。

```vb
Sub ChkFirstUntil() 
    counter = 0 
    myNum = 20 
    Do Until myNum = 10 
        myNum = myNum - 1 
        counter = counter + 1 
    Loop 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub 
 
Sub ChkLastUntil() 
    counter = 0 
    myNum = 1 
    Do 
        myNum = myNum + 1 
        counter = counter + 1 
    Loop Until myNum = 10 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub
```

## 結束在迴圈中的 Do...Loop 陳述式

您可以在**Do...Loop**迴圈中使用**Exit Do** 陳述式來結束迴圈。 為了避免陷入無限迴圈，你可以在迴圈裡使用 **If...Then...Else** 或 **Select...Case** 陳述式，當條件成立時，你可以使用**Exit Do** 陳述式來結束迴圈；當條件不成立，則繼續執行迴圈。

下列範例中，`myNum`會被指派一個值，使得陳述式陷入無限迴圈。 加入 **If...Then...Else**檢查條件，當條件成立就結束迴圈，避免陷入無限迴圈。

```vb
Sub ExitExample() 
    counter = 0 
    myNum = 9 
    Do Until myNum = 10 
        myNum = myNum - 1 
        counter = counter + 1 
        If myNum < 10 Then Exit Do 
    Loop 
    MsgBox "The loop made " & counter & " repetitions." 
End Sub
```

若要停止無限迴圈，你可以按下 `esc` 鍵或 `CTRL` + `BREAK`。

## 使用 For Each...Next 陳述式

**For Each...Next** 陳述式會對**集合**中的每個物件或陣列中的每個元素重複一個區塊的陳述式。 Visual Basic 會在每次執行迴圈時自動設定變數。 

例如，下列程序會關閉所有表單，包含正執行的程序以外的表單除外。

```vb
Sub CloseForms() 
 For Each frm In Application.Forms 
 If frm.Caption <> Screen. ActiveForm.Caption Then frm.Close 
 Next 
End Sub
```

下列程式碼會循環查看陣列中的每個元素，並將每個元素設定為索引變數 I 的值。

```vb
Dim TestArray(10) As Integer, I As Variant 
For Each I In TestArray 
 TestArray(I) = I 
Next I 
```

## 循環查看儲存格範圍

使用 **For Each...Next** 迴圈，循環查看範圍內的儲存格。 下列程序會循環查看 Sheet1 的範圍 A1:D10，並將所有絕對值小於 0.01 的任何數字設定為 0 (零)。

```vb
Sub RoundToZero() 
 For Each myObject in myCollection 
 If Abs(myObject.Value) < 0.01 Then myObject.Value = 0 
 Next 
End Sub
```

## 在 For Each...Next 迴圈完成之前結束它

您可以使用 **Exit For** 陳述式來結束 **For Each...Next** 迴圈。 比方說，當錯誤發生時，請在 **If...Then...Else** 陳述式裡面，或者在 **Select Case** 陳述式的 **True**陳述式區塊中使用 **Exit For** 陳述式。 如果錯誤沒有發生，則 **If...Then...Else** 陳述式為 **False**，且迴圈會繼續如預期般執行。

下列範例會測試範圍 A1:B5 中`不包含數字`的第一個儲存格。 如果找到這類儲存格，則會顯示一則訊息，且 **Exit For** 會結束該迴圈。

```vb
Sub TestForNumbers() 
 For Each myObject In MyCollection 
 If IsNumeric(myObject.Value) = False Then 
 MsgBox "Object contains a non-numeric value." 
 Exit For 
 End If 
 Next c 
End Sub
```

## 使用 For...Next 陳述式

你可以使用 **For...Next** 陳述式，在你所規定的次數內重複執行某個陳述式區塊。**For**迴圈使用一個計數器變數，它會隨著每次執行迴圈後，增加或減少它的值。

下列的程序會讓電腦發出50次的嗶聲。**For**陳述式會定義計數器變數**x**一個1到50的值。**Next**陳述式會把計數器變數**x**加1。

```vb
Sub Beeps() 
    For x = 1 To 50 
        Beep 
    Next x 
End Sub
```

使用**Step**關鍵字，可以根據你的定義來增加或減少計數器變數值。在下列的例子中，計數器變數`j`在每次執行完迴圈會加2，當迴圈結束時，`total`的值會是2,4,6,8,10的加總。

```vb
Sub TwosTotal() 
    For j = 2 To 10 Step 2 
        total = total + j 
    Next j 
    MsgBox "The total is " & total 
End Sub
```

你也可以使用一個負的**Step**值來遞減計數器變數。想要遞減計數器變數，你必須規定一個小於初始值的結束值。在下列例子中，計數器變數`myNum`在每次執行完迴圈會減2，當迴圈結束時，`total`的值會是16,14,12,10,8,6,4,2的加總。

```vb
Sub NewTotal() 
    For myNum = 16 To 2 Step -2 
        total = total + myNum 
    Next myNum 
    MsgBox "The total is " & total 
End Sub
```

**Next**陳述式後面的計數器變數名稱，只是為了增加程式的可讀性，因此它可加，也可不加。

你也可以使用**Exit For**陳述式，在**For...Next**陳述式的計數器還沒到達結束值前先結束迴圈。比方說，當錯誤發生時，你可以在 **If...Then...Else** 陳述式裡面，或者在 **Select Case** 陳述式的 **True**陳述式區塊中使用 **Exit For** 陳述式。 如果錯誤沒有發生，則 **If...Then...Else** 陳述式為 **False**，迴圈會繼續如預期般執行。
