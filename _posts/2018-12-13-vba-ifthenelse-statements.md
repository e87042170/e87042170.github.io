---
layout: post
title:  "VBA - 使用判斷式"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

## 使用 If...Then...Else 陳述式

您可以使用 **If...Then...Else** 陳述式執行特定陳述式或區塊陳述式 (視條件的值而定)。 **If...Then...Else** 陳述式可以使用無限多層巢狀層級。
不過為了方便閱讀理解方便，您可能需要使用 **Select Case** 陳述式，而非多層巢狀的 **If...Then...Else** 陳述式。<!--more-->

## 若條件為 True 即執行陳述式

如果要當條件為 **True** 時只執行一個陳述式，請使單行語法的 **If...Then...Else** 陳述式。 下列範例展示單行語法，省略 **Else** 關鍵字。

```vb
Sub FixDate() 
 myDate = #2/13/95# 
 If myDate < Now Then myDate = Now 
End Sub
```

若要執行一行以上的程式碼，您必須使用多行語法。 此語法包含 **End If** 陳述式，如下列範例所示。

```vb
Sub AlertUser(value as Long) 
 If value = 0 Then 
 AlertLabel.ForeColor = "Red" 
 AlertLabel.Font.Bold = True 
 AlertLabel.Font.Italic = True 
 End If 
End Sub
```

## 若條件為 True 就執行某些陳述式，而若為 False 則執行其他陳述式

使用 **If...Then...Else** 陳述式定義兩個可執行陳述式的區塊：當條件為 **True** 執行一個區塊，而當條件為 **False** 時則執行另個區塊。

```vb
Sub AlertUser(value as Long) 
 If value = 0 Then 
 AlertLabel.ForeColor = vbRed 
 AlertLabel.Font.Bold = True 
 AlertLabel.Font.Italic = True 
 Else 
 AlertLabel.Forecolor = vbBlack 
 AlertLabel.Font.Bold = False 
 AlertLabel.Font.Italic = False 
 End If 
End Sub
```

## 若第一個條件為 False，則測試第二個條件

您可以將 **ElseIf** 陳述式新增至 **If...Then...Else** 陳述式來測試當第一個條件為 **False** 時所使用的第二個條件。 例如，下列函數程序是根據職別計算獎金。 若所有的 **If** 和 **ElseIf** 陳述式皆為 **False**，則執行**Else** 陳述式後面的陳述式。

```vb
Function Bonus(performance, salary) 
 If performance = 1 Then 
 Bonus = salary * 0.1 
 ElseIf performance = 2 Then 
 Bonus = salary * 0.09 
 ElseIf performance = 3 Then 
 Bonus = salary * 0.07 
 Else 
 Bonus = 0 
 End If 
End Function
```

## 使用 Select Case 陳述式

如果你使用多層巢狀的 **If...Then...Else** 陳述式，為了方便閱讀理解，你可以改用 **Select Case** 陳述式。

在下列範例中， **Select Case** 陳述式會評估傳遞至程序的引數。 請注意，每個**Case**陳述式可以包含多個值、 一個範圍的值或值和比較運算子的組合。 如果**Select Case**陳述式不符合任何**Case**陳述式中的值，就會執行選用的**Case Else**陳述式。

```vb
Function Bonus(performance, salary) 
  Select Case performance 
    Case 1 
      Bonus = salary * 0.1 
    Case 2, 3 
      Bonus = salary * 0.09 
    Case 4 To 6 
      Bonus = salary * 0.07 
    Case Is > 8 
      Bonus = 100 
    Case Else 
      Bonus = 0 
  End Select 
End Function 
```
