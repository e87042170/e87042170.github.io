---
layout: post
title:  "VBA - 日期與時間函數"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

VBA 的 **Date** 與 **Time** 函數可以幫助開發者快速轉換格式，或者讓日期或時間的值能夠符合所指定的條件。

## Date Function

使用 **Date** 函數會返回目前的系統時間。

### 語法

```vb
date()
```

### 範例

增加一個按鈕並增加下列函數。

```vb
Private Sub Constant_demo_Click()
   Dim a as Variant
   a = date()
   msgbox "The Value of a : " & a
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
The Value of a : 19/07/2014 
```

## CDate Function

**CDate** 函數能夠將一個有效的日期或時間運算式轉換為`日期`類型。

### 語法

```vb
cdate()
```

### 範例

增加一個按鈕並增加下列函數。

```vb
Private Sub Constant_demo_Click()
   Dim a as Variant
   Dim b as Variant
   
   a = cdate("Jan 01 2020")
   msgbox("The Value of a : " & a)
   
   b = cdate("31 Dec 2050")
   msgbox("The Value of b : " & b)
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
The Value of a : 1/01/2020
The Value of b : 31/12/2050 
```

## DateAdd Function

**DateAdd** 函數會返回一個跟指定時間間隔相加的日期。

### 語法

```vb
DateAdd(interval,number,date)
```

### 參數描述

* **Interval** - 必要參數。 您要增加之時間間隔的字串運算式，它可以是下列設定值。
    * d - 日
    * m - 月
    * y - 年
    * yyyy - 年
    * w - 工作日
    * ww - 週
    * q - 季
    * h - 小時
    * n - 分鐘
    * s - 秒
* **Number** - 必要參數。您想要增加之時間間隔的數值運算式。 它可以是正數 (取得未來的日期) 或負數 (取得過去的日期)。
* **Date** - 必要參數。**Variant (Date)** 或常值，代表要新增間隔的日期。

### 範例

```vb
Private Sub Constant_demo_Click()
   ' Positive Interal
   date1 = 27-Jun-1894
   msgbox("Line 1 : " & DateAdd("yyyy",1,date1))
   msgbox("Line 2 : " & DateAdd("q",1,date1))
   msgbox("Line 3 : " & DateAdd("m",1,date1))
   msgbox("Line 4 : " & DateAdd("y",1,date1))
   msgbox("Line 5 : " & DateAdd("d",1,date1))
   msgbox("Line 6 : " & DateAdd("w",1,date1))
   msgbox("Line 7 : " & DateAdd("ww",1,date1))
   msgbox("Line 8 : " & DateAdd("h",1,"01-Jan-2013 12:00:00"))
   msgbox("Line 9 : " & DateAdd("n",1,"01-Jan-2013 12:00:00"))
   msgbox("Line 10 : "& DateAdd("s",1,"01-Jan-2013 12:00:00"))
  
   ' Negative Interval
   msgbox("Line 11 : " & DateAdd("yyyy",-1,date1))
   msgbox("Line 12 : " & DateAdd("q",-1,date1))
   msgbox("Line 13 : " & DateAdd("m",-1,date1))
   msgbox("Line 14 : " & DateAdd("y",-1,date1))
   msgbox("Line 16 : " & DateAdd("w",-1,date1))
   msgbox("Line 19 : " & DateAdd("n",-1,"01-Jan-2013 12:00:00"))
   msgbox("Line 17 : " & DateAdd("ww",-1,date1))
   msgbox("Line 18 : " & DateAdd("h",-1,"01-Jan-2013 12:00:00"))
   msgbox("Line 15 : " & DateAdd("d",-1,date1))
   msgbox("Line 20 : " & DateAdd("s",-1,"01-Jan-2013 12:00:00")) 
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
Line 1 : 27/06/1895
Line 2 : 27/09/1894
Line 3 : 27/07/1894
Line 4 : 28/06/1894
Line 5 : 28/06/1894
Line 6 : 28/06/1894
Line 7 : 4/07/1894
Line 8 : 1/01/2013 1:00:00 PM
Line 9 : 1/01/2013 12:01:00 PM
Line 10 : 1/01/2013 12:00:01 PM
Line 11 : 27/06/1893
Line 12 : 27/03/1894
Line 13 : 27/05/1894
Line 14 : 26/06/1894
Line 15 : 26/06/1894
Line 16 : 26/06/1894
Line 17 : 20/06/1894
Line 18 : 1/01/2013 11:00:00 AM
Line 19 : 1/01/2013 11:59:00 AM
Line 20 : 1/01/2013 11:59:59 AM
```

## DateDiff Function

**DateDiff** 函數會返回一隔時間間隔，它是兩個指定時間的時間差。

### 語法

```vb
DateDiff(interval, date1, date2 [,firstdayofweek[, firstweekofyear]]) 
```

### 參數描述

* **Interval** - 必要參數。  字串運算式 也就是您用來計算介於 date1 和 date2 之間差異的時間間隔，它可以是下列設定值。
    * d - 日
    * m - 月
    * y - 年
    * yyyy - 年
    * w - 工作日
    * ww - 週
    * q - 季
    * h - 小時
    * n - 分鐘
    * s - 秒
* **Date1** and **Date2** - 必要參數。您想要用於計算的兩個日期。
* **Firstdayofweek** − 非必要參數. 會指定每週的第一天。 如果未指定，則會假設星期日，它可以是下列設定值。
   * 0 = vbUseSystemDayOfWeek - 使用 NLS API 設定
   * 1 = vbSunday - 星期日(預設值)
   * 2 = vbMonday - 星期一
   * 3 = vbTuesday - 星期二
   * 4 = vbWednesday - 星期三
   * 5 = vbThursday - 星期四
   * 6 = vbFriday - 星期五
   * 7 = vbSaturday - 星期六
* **Firstdayofweek** − 非必要參數. 會指定每年的第一週。 如果未指定，則會假設1月1日那週為第一週，它可以是下列設定值。
   * 0 = vbUseSystem - 使用 NLS API 設定
   * 1 = vbFirstJan1 - 從 1 月 1 日發生當週開始 (預設值)
   * 2 = vbFirstFourDays - 從新年度的第一週至少四天開始
   * 3 = vbFirstFullWeek - 從該年第一個完整的一週開始

### 範例

增加一個按鈕並增加下列函數。

```vb
Private Sub Constant_demo_Click()
   Dim fromDate as Variant
   fromDate = "01-Jan-09 00:00:00"
   
   Dim toDate as Variant
   toDate = "01-Jan-10 23:59:00"
   
   msgbox("Line 1 : " &DateDiff("yyyy",fromDate,toDate))
   msgbox("Line 2 : " &DateDiff("q",fromDate,toDate))
   msgbox("Line 3 : " &DateDiff("m",fromDate,toDate))
   msgbox("Line 4 : " &DateDiff("y",fromDate,toDate))
   msgbox("Line 5 : " &DateDiff("d",fromDate,toDate))
   msgbox("Line 6 : " &DateDiff("w",fromDate,toDate))
   msgbox("Line 7 : " &DateDiff("ww",fromDate,toDate))
   msgbox("Line 8 : " &DateDiff("h",fromDate,toDate))
   msgbox("Line 9 : " &DateDiff("n",fromDate,toDate))
   msgbox("Line 10 : "&DateDiff("s",fromDate,toDate))
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
Line 1 : 1
Line 2 : 4
Line 3 : 12
Line 4 : 365
Line 5 : 365
Line 6 : 52
Line 7 : 52
Line 8 : 8783
Line 9 : 527039
Line 10 : 31622340
```

應用於**日期**的函數很多，大約14個，如果有需要用到的函數再自行 google 一下囉。

接下來介紹幾個常用的**時間**函數。

## Now Function

**Now**函數會返回目前系統的日期及時間。

### 語法

```vb
Now()
```

### 範例

增加一個按鈕並增加下列函數。

```vb
Private Sub Constant_demo_Click()
   Dim a as Variant
   a = Now()
   msgbox("The Value of a : " & a)
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
The Value of a : 19/07/2013 3:04:09 PM 
```

## Time Function

**Time**函數會返回目前系統的時間。

### 語法

```vb
Time()
```

### 範例

增加一個按鈕並增加下列函數。

```vb
Private Sub Constant_demo_Click()
   msgbox("Line 1: " & Time())
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
Line 1: 3:29:15 PM 
```

## Timer Function

**Timer**函數會返回一個從 12:00 AM 到現在的`秒` + `毫秒`的數值。

### 語法

```vb
Timer()
```

### 範例

增加一個按鈕並增加下列函數。

```vb
Private Sub Constant_demo_Click()
   msgbox("Time is : " & Now())
   msgbox("Timer is: " & Timer())
End Sub
```

當你執行這個函數，你會得到下列輸出。

```vb
Time is : 19/07/2013 3:45:53 PM
Timer is: 56753.4 
```

應用於**時間**的函數很多，大約8個，這邊簡單介紹幾個常用的時間函數，如果有其他需要用到的函數再自行 google 一下囉。
