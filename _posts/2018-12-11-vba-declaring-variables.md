---
layout: post
title:  "VBA - 宣告變數"
categories: [develop]
tags: [vba]
excerpt_separator: <!--more-->
---

當宣告變數時，您通常會使用 **Dim** 陳述式。 可以將宣告陳述式放置於程序內，以建立程序層級變數。 

下列範例會建立變數，並指定字串資料類型。

```vb
Dim strName As String 
```

如果此陳述式會出現在程序中，則變數 **strName** 只能用在該程序。 如果在**模組**的宣告區段中出現陳述式，則變數 **strName** 可供**模組**內所有的程序使用，而非用於專案中其他模組內的程序。<!--more-->

若要讓此變數可用於**專案**中的所有程序，請將 **Public** 陳述式置於變數之前，如下列範例中所示：

```vb
Public strName As String 
```

可將變數宣告為下列其中一種資料類型：**Boolean、Byte、Integer、Long、Currency、Single、Double、Date、String** (適用於可變長度變數)、**String * length** (適用於固定長度字串)、**Object** 或 **Variant**。 

如果您未指定資料類型，便會依預設指派 **Variant** 資料類型。 您也可以使用 **Type** 陳述式來建立使用者定義類型。

您可以在一個陳述式中宣告多個變數。 若要指定資料類型，您必須包含每個變數的資料類型。

在下列陳述式中，變數 `intX`、`intY` 及 `intZ` 宣告為類型 **Integer**。

```vb
Dim intX As Integer, intY As Integer, intZ As Integer 
```

在下列陳述式中，`intX` 和 `intY` 宣告為類型 **Variant**，而且僅 `intZ` 宣告為類型 **Integer**。

```vb
Dim intX, intY, intZ As Integer 
```

您沒有在宣告陳述式中提供變數的資料類型。 如果您省略資料類型，變數將會是 **Variant** 類型。

在上述陳述式將 `x` 和 `y` 宣告為整數的簡略表示法為：

```vb
Dim intX%, intY%, intZ as Integer
```

類型的簡略表示法為：% -integer; & -long; @ -currency; # -double; ! -single; $ -string

## Public 陳述式

您可以使用 **Public** 陳述式來宣告公用模組層級變數。

```vb
Public strName As String 
```

**公用變數**可以用於**專案中**的任何程序。 如果在標準模組或類別模組中宣告公用變數，它也可用於任何在參考專案中已宣告公用變數的專案。

## Private 陳述式

您可以使用 **Private** 陳述式來宣告私人模組層級變數。

```vb
Private MyName As String 
```

只有**相同模組**中的程序，才可使用**私人變數**。

當用於**模組層級**時，**Dim** 陳述式便相當於 **Private** 陳述式。 您可能想要使用 **Private** 陳述式，讓您的程式碼更容易讀取與解譯。

## Static 陳述式

當您使用 **Static** 陳述式，而不是 **Dim** 陳述式來宣告程序中的變數時，宣告的變數會保留其呼叫該程序之間的值。

## Option Explicit 陳述式

您只要在指派的陳述式中使用變數，便等於在 Visual Basic 中隱含宣告變數。 隱含宣告的所有變數都是屬於 **Variant** 類型。 **Variant** 類型的變數會比其他大部分的變數，需要更多的記憶體資源。 

如果您能夠明確地和特定資料類型宣告變數，應用程式將會更有效率。 明確宣告所有變數，可減少命名衝突錯誤及拼字錯誤的發生率。

如果您不想讓 Visual Basic 作隱含的宣告，可以在進行任何程序之前，將 **Option Explicit** 陳述式放置於模組中。 此陳述式會要求您在模組內明確地宣告所有變數。 如果模組包含 **Option Explicit** 陳述式，當 Visual Basic 遇到未於之前宣告的變數名稱或名稱拼寫不正確時，就會發生編譯時期錯誤。
您可以在 Visual Basic 程式設計環境中設定選項，在所有的新模組中自動包含 **Option Explicit** 的陳述式。 請注意這個選項不會變更已寫入的現有程式碼。

您必須明確地宣告**固定陣列**和**動態陣列**。

## 宣告自動化物件變數

當您使用一個應用程式來控制另一個應用程式的物件時，您應該設定參照至另一個應用程式的**型別程式庫**。 在您設定參照之後，可以根據其最特定的類型來宣告物件變數。 例如，將參照設定為 Microsoft Excel 的型別程式庫時，如果是在 Microsoft Word 中，您可以代表 Excel 的**工作表**物件，從 Word 中來宣告類型**工作表**的變數。

如果您正在使用另一個應用程式來控制 Microsoft Access 物件，在大部分的情況下，您可以根據其最特定的類型來宣告物件變數。 您也可以使用**New**關鍵字以自動建立物件的新執行個體。 不過，您可能需要指示其為 Microsoft Access 物件。 例如，當您從 Visual Basic 內代表 Access 表單來宣告物件變數時，您必須將 Access 表單物件，和 Visual Basic 的表單物件作區別。 包含在變數宣告中的型別程式庫名稱，如下列範例所示：

```vb
Dim frmOrders As New Access.Form 
```

某些應用程式無法識別個別的 Access 物件類型。 即使您從這些應用程式設定參照至 Access 類型程式庫，您必須將所有的 Access 物件變數宣告為 Object 類型。 

您也不可以使用**New**關鍵字來建立物件的新執行個體。

下列範例顯示如何從無法識別的 Access 物件類型的**Application**，來宣告 Access **Application**物件的執行實體物件變數。 然後，應用程式會建立**Application**物件的執行實體物件。

```vb
Dim appAccess As Object 
Set appAccess = CreateObject("Access.Application")
```
