# Excel VBA | 生物科技原料採購自動化報表系統

# 問題概述
痛點：每月需彙整上百種產品的配料表，人工計算耗時且容易出錯  
影響：生管人員常需假日加班整理資料，容易延誤原料請購時程

# 功能摘要
輸入：輸入各產品的預計生產數量與製損率  
處理：自動計算各原料需求量，彙整相同原料的總需求量與指定供應商  
輸出：產出可直接用於請購的總表，並保留對應訂單明細

# 面對挑戰，找到解法
考慮到未來資料大量增加的時候，依然保持程式的執行效率，比較了多種資料結構之後，我決定採用Dictionary + Dictionary + Array的結構來完成，代碼最核心的邏輯是先依原料名稱進行彙總，再把每個原料對應到不同產品的需求明細保存起來。第一層 Dictionary 負責累加原料總需求量，第二層 Dictionary 負責保留該原料對應的各產品明細


```vba
Dim IngredientName As String
Dim Supplier As String
Dim Dosage As Double
Dim Demand As Double
Dim DetailList As Variant
Dim IngredientDict As Object
Dim IngredientDetailDict As Object
Set IngredientDict = CreateObject("Scripting.Dictionary")
Set IngredientDetailDict = CreateObject("Scripting.Dictionary")

For j = 2 To WS_Formula_Lastrow
        ProductName = WS_Formula.Range("A" & j).Value
        If ProductDict.Exists(ProductName) Then
                IngredientName = WS_Formula.Range("C" & j).Value
                Supplier = WS_Formula.Range("B" & j).Value
                Dosage = Val(WS_Formula.Range("D" & j).Value)
                Demand = Round(Dosage * ProductDict(ProductName), 2)
                
                If Not IngredientDict.Exists(IngredientName) Then
                        IngredientDict(IngredientName) = Array(Supplier, Demand)
                Else
                       Dim TempArr As Variant
                       TempArr = IngredientDict(IngredientName)
                       TempArr(1) = TempArr(1) + Demand
                       IngredientDict(IngredientName) = TempArr
                End If
                
                If Not IngredientDetailDict.Exists(IngredientName) Then
                        IngredientDetailDict(IngredientName) = Array(ProductName & ": " & Demand & "  g")
                Else
                        DetailList = IngredientDetailDict(IngredientName)
                        ReDim Preserve DetailList(UBound(DetailList) + 1)
                        DetailList(UBound(DetailList)) = ProductName & ": " & Demand & "  g"
                        IngredientDetailDict(IngredientName) = DetailList
                End If
        End If
Next j
```

# 成果展示
如下圖，輸入應生產數量，並修改預期的製損率
![輸入數量](https://github.com/clin1155/VBA_Project_Demonstration/blob/ee07e944c17ae220d4da062f89e5cea961dea907/ScreenShot/Input.JPG)

本來需要幾個小時彙整運算，現在只需要不到5秒鐘，大幅縮短作業時間，同時降低重複加總和運算錯誤  
![結果報表](https://github.com/clin1155/VBA_Project_Demonstration/blob/ee07e944c17ae220d4da062f89e5cea961dea907/ScreenShot/Result.JPG)

# 預想未來還可以拓展的功能
+ 跨部門資料整合：若未來串接倉儲庫存資料，可同步顯示可用量、缺口量與建議採購量
+ 報表自動化：可將請購結果輸出為 PDF 或標準格式檔案，直接供採購部門使用
+ 版本管理與追蹤：保留每月計算結果，方便後續比對需求變化

# 備註
+ 如果要下載試用，必須在Excel檔按右鍵，選擇「內容」->「解除封鎖」，這是Windows避免VBA代碼中含有惡意程式的保護措施，否則程式無法運行
+ 實際上的資料庫量更大，本專案只截取部分資料展示核心運算邏輯
