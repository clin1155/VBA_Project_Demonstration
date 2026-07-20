# VBA專案名稱
生物科技原料採購自動化報表

# 問題概述
生管部門每月要計算次月的原料需求量，要生產的產品上百種，每種產品的配料表有10-20項不等的原料。過去的做法，是先把每張訂單都分別算完需求量，再把其中相同供應商且重複的原料進行加總(例如不同口味的乳清蛋白粉，可能共同使用同一款乳清蛋白)，最後再依照彙整好的資料開單請購。依據生管人員表示，因為資料龐雜，他時常必須假日加班，才不會耽誤購買原料的時程。希望我可以協助。

# 功能摘要
讓使用者在產品列表裡面輸入應生產數量，並讓使用者依照過去經驗填入預期的製損率，一鍵執行程式，直接得到包含以下內容的報表：
+ 原料的指定供應商
+ 原料的應採購量
+ 本採購所對應的訂單(例如總共需要買50KG的香料粉，這50KG分別為了那些訂單採購的)

# 面對挑戰，找到解法
考慮到未來資料大量增加的時候，依然保持程式的執行效率，我決定採用Dictionary，但是無法達到我預期的效果。我把問題交給ChatGPT、CoPilot、DeepAI等模型，最後決定採用雙層Dictionary + Array的方式來完成，代碼最核心的邏輯如下：


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
如下圖輸入應生產數量，並調整預期的製損率
![輸入數量](https://github.com/clin1155/VBA_Project_Demonstration/blob/ee07e944c17ae220d4da062f89e5cea961dea907/ScreenShot/Input.JPG)

不用5秒鐘，即可得到最終報表--
![結果報表](https://github.com/clin1155/VBA_Project_Demonstration/blob/ee07e944c17ae220d4da062f89e5cea961dea907/ScreenShot/Result.JPG)

# 預想未來還可以拓展的功能
+ 寫這個程式的時候，公司的倉儲資料尚未建置，若導入倉儲部門的資料，可以使程式再進一步跨部門整合，使結果更有參考價值
+ 如果採購單的表格形式完成，可以直接生成多張PDF文件報表，進一步大大提高開單效率

# 備註
+ 如果要下載試用，必須在Excel檔按右鍵，選擇「內容」->「解除封鎖」，這是Windows避免VBA代碼中含有惡意程式的保護措施，否則程式無法運行
+ 實際上的資料庫量更大，本專案只截取部分資料展示核心運算邏輯
