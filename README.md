### 演算法說明 ###
---
Floyd’s Algorithm 是一種各點間尋找最短路徑的演算法。它利用兩個矩陣進行計算，一是距離矩陣 (Distance Matrix)；另一是節點矩陣 (Node Matrix)。其中距離矩陣用來表示某點 A 到某點 B 的距離，節點矩陣則用來表示某點 A 到某點 B 間須經過的其他節點。

舉個例子實作一次，假設現有 5 個節點尋找最短路徑，節點路徑圖如下所示

<p Align=center><img src=/img/exorigpic.png alt=範例原始圖片></p>

則距離矩陣與節點矩陣可以表示如下

<p Align=center><img src=/img/exormatpic.png alt=範例初始矩陣圖片></p>

在距離矩陣中，無法一次抵達的節點，以無窮的符號標示，並且節點到節點本身不計。而在節點矩陣中，第一行 (column) 為 1，第二行為 2 ...，以此類推，並且節點到節點本身不計。

接著取第 k 列 (row)、第 k 行作為基準 (k 從 1 開始)，檢查第 i 列、第 j 行的元素 (簡寫為 d_ij，i 與 j != k) 是否大於檢查行列之和，也就是檢查是否：d_ij > d_kj + d_ik；若是，則下一個矩陣的 d'_ij = d_kj + d_ik，並且下一個節點矩陣的 n'_ij = k。檢查完距離矩陣內所有元素後，k + 1 再繼續檢查，以本範例來說需要檢查到 k = 5 後才結束。

完成演算後的結果如下

<p Align=center><img src=/img/exrematpic.png alt=範例最終矩陣圖片></p>

現在就可以根據上面的矩陣回答下列問題：

1. 試問點 A 到 B 的距離為？需要經過哪些節點？

2. 試問點 A 到 D 的距離為？需要經過哪些節點？

3. 試問點 B 到 E 的距離為？需要經過哪些節點？

解答如下：

1. 距離為 5 (從距離矩陣查得)，移動一次即可到達點 B (節點矩陣中的數字 1 代表 A、2 代表 B ...，以此類推)。

2. 距離為 12，移動的點依序為 A、C、D。先查節點矩陣得知 A 到 D 要先經過 C，再查 A 到 C 發現為直達，因此答案為 A、C、D。

3. 距離為 23，移動順序為 B、A、C、D、E。

### 程式介紹 ###
---
使用者只需要將距離矩陣填入 Excel 表單中，再利用本程式的 RefEdit 選取距離矩陣所在的 Range 即可。

自訂表單樣板如下

<p Align=center><img src=/img/userform.jpg alt=自訂表單圖片></p>

選取距離矩陣所在的 Range，須注意先前以無窮符號表示的地方，現在需要用一個很大的數取代，如下圖

<p Align=center><img src=/img/ormatpic.png alt=初始矩陣圖片></p>

按下 OK 按鈕後即可得到計算過程與最終的解答，如下圖所示

<p Align=center><img src=/img/rematpic.png alt=最終矩陣圖片></p>
