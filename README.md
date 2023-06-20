# makeMusicFormat
## 画面(例)
![image](https://github.com/tadashiro4610/makeMusicFormat/assets/61111774/39beb9c5-a9ef-480b-b42d-09736b9039de)
## 注意
* **<font color="Red">ただの素人が作ったやつなので自己責任で</font>**
* excelが必要

## ダウンロード
* すぐ起動してほしい人は「app」フォルダをダウンロード
* 軽くてファイル数が小さいのがいい人は「onlyexe」の「makeMusicFormat.exe」をダウンロード(名前は分かりやすいのに変えて下さい)
* ソースコード(python)が欲しい人は「source」をダウンロード

## 使い方
* 曲の題名，一行に表示する小節数，セクション名(Aメロ,Bメロとか)，小節数を入力すると自動でexcelのシートに生成してくれます
* 後は幅とか調整して印刷するだけ

## 例
### 入力
* 題名とか色々入力する
* セクション入力はプルダウンから選んでも直接入力でも可
![image](https://github.com/tadashiro4610/makeMusicFormat/assets/61111774/67f4325e-912c-43f4-a0be-f1bd80c2c1d8)
* 小節数も矢印でも直接入力でも可
* 色のボタンを押すとカラーピッカーが開く
* 「Random」ボタンで色をランダムで選択される
* 右下の「+」ボタンを押すとリストに追加されていく(消すときは消したいセクションを選択して「-」ボタン)
* リストで選択すると、「セクション名」「小節数」「色」が自動で入力される
![md1](https://github.com/tadashiro4610/makeMusicFormat/assets/61111774/af9bde16-c89e-43c3-b3a4-95cee57c29f4)
* 選択の解除は右下の「選択解除」ボタンで可能
### 保存
* 保存は下の「参照」ボタンを押して，保存ファイルを選択または作成
![md2](https://github.com/tadashiro4610/makeMusicFormat/assets/61111774/bc456ed4-64e1-408e-9b81-87ad8cdeba16)


### excelシートの生成
* 左下の「generate」ボタンを押してexcelシートに焼き付ける
* 生成されたexcelシートは以下のようになる
![image](https://github.com/tadashiro4610/makeMusicFormat/assets/61111774/87055412-c93e-490f-8c8b-5e9f5463481f)

### 印刷する
* 印刷プレビューを見ながら印刷する
* 印刷範囲を「選択した部分を印刷」にして画面の「余白を表示」でいじるとやりやすい(ご自由に)
![image](https://github.com/tadashiro4610/makeMusicFormat/assets/61111774/b11aeb71-3cb1-46f1-abbd-da143dd0645d)

### 今後
* 暇だったら色々追加するかも

### その他
* コードは勝手にいじってください
* appフォルダ内の「pullDownList.csv」をいじれば、セクション名の候補を変えられる