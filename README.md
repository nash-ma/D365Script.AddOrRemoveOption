# D365Script.AddOrRemoveOption
## オプションセットコントロールにオプションの追加とリムーブ（adds an option to a control or removes an option from a control）
## ■変更前：スクリプトを追加した前に項目「案件ステータス」のオプションは「A、B、C、D、E」があります。👇
![スクリプト追加前：受注済がはい](WebSite/image/image1.png "スクリプト追加前：受注済がはい")

![スクリプト追加前：受注済がいいえ](WebSite/image/image2.png "スクリプト追加前：受注済がいいえ")

## ■変更後：スクリプトを追加した後に案件ステータスのオプションは「受注済」によって動的に変わります。👇
##### ①「受注済」が”はい”の時、AとBしか表示されない。
##### ②「受注済」が”いいえ”の時、C、D、Eしか表示されない。

![スクリプト追加後：受注済がはい](WebSite/image/image3.png "スクリプト追加後：受注済がはい")

![スクリプト追加後：受注済がいいえ](WebSite/image/image4.png "スクリプト追加前：受注済がいいえ")

### ■参考：
##### [addOption (Client API reference)](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/controls/addoption)
##### [removeOption (Client API reference)](https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/controls/removeoption)

