## graph-client

このソリューションはSharePointのWebパーツにPlannerのタスクを表示するためのサンプルです。  
yo @microsoft/sharepoint コマンドで作成した雛形に対して、以下の変更を加えています。  
 * パッケージ追加(npm install @microsoft/microsoft-graph-types --save-dev)
 * config > pakage-solution.jsonファイルにwebApiPermissionRequestsを追加 (アクセス許可としてGroup.Read.Allが必要)
 * 参考：https://docs.microsoft.com/ja-jp/graph/api/planneruser-list-tasks?view=graph-rest-1.0&tabs=http
 * src > webparts > plannerWebpart > PlannerWebpartWebpart.tsファイルにコードを追加
 * 同フォルダ > PlannerWebpartWebpart.module.scssファイルにレイアウト用スタイルを追加

### ビルド方法

* PlannerWebpartフォルダをVisual Studio Codeで開く
* ターミナルで以下コマンドを順次実行
* npm i
* gulp build --ship
* gulp bundle --ship
* gulp package-solution --ship
* sharepointフォルダ > solutionフォルダ > planner-webpart.sppkgが出来れば成功

### デプロイ方法

* ビルド方法に従い作成したplanner-webpart.sppkgをSharePointのアプリカタログサイトにアップロード
* エラーが無く、展開済であることを確認
* SharePoint管理センター > APIの管理 画面で、Microsoft Graphのアクセス許可を承認
* 任意のSharePointサイトでアプリを追加(アプリ名：planner-webpart-client-side-solution)
* 同サイトの任意のページにWebパーツを追加(Webパーツ名：PlannerWebpart)