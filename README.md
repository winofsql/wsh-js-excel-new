# wsh-js-excel-new
```javascript
    "code-runner.executorMapByFileExtension": {
        ".wsf": "cscript //Nologo"
    }
```
ブックを作成してセルに値をセットして Excel で起動
```xml
<?xml version="1.0" encoding="utf-8" ?>
```
GitHub 用に utf-8 で記述する為に wsf 形式を使用
```javascript
var App = new ActiveXObject( "Excel.Application" );
var WshShell = new ActiveXObject( "WScript.Shell" );
```
HTML アプリケーションでもコピペで使いたいので new ActiveXObject を使用

## 重要
```javascript
// セルに値をセット
Book.Sheets(1).Cells(1, 1).Value = "社員名";
Book.Sheets(1).Cells(2, 1).Value = "山田　太郎甚左衛門";
Book.Sheets(1).Cells(3, 1).Value = "鈴木　一郎";
Book.Sheets(1).Cells(4, 1).Value = "佐藤　洋子";
```
