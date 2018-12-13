# スプレッドシート to JSON

## 機能
アクティブなスプレッドシートからJSONファイルを作成する。  

## 準備
保存先のドライブのディレクトリのidを調べ、
`var folder = DriveApp.getFolderById('XXXXXXXXXXXXXXXXXXXXXXXXXXX');`
の部分を置き換える。

## シートの種類
シート名の後ろに`?パラメータ`をつけることでデータの形式を選ぶことができる。

### デフォルト(sheetName パラメータなし)
オブジェクトの配列

|key_1|key_2|...|key_n|
|---|---|---|---|
|value_01|value_02|...|value_0n|
|value_11|value_12|...|value_1n|
|...|...|...|...|
|value_n1|value_n1|...|value_nn|

1行目はキー。
```
[
  {
    key_1: value_01,
    key_2: value_02,
    ...
    key_n: value_0n
  },
  ...
  {
    key_1: value_n1,
    key_2: value_n2,
    ...
    key_n: value_nn
  }
]
```

### Array(sheetName?Array)
配列

|A列に値を埋める|
|---|
|value_0|
|value_1|
|...|
|value_n|

1行目から埋めていく。

```
[
  value_0,
  value_1,
  ...
  value_n
]
```

### Object(sheetName?Object)
|A列はキー|B列は値|
|---|---|
|key_0|value_0|
|key_1|value_1|
|...|...|
|key_n|value_n|

1行目から埋めていく。
```
{
  key_0: value_0,
  key_1: value_1,
  ...
  key_n: value_n
}
```

## 値の埋込み
{{#シート名}}とすることで、別のシートのオブジェクトを埋め込むことができる。

### sheet: log?Objcet
grade, 1  
score, {{#score}}

### sheet: score?Array
90,  
87,  
53,  
69

log?ObjcetシートでGASを実行すると次のJSONがドライブに作成される。

### log.json
```
{
  "grade": 1,
  "score": [
    90,
    87,
    53,
    69
  ]
}
```
