# jxls-sample

## 概要

Jxls を利用した帳票出力処理のサンプル実装。

## 実行手順

以下のコマンドを実行すると、 ".output" ディレクトリに PDF ファイルが出力される。

```bash
mvn clean package

mvn exec:java -Dexec.mainClass="com.qwerty0121.jxls.sample.JxlsSample"
```
