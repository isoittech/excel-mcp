# EXCEL 操作用 Python MCP サーバー

このリポジトリは、Java で実装した Excel 操作ツール群を Python からラップし、
Microsoft Copilot / Claude などの MCP クライアントから利用できる MCP サーバーを提供します。

Excel 操作のコア部分は Java + Apache POI で実装しています。
その上に Python で MCP サーバーを構築することで、次のような利点があります。

- Excel 操作は Apache POI が非常に成熟しており、細かい機能や互換性の面で有利
- Python 製の MCP サーバーは実装が容易で、fastmcp などのライブラリを活用しやすい

Python から Java を呼び出すラッパーを用意することで、
「Excel 操作は Java/Apache POI に任せつつ、MCP サーバー自体は Python で軽量に構築する」
という構成を実現しています。

## 目次

1. [システム要件](#システム要件)
2. [インストール](#インストール)
3. [MCP クライアントとの連携](#mcp-クライアントとの連携)
4. [主なツール一覧](#主なツール一覧)
5. [注意事項](#注意事項)

## システム要件

- Java 11 以上
- Python 3.11 以上
- Excel ファイル形式: .xlsx
- 対応 OS: Windows 10/11, macOS 12+, Linux (Ubuntu 20.04+)

## インストール

### 1. リポジトリのクローン

```bash
git clone git@github.com:isoittech/excel-mcp.git
cd excel-mcp
```

### 2. Java ツールのビルド

Java 側の Excel ツールは、`java/tools/compile.sh` でコンパイルできます:

```bash
cd java
./tools/compile.sh
cd ..
```

### 3. Python MCP サーバーのセットアップ

Python 版 MCP サーバーは `py` ディレクトリにあります。
`uv` を使ってローカル環境から起動することを想定しています。

```bash
cd py
uv run excel-mcp-server --help
```

初回実行時に、`pyproject.toml` に基づいて必要な依存関係がインストールされます。

## MCP クライアントとの連携

### Claude Desktop などから利用する場合

1. `excel-mcp/py` を MCP 設定ディレクトリとして指定します。
2. `py/mcp-config.json` に、MCP サーバーの定義が含まれています。

`mcp-config.json` の例:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "uv",
      "args": [
        "run",
        "excel-mcp-server"
      ],
      "env": {
        "PYTHONPATH": "./src"
      }
    }
  }
}
```

MCP クライアントを再起動すると、`excel-mcp-server` というサーバー名で
Excel 操作用の MCP ツール群が利用できるようになります。

### Docker を使う場合

`py` ディレクトリには Dockerfile / docker-compose.yml も用意しています。
SSE モードで HTTP 経由の接続を行いたい場合に利用できます。

```bash
cd py
docker build -t excel-mcp-server .
docker run --rm -p 8585:8585 excel-mcp-server excel-mcp-server -t sse -p 8585
```

`docker-compose.yml` を使う場合:

```bash
cd py
docker compose up -d
```

MCP クライアントからは、`http://host.docker.internal:8585/sse` などの SSE エンドポイントを指定します。

## 主なツール一覧

以下は、MCP クライアントから利用できる主なツールの例です。

### EXCEL ファイルの読み込み

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "read_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "range": "A1:C10"
  }
}
```

### EXCEL ファイルへの書き込み

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "write_excel",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "data": [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"]
    ]
  }
}
```

### 新しいシートの作成

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_sheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "NewSheet"
  }
}
```

### 新しい EXCEL ファイルの作成

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_excel",
  "arguments": {
    "filePath": "/path/to/new_file.xlsx",
    "sheetName": "Sheet1"  // 省略可、デフォルトは "Sheet1"
  }
}
```

### ワークブックのメタデータ取得

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "get_workbook_metadata",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "includeRanges": false  // 省略可、範囲情報を含めるかどうか
  }
}
```

### シート名の変更

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "rename_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "oldName": "Sheet1",
    "newName": "NewName"
  }
}
```

### シートの削除

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "delete_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1"
  }
}
```

### シートのコピー

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "copy_worksheet",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sourceSheet": "Sheet1",
    "targetSheet": "Sheet1Copy"
  }
}
```

### セルへの数式適用

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "apply_formula",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### 数式構文の検証

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "validate_formula_syntax",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "cell": "C1",
    "formula": "=SUM(A1:B1)"
  }
}
```

### セル範囲の書式設定

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "format_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "bold": true,
    "italic": false,
    "fontSize": 12,
    "fontColor": "#FF0000",
    "bgColor": "#FFFF00"
  }
}
```

### セルの結合

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "merge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### セルの結合解除

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "unmerge_cells",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C1"
  }
}
```

### セル範囲のコピー

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "copy_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "sourceStart": "A1",
    "sourceEnd": "C3",
    "targetStart": "D1",
    "targetSheet": "Sheet2"  // 省略可、省略時は同じシート
  }
}
```

### セル範囲の削除

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "delete_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3",
    "shiftDirection": "up"  // "up" または "left"、デフォルトは "up"
  }
}
```

### Excel 範囲の検証

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "validate_excel_range",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "startCell": "A1",
    "endCell": "C3"  // 省略可
  }
}
```

### グラフの作成

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_chart",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:C10",
    "chartType": "column",  // "column", "line", "bar", "area", "scatter", "pie"
    "targetCell": "E1",
    "title": "サンプルグラフ",  // 省略可
    "xAxis": "X軸ラベル",      // 省略可
    "yAxis": "Y軸ラベル"       // 省略可
  }
}
```

### ピボットテーブルの作成

```json
{
  "server_name": "excel-mcp-server",
  "tool_name": "create_pivot_table",
  "arguments": {
    "filePath": "/path/to/file.xlsx",
    "sheetName": "Sheet1",
    "dataRange": "A1:D100",
    "rows": ["Category"],
    "values": ["Sales"],
    "columns": ["Region"],  // 省略可
    "aggFunc": "sum"  // "sum", "count", "average", "max", "min" など
  }
}
```

## 注意事項

- ファイルパスは絶対パスで指定すること
- シート名を指定しない場合、最初のシートが対象になる
- 範囲指定は "A1:C10" のような形式で記述する
- create_excel で既存ファイルパスを指定するとエラーになる
- ピボットテーブル機能は、現在はメタデータ構築のみで実際の Excel ピボットテーブルオブジェクトは作成しない実装になっている場合があります

## 作者

- 作者: isoittech

## ライセンス

この MCP サーバーは MIT ライセンスで提供されています。
自由に利用・改変・再配布できます。詳細はリポジトリ内の LICENSE ファイルを参照してください。


## 注意事項

- ファイルパスは絶対パスで指定すること
- シート名を指定しない場合、最初のシートが対象になる
- 範囲指定は "A1:C10" のような形式で記述する
- create_excel で既存ファイルパスを指定するとエラーになる
- ピボットテーブル機能は、現在はメタデータ構築のみで実際の Excel ピボットテーブルオブジェクトは作成しない実装になっている場合があります
