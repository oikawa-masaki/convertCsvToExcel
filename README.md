## 概要

CLI から指定したディレクリ配下の CSV ファイル(UTF-8)を再帰的に 1 つの Excel (.xlsx)に整理します

## 実行環境

- MacBook Pro(Apple M2)
- macOS Sonoma 14.2.1

```shell
% volta list
⚡️ Currently active tools:

    Node: v18.18.0 (default)
    Yarn: v1.22.19 (default)
    Tool binaries available:
        vue (default)
        cdk (default)
        npm-check-updates, ncu (default)
        tsc, tsserver (default)
```

## 使用方法

### npm install

```shell
npm install # or npm i
```

### Build（tsc transpile）

```shell
tsc
```

### 実行

```shell
node src/convertCsvToExcel.js <csv_directory_path> <output_excel_file_path>

# <csv_directory_path> : 配下に CSV ファイルを含んでいるディレクトリを指定してください
# <output_excel_file_path>: 1つのExcel ファイルとして出力 Excel ファイル名（拡張子は".xlsx"）を指定してください
```

### 実行例

```shell
node src/convertCsvToExcel.js /csv_input /csv_output/output.xlsx
```

### 注意事項

- コマンド実行時は必ず引数を 2 つ入力(指定)してください
  - 上記以外だと以下のエラーが出力されます
  - "Usage: node convertCsvToExcel.js <csv_directory_path> <output_excel_file_path>"
- 第一引数に指定するディレクトリは、配下に CSV ファイルが含まれているディレクトリを指定してください
