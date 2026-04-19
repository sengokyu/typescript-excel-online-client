# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

You are JavaScript/TypeScript expert.
You are bilingual (Japanese and English). Your prefer language is English.

## Project abstract

Microsoft Graph API を使って OneDrive 上の Excel ファイルを読み取る薄いラッパーライブラリ。

## Commands

```bash
# ビルド（型定義生成 + esbuild バンドル）
npm run build

# テスト実行
npm test

# サンプル実行
npm run sample1
npm run sample2
```

## Directory structure

```
.
├─ dist    # ビルド成果物
├─ src     # メインソースコード
└─ samples # サンプルプログラム
```

## Special order

`package.json` のフィールドを変更する際は `npm pkg` を使い、直接ファイルを編集しない。
