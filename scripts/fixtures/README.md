# Power Automate の互換性フィクスチャ

`microsoft-bypass-flow.pre-attachments.json` は、コミット `0da7009` に由来する、秘密情報を除いた元の 3 操作のエクスポートです。**変更しない履歴資料・生成器の基準入力**であり、別フローや配備用パッケージではありません。

現在の唯一の正規ソースは [`power-automate/microsoft-bypass-flow/definition.json`](../../power-automate/microsoft-bypass-flow/definition.json) です。[生成・更新の説明](../../power-automate/microsoft-bypass-flow/README.md)も参照してください。

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
bun run test:flow
```

先頭コマンドは、このフィクスチャに固定操作の拡張と actions の互換変換を加え、正規ソースへ書き出します。`--check` はファイルを書き換えず、生成結果との一致を確認します。

テストはフィクスチャの checksum、既存の認証・パラメーター・default 分岐の保持、元の添付分岐と互換変換、現行の決定的な生成結果を検証します。メール分岐は現行仕様に更新されているため、「生成結果から添付 2 分岐を除けば元フィクスチャと完全一致する」という古い説明は、現在の契約には当てはまりません。

生成定義でこのフィクスチャを上書きしたり、生のエクスポートで更新したりしないでください。認証情報・接続情報・実 URL を持つ非公開エクスポートをここへ追加してはいけません。オフラインテストは合成データだけを使い、Microsoft 側の import / 保存 / 実行は確認しません。
