# TASK-260723: README 整備と配布 (publish) 方針の確定

## 目的
利用者向け README (Installation / Usage) を作成し、単一 `.ps1` スクリプトの配布チャネルと publish 手順を確定する。

## 決定事項
- リポジトリは **public**
- 利用者の導入方法は **GitHub Release asset** (Releases ページから手動 DL、または `gh release download`)
    - `main` raw URL 直 DL は案内しない (バージョン固定不可のため)
- publish 手順は **タグ + GitHub Release 作成** (asset として `Export-PptxPng.ps1` を添付)。実行は人間のみ
- README 構成: 利用者向け Getting Started (Requirements / Installation / Usage / Known limitations) を先頭に、Structure / Depends は Development 配下へ移動
- ブラウザ DL 時の MOTW 対策として `Unblock-File`、実行ポリシーはプロセス単位 Bypass を案内 (SPEC §Trouble Shooting 準拠)

## 成果物
- `README.md` (全面書き換え。SPEC-0001 README requirements の全項目 + 機密情報警告を反映)
- `docs/ADR-0001-powerpoint-com-export.md` — Distribution policy へ配布チャネル (public / Release asset / raw 直 DL 非案内) を追記
- `docs/BACKLOG.md` — README 項目を完了として削除、初回リリース作業 (人間) を追加

## 残項目
- BACKLOG へ集約済み (本体実装、LICENSE、初回 Release 作成など)
