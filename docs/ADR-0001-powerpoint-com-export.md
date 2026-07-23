# ADR-0001: PowerPoint COM Automation による高解像度 PNG 書き出し

- Status: Accepted
- Date: 2026-07-23

## Context

Windows 版 Microsoft PowerPoint では、通常の PNG 書き出し UI から出力ピクセルサイズを自由に指定できない。

16:9 の PowerPoint スライドを Web コンテンツなどで利用するため、次のような高解像度 PNG 書き出しが必要である。

- 基準サイズ: 1920 × 1080
- 2 倍書き出し: 3840 × 2160
- 任意倍率による書き出し
- 元スライドと同じ縦横比の維持
- 複数スライドの一括処理

利用者の Windows PC には、デスクトップ版 Microsoft PowerPoint がインストールされている前提が置ける。

## Decision

PowerShell から Microsoft PowerPoint の COM Automation API を操作し、各スライドを PNG として書き出す CLI スクリプト `Export-PptxPng.ps1` を実装し、GitHub リポジトリ **ppt2png** として配布する。

PPTX ファイルの構造や描画内容は独自解析しない。レンダリングはインストール済み PowerPoint 自身に委譲し、PowerPoint の `Slide.Export` メソッドへ出力幅・高さをピクセル単位で渡す。

利用例:

```powershell
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2
.\Export-PptxPng.ps1 .\presentation.pptx -Width 3840
.\Export-PptxPng.ps1 .\presentation.pptx -Width 3840 -Height 2160
```

### 1. PowerPoint と同じレンダリング結果を得る

LibreOffice などの互換ソフトではなく、PowerPoint 本体をレンダラーとして利用する。これにより次の互換性問題を避ける。

- フォント配置の差
- テキスト折り返しの差
- SmartArt やグラフの差
- SVG、影、透過、テーマの差
- PowerPoint 固有機能の描画差

ただし、利用 PC に存在しないフォントや外部リンク画像などは、PowerPoint 上でも正確に再現できない。

### 2. PPTX レンダラーを独自実装しない

PPTX は単純な画像コンテナではない。独自レンダリングを行うと DrawingML、テーマ／マスタースライド、フォントメトリクス、SmartArt、グラフ、SVG、影・透過・グラデーション、Office 固有の互換処理などの再実装が必要になり、目的に対して実装範囲が過大になる。

### 3. PowerShell 単一スクリプトとして提供する

Windows 標準環境で利用しやすく、追加ランタイムなしで COM Automation を利用できるため PowerShell を採用する。初期バージョンではモジュール化や GUI 化は行わず、単一の `.ps1` として提供する。

### 4. COM lifecycle policy

スクリプトは、自身が生成した PowerPoint Application インスタンスだけを終了する。既にユーザーが起動している PowerPoint へ接続してはならない。

処理は必ず `try/finally` 相当の構造で実装し、成功・失敗を問わず以下を実行する。

1. Presentation を保存せず閉じる
2. PowerPoint Application を終了する
3. COM オブジェクトを逆順で解放する
4. 変数参照を破棄する
5. 必要に応じて GC を実行する

既存の POWERPNT.EXE を名前指定で一括終了してはならない。`Stop-Process -Name POWERPNT` のような処理は禁止する。ユーザーが編集中の PowerPoint まで終了する危険があるためである。

### 5. Security policy

- 入力ファイルは読み取り専用で開く
- マクロを実行しない
- 入力ファイルを上書きしない
- 出力先以外へファイルを書き込まない
- PowerPoint ファイル内の会社名・製品名・顧客情報が PNG へ含まれる可能性を README で警告する
- 機密資料を公開 GitHub リポジトリのテストデータとしてコミットしない
- テスト用 PPTX は架空の内容だけで作成する
- 出力 PNG にも元資料と同じ情報分類を適用する

### 6. Operational policy

本ツールは、ログイン済み Windows ユーザーが対話セッション内で実行するローカル CLI とする。以下はサポート対象外とする。

- Windows サービスからの実行
- GitHub Actions 上での PowerPoint 実行
- 複数ファイルの並列変換
- PowerPoint がインストールされていない環境
- パスワード保護されたファイル
- 修復確認ダイアログが必要な破損ファイル
- マクロや外部リンクの実行
- アニメーション途中状態の画像化
- 動画フレームの書き出し

### 7. Distribution policy

GitHub リポジトリ (ppt2png) では、スクリプトとドキュメントのみを管理する。リポジトリは public とし、ライセンスは MIT とする。

配布はタグ + GitHub Release で行う。`Export-PptxPng.ps1` を Release asset として添付し、利用者は Releases ページまたは `gh release download` で取得する。`main` ブランチ raw URL の直接ダウンロードは案内しない (バージョン固定できず、利用者側の挙動が予告なく変わるため)。

初期リリースの提供物:

- `Export-PptxPng.ps1`
- `README.md`
- `LICENSE` (MIT)
- `docs/ADR-0001-powerpoint-com-export.md`
- `docs/SPEC-0001-pptx-png-export.md`
- `tests/`

リポジトリへの push、タグ作成、Release 公開は、人間が内容を確認して実施する。

## Alternatives Considered

- **Windows レジストリの ExportBitmapResolution 変更**: 不採用。PC 単位の恒久設定になりファイルごとに倍率を変更しにくい。DPI とスライド物理サイズからピクセル数を逆算する必要があり、最終出力サイズを直感的に指定できない。チーム PC の設定状態に依存する。単一固定解像度を常用する個人環境では有効だが、チーム配布ツールとしては扱いにくい。
- **VBA マクロまたは PowerPoint アドイン**: 不採用。PPTX/PPTM への埋め込みやアドイン配布・署名・マクロポリシー対応が必要になり、PowerPoint ファイルそのものへツールの都合を持ち込む。CLI や自動処理として利用しにくい。
- **LibreOffice CLI**: 不採用。PowerPoint と異なるレンダリングエンジンでありレイアウト互換性を保証できない。PowerPoint が利用可能という前提を活かせない。
- **Aspose.Slides などの商用ライブラリ**: 不採用。ライセンス費用が発生し、ローカル変換用途に対して依存関係と導入負担が大きい。PowerPoint をインストールできないサーバー環境での大量変換が必要になった場合は再検討する。
- **C# による COM Automation**: 初期バージョンでは不採用。PowerShell 版で運用上の限界が確認された場合の移行候補とする。移行を検討する条件: 実行ポリシーが配布上の障害になる / 単一 EXE として配布したい / Explorer 右クリックメニューへ統合したい / GUI が必要になる / エラー処理・プロセス制御が PowerShell では複雑になる。

## Consequences

### 良くなること

- 任意の出力倍率・任意のピクセル幅を指定できる
- PowerPoint 本体と同じレンダリング結果を得られる
- 各 PPTX へのマクロ埋め込みが不要
- レジストリ変更が不要
- GitHub で単一スクリプトとして配布でき、CI やビルド工程を必要としない

### リスク・コスト

- Windows 専用となり、デスクトップ版 PowerPoint が必要になる
- PowerPoint COM Automation に依存する
- 非対話サーバーや Windows サービス上での実行は対象外となる
- PowerPoint のダイアログやアドインが処理を妨げる可能性がある
- 異常終了時に POWERPNT.EXE が残留する可能性がある
- 同時並列実行には適さない

## Migration Notes

新規リポジトリのため既存実装への影響はない。将来 C# 単一 EXE へ移行する場合は、本 ADR の「C# による COM Automation」の条件を満たしたことを確認し、別 ADR を作成する。

## Open Questions

- なし (未着手事項は `docs/BACKLOG.md` を参照)

## Progress

- 2026-07-23: ADR 作成・承認。仕様は [SPEC-0001](SPEC-0001-pptx-png-export.md) に記録

## References

- Microsoft Learn: Slide.Export method
- Microsoft Learn: Presentations.Open method
- Microsoft Learn: Creating .NET and COM objects in PowerShell
