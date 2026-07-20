# WinUI 3 アプリ アーキテクチャガイド

このドキュメントは、BlueShift / SmartPowerManager で確立した WinUI 3（Windows App SDK, unpackaged）アプリのアーキテクチャを整理したものです。
このテンプレート（`WinUiAppTemplate`）はこのアーキテクチャに従って実装されています。**新しい C# WinUI 3 アプリを作る際は、このテンプレートをコピーし、本ドキュメントのルールを維持したまま拡張してください。**

## 1. 全体像（プロセス寿命とウィンドウ寿命の分離）

最も重要な原則は「**プロセスの寿命**」と「**MainWindow の寿命**」を分離することです。

```
Program.Main
  └─ App (Application)                 … プロセス全体で 1 個。DispatcherShutdownMode = OnExplicitShutdown
       └─ AppRuntime                    … プロセスが生きている間ずっと 1 個。トレイ・二重起動リスナーを所有
            └─ MainWindow（0 個か 1 個） … 表示するたびに new、閉じるたびに破棄（ADM = Auto Dark Mode 方式）
```

- **AppRuntime** がプロセスの実体（トレイアイコン・バックグラウンドサービス・二重起動時の Show/Exit イベント監視）を持ちます。
- **MainWindow** は「今画面に表示されているウィンドウ」でしかありません。閉じたら本当に破棄し、再度開くときは new して作り直します。
- ユーザーが × を押しても **プロセスは終了しません**（トレイを使う場合）。ExitApplication が呼ばれるまで `AppRuntime` は生き続けます。

### なぜこの構造にするか

- WinUI の `Window` はいちど Close すると内部状態が壊れやすく、同じインスタンスを Hide/Show して再利用すると白画面・入力不能などの不具合が起きやすい。
- タスクトレイ常駐アプリでは「ウィンドウを閉じる」と「アプリを終了する」は別の操作。この区別をコード上でも明確に分離する。

## 2. 各コンポーネントの責務

### App（`App.xaml.cs`）

- `DispatcherShutdownMode = DispatcherShutdownMode.OnExplicitShutdown` を **必ず** コンストラクタで設定する。
  - これを忘れると、最後の `Window` が閉じた瞬間にプロセスが終了してしまい、トレイ常駐が機能しない。
- `OnLaunched` の役割は最小限にする:
  1. `UpdateChecker.LatestReleaseApiUrl` を設定する（未リリースなら `null` のままでよい）。
  2. `Settings.Load()` で設定を読み込む。
  3. `--background` 引数を見て `launchInBackground` / `requestInteractiveShow` を決める。
  4. `SingleInstanceManager.TryBecomePrimaryInstance(...)` で二重起動をチェックする。二重起動なら既存インスタンスに表示を依頼して `Exit()`。
  5. `AppRuntime` を生成して `Start(...)` を呼ぶ。
- ドメインロジック（バックグラウンド処理・トレイ操作など）は **App に書かない**。すべて `AppRuntime` に委譲する。

### AppRuntime（`AppRuntime.cs`）

プロセス寿命そのものを表すクラス。責務は次の 4 つだけ：

1. **タスクトレイの所有**（`TrayMessageWindow` / `TrayIconService`）
2. **二重起動イベントの監視**（`SingleInstanceManager.InteractiveShowEvent` / `ExitEvent` を別スレッドで待機）
3. **MainWindow の生成・表示・破棄の管理**（`ShowOrCreateMainWindow` / `OnMainWindowClosing` / `MainWindow_Closed`）
4. **アプリ終了処理の一本化**（`ExitApplication()`）

バックグラウンドサービス（スケジューラ・監視タイマーなど）を追加する場合も、この `AppRuntime` が所有し、起動・停止のタイミングを管理します（BlueShift のガンマ制御、SmartPowerManager のスケジュール実行サービスが実例）。

```csharp
public void ShowOrCreateMainWindow(string? pageTag = null)
{
    if (_isExitingProcess) return;
    GetDispatcherQueue()?.TryEnqueue(() => ShowOrCreateMainWindowCore(pageTag));
}

private void ShowOrCreateMainWindowCore(string? pageTag = null)
{
    if (_mainWindow != null)
    {
        BringWindowToForeground(_mainWindow);
        if (pageTag != null) _mainWindow.NavigateToPageTag(pageTag);
        return;
    }

    _mainWindow = new MainWindow(this);
    _mainWindow.Closed += MainWindow_Closed;
    _mainWindow.PrepareAndActivate(pageTag);
}
```

`ExitApplication()` はアプリを完全終了する **唯一の入口** です。トレイの「終了」メニューも、`--exit` コマンドライン引数（インストーラからの終了依頼）も、最終的にこの 1 メソッドだけを呼びます。

```csharp
public void ExitApplication()
{
    if (_isExitingProcess) return;
    _isExitingProcess = true;

    // リスナー停止 → トレイ破棄 → SingleInstanceManager.Release() → MainWindow.Close() → _app.Exit()
}
```

### MainWindow（`MainWindow.xaml.cs`）

MainWindow は **UI 専業** です。プロセス寿命には関与しません。

- コンストラクタで `AppRuntime` を受け取り、`AppState`（設定など共有状態）を取得する。
- `PrepareAndActivate(pageTag)` の順序を厳守する:
  1. `NavigateToPage(...)`（最初のページへ遷移）
  2. `RestoreWindowBounds()`（`WindowPlacementHelper` で位置・サイズ・最大化状態を復元）
  3. `Activate()`
  - この順序は Auto Dark Mode と同じで、ちらつきや位置ズレを防ぎます。
- **× ボタン（`AppWindow.Closing`）は本当に `Close` させます。** `args.Cancel = true` にして `Hide()` するような実装は禁止です（詳細は 4 章）。
- ウィンドウの位置・サイズは `AppWindow.Changed` で変化を検知し、`WindowPlacementHelper.Save` で `Settings` に保存します。

### TitleBarThemeHelper（`TitleBarThemeHelper.cs`）

タイトルバーのテーマ同期は **3行だけ** で完結させます。

```csharp
_window.ExtendsContentIntoTitleBar = true;
_window.SetTitleBar(_titleBar);   // titleBar は XAML 上の <TitleBar x:Name="AppTitleBar">
_window.AppWindow.TitleBar.ButtonHoverBackgroundColor = hover; // テーマに応じた半透明色のみ
```

- MainWindow.xaml 側は `Grid.RowDefinitions="Auto,*"` の 1 行目に `<TitleBar x:Name="AppTitleBar">` を置くだけ。カスタム Grid で余白列を作ったり、キャプションボタンの座標を計算したりしない。
- `ButtonBackgroundColor` / `ButtonForegroundColor` / `ButtonPressedBackgroundColor` などを個別に設定する「キャプションボタンの手動色付け」は **行わない**。WinUI 3 の `TitleBar` コントロール（`ExtendsContentIntoTitleBar` 前提）はテーマに応じて自動的に描画されるため、必要なのはホバー色の半透明オーバーレイのみです。
- ウィンドウをオフスクリーンに一瞬 Move してから戻す、`ResetCaptionButtons` 的な再描画ハックは不要です（旧実装の名残があれば削除してください）。

## 3. タスクトレイとバックグラウンド起動

### タスクトレイ（`TrayMessageWindow.cs` / `TrayIconService.cs`）

- トレイ通知には **専用の非表示ウィンドウ**（`TrayMessageWindow`、Win32 の `CreateWindowEx` + `WS_POPUP`）を使います。MainWindow の `WndProc` を差し替えて `WM_TRAYICON` を処理する実装は禁止です（WinUI 側の再描画が壊れて白画面になることがあるため）。
- `TrayIconService` は `Shell_NotifyIcon` による最小実装（ダブルクリックでメイン画面、右クリックで「設定を開く / 終了」の 2 項目メニュー）。サードパーティのトレイライブラリは必須ではありません。
- `Settings.HideTrayIcon` が true のときはトレイアイコンを非表示にできますが、`AppRuntime` 自体は常駐を続けます（トレイを消す ≠ 常駐をやめる）。

### バックグラウンド起動（`--background`）

- `--background` 引数付きで起動すると、`MainWindow` を表示せずトレイ・バックグラウンドサービスだけを起動します。
- 二重起動時、既存インスタンスに「画面を開いてほしい」と伝えるかどうかは `requestInteractiveShow`（= `!launchInBackground`）で判定します。`--background` の二重起動では通知しません。

### 終了経路の一本化（`--exit`）

- `Program.cs` で `--exit` を検出したら、**WinUI を起動せず** `SingleInstanceManager.SignalExit()` を呼んで既存プロセスに終了イベントを送り、即座に戻ります（インストーラのアンインストール前フックなどで使用）。
- 受け取った側は `AppRuntime` のリスナーが `ExitApplication()` を呼び、`ExitApplication → Application.Exit()` の 1 経路だけで終了します。トレイの「終了」メニューも同じ `ExitApplication()` を呼びます。

### 二重起動制御（`SingleInstanceManager.cs`）

- 名前付き `Mutex` で多重起動を検出し、名前付き `EventWaitHandle`（`InteractiveShowEvent` / `ExitEvent`）で既存インスタンスに「表示」「終了」を伝えます。
- 複製したプロジェクトでは `Mutex` / イベント名に含まれるアプリ名部分を変更し、他アプリと衝突しないようにしてください。

## 4. やってはいけないこと（アンチパターン）

過去の実装で問題が出た/回避した方法です。**このテンプレート・派生プロジェクトでは使用しないでください。**

| アンチパターン | 問題点 | 代わりにすること |
|---|---|---|
| × ボタンで `args.Cancel = true` にして `Window.Hide()` する | WinUI の `Window` は Hide/Show の再利用でレイアウト崩壊・入力不能になりやすい | 本当に `Close()` する。プロセス常駐は `AppRuntime` が担う |
| タイトルバーのキャプションボタン背景・前景色を個別に手動設定する | テーマ切替時に反映漏れが起きやすく、保守コストが高い | `ExtendsContentIntoTitleBar` + `SetTitleBar` + `ButtonHoverBackgroundColor` のみ |
| ウィンドウをオフスクリーンに `Move` してから戻す再描画ハック | 副作用が読みにくく、フリッカーの原因にもなる | 上記の TitleBar 方式にすれば不要 |
| MainWindow の `WndProc` を差し替えてトレイメッセージを処理する | WinUI の内部メッセージ処理と衝突し、白画面などの不具合が起きる | `TrayMessageWindow`（専用の非表示ウィンドウ）を使う |
| `App.Exit()` 相当の終了処理を複数箇所に書く | 終了漏れ・二重終了・トレイだけ残る等の不整合が起きる | `AppRuntime.ExitApplication()` に一本化する |
| 起動時に無条件でログオン自動起動を登録する | ユーザーの意図しない常駐増加、セキュリティ上の懸念 | 5 章参照。既定では実装しない |

## 5. オプション機能（既定では未実装）

以下はテンプレートに **意図的に含めていません**。プロジェクトの要件に応じて、必要になった時点で追加してください。

- **ログオン時自動起動（AutoStart / StartupManager）**
  - ⚠️ **ログオン時自動起動はユーザーから明示的な指示があるときのみ実装すること。** テンプレートやこのアーキテクチャに従う新規プロジェクトへ、AI エージェントが独断でレジストリ `Run` キー登録・タスクスケジューラのログオンタスク・`StartupManager` クラス等を追加することは禁止です。
  - 実装が必要になった場合の参考実装は BlueShift / SmartPowerManager の `StartupManager.cs`（タスクスケジューラのログオンタスク方式、`--sync-autostart` / `--cleanup-autostart` 引数、`Settings.AutoStart` との同期）にあります。本テンプレートにはこのクラスは含まれていません。
- **バックグラウンドサービス（スケジューラ・監視タイマー・外部デバイス連携など）**
  - `AppRuntime` に追加する形で実装します（`EnsureXxx()` / 対応する `Start` 呼び出し / `ExitApplication` での破棄、というパターンに揃えてください）。
- **インストーラ（Inno Setup 等）・自動更新の適用処理**
  - `UpdateChecker` は GitHub Releases の確認のみを行います。ダウンロード・適用（`UpdateInstallerService` 等）は必要になった時点で追加してください。

## 6. 設定・テーマ・アップデート確認

- **`Settings.cs`**: JSON（Newtonsoft.Json）で `%AppData%\<AppName>\settings.json` に保存。ウィンドウ位置・サイズ・最大化状態、テーマ設定、`HideTrayIcon` を保持します。
- **テーマ**: `AppThemePreference`（`System` / `Light` / `Dark`、既定値は `System`）。`ThemeService` がルート要素（`RootGrid`）へ反映し、`ThemeChanged` イベントで `TitleBarThemeHelper` 等に通知します。設定画面には必ず 3 択のテーマ切替 UI を用意してください。
- **アップデート確認**: `UpdateChecker.LatestReleaseApiUrl` に GitHub Releases API の URL（`https://api.github.com/repos/<owner>/<repo>/releases/latest`）を設定すると有効になります。未設定時は `NotConfigured` を返すだけで、エラーにはなりません。
- **多言語化**: `Strings/ja-JP/Resources.resw` と `Strings/en-US/Resources.resw` に日本語・英語の文言を用意し、`Strings.Get` / `Strings.Format` 経由で取得します（`ApplicationDefaultLanguage` は `ja-JP`）。新しい文言を追加したら両方の `.resw` に追記してください。

## 7. ファイル対応表

| ファイル | 役割 |
|---|---|
| `Program.cs` | エントリポイント。`--exit` の早期処理、WinUI 起動 |
| `App.xaml.cs` | `OnExplicitShutdown` 設定、二重起動チェック、`AppRuntime` 生成 |
| `AppRuntime.cs` | プロセス寿命・トレイ・リスナー・`ShowOrCreateMainWindow`・`ExitApplication` |
| `SingleInstanceManager.cs` | Mutex + 名前付きイベントによる二重起動制御 |
| `TrayMessageWindow.cs` / `TrayIconService.cs` | タスクトレイ（専用非表示ウィンドウ + Shell_NotifyIcon） |
| `MainWindow.xaml` / `.xaml.cs` | UI 専業のウィンドウ。TitleBar + NavigationView + Frame |
| `TitleBarThemeHelper.cs` | タイトルバーのテーマ同期（最小実装） |
| `WindowPlacementHelper.cs` | ウィンドウ位置・サイズ・最大化状態の保存／復元（WINDOWPLACEMENT） |
| `AppState.cs` | ページ間で共有する状態（`Settings` + トレイ更新フックなど） |
| `Settings.cs` | JSON 設定の読み書き |
| `ThemeService.cs` | ライト/ダーク/システム連動テーマの適用 |
| `UpdateChecker.cs` | GitHub Releases 経由のアップデート確認 |
| `Strings.cs` + `Strings/*/Resources.resw` | 多言語文字列リソース |
| `Views/HomePage.xaml*`, `Views/InfoPage.xaml*`, `Views/SettingsPage.xaml*` | 各ページ（Settings ページは必須） |
| `Views/PageScrollHost.xaml*` | ページ共通の必要時のみスクロールする ScrollViewer ラッパー |

## 8. 新規プロジェクトを始める手順

1. `WinUiAppTemplate` フォルダをコピーし、新しいプロジェクト名にリネームする。
2. `WinUiAppTemplate.csproj` の `RootNamespace` / `AssemblyName` / `Version` を変更し、全 `.cs` / `.xaml` の namespace（`WinUiAppTemplate`）を新しい名前に置換する。
3. `SingleInstanceManager.cs` / `TrayMessageWindow.cs` の Mutex 名・イベント名・ウィンドウクラス名を新しいアプリ名に変更する（他アプリと衝突させないため）。
4. `Strings/*/Resources.resw` の `AppName` 等を変更する。
5. `Assets/AppIcon.ico` を配置する（`MainWindow.xaml` の `TitleBar.IconSource`、`TrayIconService`、`MainWindow.ApplyWindowIcon` が自動的に参照する）。
6. `App.xaml.cs` の `UpdateChecker.LatestReleaseApiUrl` に GitHub Releases API URL を設定する。
7. 本アーキテクチャ（`AppRuntime` 中心の構造、× で Close、TitleBar 最小実装）を維持したまま、ページやバックグラウンド処理を追加していく。
8. ログオン自動起動が必要になった場合は、必ずユーザーへ確認したうえで実装する（5 章参照）。
