# LeftHandDevice

Raspberry Pi Pico 2 W (RP2350) をベースに開発した左手デバイス向けの設定用 Windows アプリケーション（WinUI 3 / Windows App SDK）です。

ファームウェアは別場所で管理しています。

- **ファームウェア**: `C:\Users\kazuh\Documents\Arduino\LeftHandDevice`

PC 側のアプリケーションから複雑なマクロパターンを構築し、シリアル通信経由でマイコン本体の EEPROM に直接記憶させるアーキテクチャを採用しています。これにより、専用アプリがバックグラウンドで起動していなくてもどの PC でも同一の動作を行えます。

## ギャラリー

### PC アプリケーション

マクロ設定のメイン画面。ドラッグ＆ドロップによるステップの順番入れ替えに対応。

マウス動作の一括キャプチャ機能。画面上で実際にクリックした座標を自動で取得し、マクロとして記録する。

レジストリを走査し、インストール済みアプリケーション（.exe）を一覧表示。任意のアプリ起動マクロの登録を簡略化。

システム設定画面
<img src="https://github.com/user-attachments/assets/738ec8f7-f14e-4db6-aab9-8d08c0f81a96" alt="soft" width="1000" />

### ハードウェア（デバイス本体）

Raspberry Pi Pico 2 W を組み込んだデバイス本体。

各ボタンを押したときに LED によって状態を判別
<img src="https://github.com/user-attachments/assets/c0cc62f8-bf68-4903-8992-a4d63e140cd0" alt="hard" width="500" />

## 主な機能

- **Pico 連携**: シリアル通信でマクロパターンを送信し、デバイスの EEPROM に同期
- **マウス一括登録**: 低レベルフックでクリック座標をキャプチャしマクロ化
- **音量制御**: ロータリーエンコーダからの `VOL_STEP` を受信し、OS 音量を ±2% 刻みで変更
- **テーマ**: ライト / ダーク / システム連動
- **アップデート確認**: GitHub Release との照合（Inno setup.exe）
- **ナビ UI**: 左サイドバー（ホーム / 情報 / 設定）

## ディレクトリ構成

```text
LeftHandDevice/
├── WindowsApp/
│   └── WinApp/
│       ├── LeftHandDevice.csproj
│       ├── MainWindow.xaml / .cs
│       ├── DeviceService.cs
│       └── Views/
│           ├── HomePage.xaml*
│           ├── InfoPage.xaml*
│           └── SettingsPage.xaml*
├── installer/
├── scripts/
├── ARCHITECTURE.md
├── version_history.txt
└── README.md
```

## ビルド・実行

1. [.NET 8 SDK](https://dotnet.microsoft.com/download) と Windows App SDK 対応環境を用意
2. Visual Studio 2022 等で `WindowsApp/WinApp/LeftHandDevice.csproj` を開く（プラットフォームは x64）
3. ビルドして起動し、COM ポートからデバイスに接続

```powershell
dotnet build WindowsApp\WinApp\LeftHandDevice.csproj -c Debug -p:Platform=x64
```

インストーラ作成:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build-installer.ps1
```

### リリース版を使う場合

1. [Releases](../../releases) から最新の `LeftHandDevice-vX.X.X-win-x64-setup.exe` をダウンロード
2. インストール後に起動し、COM ポートを選んで接続

## 免責事項

本ソフトウェアの使用により生じたいかなる損害についても、開発者は一切の責任を負いません。自己責任でご使用ください。

## ライセンス

MIT License — 詳細は [LICENSE](LICENSE) を参照してください。
