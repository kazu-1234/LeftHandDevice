# LeftHandDevice（Windows アプリ）

左手デバイス（Raspberry Pi Pico 2 W）用の設定・同期アプリケーション（WinUI 3 / Windows App SDK）です。

ファームウェアは別リポジトリで管理しています。

- **ファームウェア**: `C:\Users\kazuh\Documents\Arduino\LeftHandDevice`

## 主な機能

- **Pico 連携**: シリアル通信でマクロパターンを送信し、デバイスの EEPROM に同期
- **マウス一括登録**: 低レベルフックでクリック座標をキャプチャしマクロ化
- **音量制御**: ロータリーエンコーダからの `VOL_STEP` を受信し、OS 音量を ±2% 刻みで変更
- **テーマ**: ライト / ダーク / システム連動
- **アップデート確認**: GitHub Release との照合
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

リリース版の `.exe` を使う場合は GitHub Releases からダウンロードしてください。

## 免責事項

本ソフトウェアの使用により生じたいかなる損害についても、開発者は一切の責任を負いません。自己責任でご使用ください。

## ライセンス

MIT License — 詳細は [LICENSE](LICENSE) を参照してください。
