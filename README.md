# LeftHandDevice (左手デバイス)

Raspberry Pi Pico W を用いた自作の左手デバイス用ファームウェアおよびPC側（Windows WPF）設定アプリケーションです。
デバイス側のキー入力やマウス操作のキャプチャ機能、マクロ設定などをPCから動的に構成することができます。

## 主な機能

- **Pico W 連携**: アプリケーションからシリアル通信経由で各ボタンの動作（単発キー、マウスクリック一括登録など）を送信し、Pico Wに記憶。
- **マクロ/マウス一括登録**: マウスカーソルの動きをキャプチャし、指定したボタンの1アクションとして連続実行させることが可能。
- **ダーク/ライトモード対応**: Windowsのテーマ設定に合わせたUI変更、手動切り替えにも対応。
- **自動アップデート確認**: GitHubのリリース機能と連携し、新しいバージョンがある場合は通知を表示。
- **COMポート自動接続**: 前回接続したCOMポートを記憶し起動時に自動で再接続を試みます。

## ディレクトリ構成

- `LeftHandDevice.ino` : Raspberry Pi Pico W 側のファームウェア（Arduino IDEベース）
- `LeftHandDeviceApp/` : Windows用設定アプリケーション（C# / WPF / .NET）
- `version_history.txt` : 各バージョンの更新履歴

## 環境・ビルド方法

### Raspberry Pi Pico W
- Arduino IDE を使用し、ボードマネージャから Raspberry Pi Pico/RP2040 をインストールしてください。
- `LeftHandDevice.ino` を書き込みます。

### Windows アプリ (LeftHandDeviceApp)
- .NET SDK (net10.0-windows)
- Visual Studio 2022 等で `LeftHandDeviceApp.csproj` を開き、ビルド・実行します。

## 免責事項

本ソフトウェアの使用により生じたいかなる損害（データ消失、システム不具合、セキュリティ上の問題など）についても、開発者は一切の責任を負いません。自己責任でご使用ください。特に重要な作業を行うPCでの使用には十分ご注意ください。

## ライセンス

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
