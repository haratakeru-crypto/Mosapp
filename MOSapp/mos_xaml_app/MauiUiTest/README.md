# MOS Excel Mogi App - MAUI UIテスト版

このプロジェクトは、MOS Excel模擬アプリのUIテスト機能をMAUI（.NET Multi-platform App UI）で実装したものです。

## 概要

WPF版のUIテスト機能（`UiTestAppBarWindow`）をMAUIに移植したアプリケーションです。デザインと機能を可能な限り維持しながら、モダンなMAUIフレームワークで再構築しています。

## 機能

- **タイマー機能**: 50分のカウントダウンタイマー
- **プロジェクト管理**: プロジェクト1-10の切り替え
- **タスクナビゲーション**: タスク1-7の表示と切り替え
- **採点機能**: プロジェクトの採点実行（準備中）
- **一時停止機能**: タイマーの一時停止/再開
- **リセット機能**: プロジェクトのリセット

## 要件

- .NET 9.0 SDK
- Windows 10/11（Windows専用MAUIアプリ）

## ビルドと実行

```bash
cd MauiUiTest
dotnet build
dotnet run
```

## プロジェクト構造

```
MauiUiTest/
├── Pages/
│   └── UiTestAppBarPage.xaml      # UIテストページ（メインUI）
├── MainPage.xaml                   # ホームページ
├── AppShell.xaml                  # アプリケーションシェル
└── MOSExcelMogiApp.Maui.csproj    # プロジェクトファイル
```

## 主な変更点（WPF版との比較）

### UI要素の違い

| WPF | MAUI |
|-----|------|
| `Window` | `ContentPage` |
| `TextBlock` | `Label` |
| `DispatcherTimer` | `System.Timers.Timer` |
| `SystemParameters` | `DeviceDisplay.MainDisplayInfo` |
| `MessageBox` | `DisplayAlert` |

### 実装の違い

1. **タイマー**: `System.Timers.Timer`を使用し、`MainThread.BeginInvokeOnMainThread`でUI更新
2. **ナビゲーション**: MAUIの`Navigation.PushAsync`を使用
3. **ダイアログ**: `DisplayAlert`を使用（`MessageBox`の代替）

## 今後の拡張予定

- [ ] 採点機能の完全実装
- [ ] プロジェクトデータの読み込み（config.json）
- [ ] タスク状態の永続化
- [ ] レビューページ機能
- [ ] タイマー無効化チェックボックスの統合

## 注意事項

- 現在はWindows専用のMAUIアプリとして実装されています
- Excel Interop機能は含まれていません（UIテスト専用）
- 一部の機能（採点、レビューページ）は準備中です

## ライセンス

元のWPFアプリケーションと同じライセンスに従います。

