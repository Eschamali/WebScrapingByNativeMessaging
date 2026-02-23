# VBA Native Messaging PoC

## 概要

本プロジェクトは、ブラウザ拡張機能の通信規格である「ネイティブ メッセージング (Native Messaging)」を、Excel VBAから直接制御するための概念実証（PoC）モデルです。  
本検証における通信インターフェースの「犠牲（ターゲット）」として、[Microsoft Power Automate 拡張機能](https://microsoftedge.microsoft.com/addons/detail/microsoft-power-automate/kagpabjoboikccfdghpdlaaopmgpgfdc)の通信プロトコルを解析・ハイジャックして利用しています。

## ⚠️ 免責事項

完全なる自己責任でのご利用をお願いいたします。

本ツールの使用により発生する可能性のあるいかなる損害についても、開発者は一切の責任を負いません。  
これには以下の事象が含まれますが、これらに限定されません。

* デバイスやシステムへの物理的・論理的な損害
* データの損失や破損
* 拡張機能のライセンス違反等の法的なトラブル

本プロジェクトはあくまで **「ネイティブ メッセージングのアーキテクチャや、Windowsのプロセス間通信（パイプ通信）の仕組みを理解するための学習・研究用」** として作成されたものです。

## 💡 実務でのブラウザ自動操作について（推奨代替案）

本手法はOSの低レイヤー（WinAPI、名前付きパイプ）を直接叩く極めてトリッキーなアプローチであり、対象の拡張機能がアップデートされた瞬間に動作しなくなる致命的なリスク（脆弱なメンテナンス性）を抱えています。

そのため、実務環境のVBAからブラウザ自動操作を実装する場合は、以下の堅牢なアプローチのいずれかを使用することを強く推奨します。

* [SeleniumVBA](https://github.com/GCuser99/SeleniumVBA)
  * 特徴: WebDriver.exe を利用する王道のアプローチ
  * メリット: 最も情報が多く、直感的で安定した操作が可能

* [StarterWebScrapingKit](https://github.com/Eschamali/StarterWebScrapingKit)
  * 特徴: WebDriver.exe 不要。CDP (Chrome DevTools Protocol) を介してブラウザを制御するアプローチ
  * メリット: 外部実行ファイルの配布が不要で、厳しいセキュリティ環境下でも動作しやすい
  * 本家：[Chromium-Automation-with-CDP-for-VBA](https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA)

