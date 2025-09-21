# VOICE_CORRECTOR

音声入力された文字列を文法的に正しい文章に変換するGUIアプリケーションです。

**GitHub Repository**: https://github.com/gpsnmeajp/VOICE_CORRECTOR

## 機能

- 音声認識で入力されたテキストの文法修正
- OpenRouter APIを使用した高精度な文章校正
- 参考用テキストによる文体調整
- クリップボードへの自動コピー
- 変換完了時の音声通知
- 設定の自動保存・復元

## 必要な環境

- Python 3.7以上
- Windows OS（音声通知機能のため）
- OpenRouter APIキー

## インストール

1. このディレクトリで以下のコマンドを実行して必要なライブラリをインストールします：

```powershell
pip install -r requirements.txt
```

2. OpenRouter APIキーを環境変数に設定します：

```powershell
$env:OPENROUTER_API_KEY = "your-api-key-here"
```

または、システム環境変数として`OPENROUTER_API_KEY`を設定してください。

## 使用方法

1. アプリケーションを起動します：

```powershell
python main.py
```

2. 各ボックスに以下を入力します：
   - **入力ボックス**: 音声入力されたテキスト
   - **変換の方針**: 修正の方針（例：「数字は半角、固有名詞は漢字で」）
   - **参考用ボックス**: 文体の参考となるテキスト

3. **変換**ボタンをクリックするか、入力ボックスで**改行を3回連続**入力して文章を修正します

4. 修正された文章は自動的にクリップボードにコピーされます

## 参考用ファイル

`reference`フォルダにテキストファイル（.txt）を保存すると、参考用ボックスセレクターから選択して読み込むことができます。

## 設定

アプリケーションは以下の設定を自動的に保存・復元します：
- 変換の方針
- 参考用テキスト
- 選択された参考用ファイル

設定は`settings.json`ファイルに保存されます。

## APIについて

このアプリケーションはOpenRouter APIの言語モデルを使用してテキストの修正を行います。APIキーの取得方法：

1. [OpenRouter](https://openrouter.ai/)にアクセス
2. アカウントを作成
3. APIキーを取得
4. 環境変数`OPENROUTER_API_KEY`に設定

## トラブルシューティング

### APIキーエラー
- `OPENROUTER_API_KEY`環境変数が正しく設定されているか確認してください

### ライブラリエラー
- `pip install -r requirements.txt`でライブラリが正しくインストールされているか確認してください

## ライセンス

このプロジェクトはMITライセンスの下で提供されています。