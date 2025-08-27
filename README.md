# PowerPoint Notes to Speech

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue)](https://www.python.org/)
[![uv](https://img.shields.io/badge/uv-latest-green)](https://github.com/astral-sh/uv)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

PowerPoint の発表者ノートを自動的に**抽出 → 整形 → 音声化（MP3）**するPythonツールキットです。  
研究発表や授業準備の効率化、教材のアクセシビリティ向上にご活用ください。

## 特徴

- **柔軟な出力形式**: Markdown, CSV, JSON に対応
- **自動テキスト整形**: Markdown記法を自動除去
- **高品質な音声合成**: gTTS による多言語対応
- **パイプライン処理**: Unix パイプで効率的な処理
- **進捗表示**: tqdm による処理状況の可視化
- **カスタマイズ可能**: 速度・音量・無音挿入を細かく制御

## クイックスタート

### 前提条件

- Python 3.9 以上
- ffmpeg（音声処理に必要）
- WSL (Windows) / Linux / macOS

### インストール

```bash
# リポジトリをクローン
git clone https://github.com/yourname/powerpoint-notes-to-speech.git
cd powerpoint-notes-to-speech

# uvをインストール（未インストールの場合）
curl -LsSf https://astral.sh/uv/install.sh | sh

# プロジェクトをセットアップ
uv init --python 3.11
uv add python-pptx gtts pydub tqdm

# ffmpegをインストール
# Ubuntu/WSL:
sudo apt update && sudo apt install ffmpeg
# macOS:
brew install ffmpeg
```

### 使い方

#### 最も簡単な方法（シェルスクリプト使用）

```bash
# 実行権限を付与
chmod +x pptx2mp3.sh

# PPTXファイルを音声化
./pptx2mp3.sh presentation.pptx output.mp3
```

#### ステップバイステップ実行

```bash
# 1. PPTXからノートを抽出
uv run python notes_export.py slides.pptx -o notes.md --format md

# 2. Markdownをクリーニング
uv run python clean_markdown.py notes.md -o notes_clean.txt

# 3. テキストを音声化
uv run python tts_gtts_make_mp3.py notes_clean.txt -o audio.mp3 --no-clean
```

#### パイプライン処理（推奨）

```bash
uv run python notes_export.py slides.pptx -o - --format md \
| uv run python clean_markdown.py - -o - \
| uv run python tts_gtts_make_mp3.py -o audio.mp3 --no-clean
```

## 詳細な使い方

### notes_export.py - ノート抽出

PowerPoint ファイルから発表者ノートを抽出します。

```bash
# Markdown形式で出力（区切り線付き）
uv run python notes_export.py presentation.pptx -o notes.md \
  --format md --md-separator="---"

# 印刷用改ページ付き
uv run python notes_export.py presentation.pptx -o notes.md \
  --format md --md-pagebreak

# CSV形式（空行区切り付き）
uv run python notes_export.py presentation.pptx -o notes.csv \
  --format csv --csv-blank-row

# JSON形式
uv run python notes_export.py presentation.pptx -o notes.json \
  --format json
```

### clean_markdown.py - テキスト整形

Markdown記法を除去してプレーンテキストに変換します。

```bash
# ファイル入出力
uv run python clean_markdown.py input.md -o output.txt

# パイプ処理（標準入出力）
cat input.md | uv run python clean_markdown.py - -o - > output.txt
```

### tts_gtts_make_mp3.py - 音声合成

テキストを音声ファイル（MP3）に変換します。

```bash
# 基本的な使用方法
uv run python tts_gtts_make_mp3.py input.txt -o output.mp3

# パラメータのカスタマイズ
uv run python tts_gtts_make_mp3.py input.txt -o output.mp3 \
  --chunk 1500       # チャンクサイズ（文字数）
  --speed 1.2        # 再生速度（1.0が標準）
  --gain 3.0         # 音量増幅（dB）
  --silence-ms 250   # チャンク間の無音（ミリ秒）
  --lang ja          # 言語（ja/en/fr等）
  --no-clean         # クリーニング処理をスキップ
```

### 対応言語

gTTS がサポートする言語に対応しています：

- 日本語: `--lang ja`
- 英語: `--lang en`
- フランス語: `--lang fr`
- スペイン語: `--lang es`
- 中国語: `--lang zh`
- その他多数

## ファイル構成

```
powerpoint-notes-to-speech/
├── README.md                # このファイル
├── pyproject.toml          # uvプロジェクト設定
├── requirements.txt        # 従来の依存関係リスト（互換性用）
├── notes_export.py         # PPTXノート抽出ツール
├── clean_markdown.py       # Markdown整形ツール
├── tts_gtts_make_mp3.py   # 音声合成ツール
├── pptx2mp3.sh            # 一括処理用シェルスクリプト
└── examples/              # サンプルファイル（オプション）
    └── sample.pptx
```

## 活用例

- **発表練習**: 通勤・通学中に音声で発表内容を確認
- **授業準備**: 講義ノートを音声化して内容確認
- **アクセシビリティ**: 視覚に頼らない教材提供
- **復習教材**: 学生向けの音声教材として配布

### カスタマイズ例

#### 英語プレゼンテーションの高速レビュー用

```bash
uv run python notes_export.py presentation.pptx -o - --format md \
| uv run python clean_markdown.py - -o - \
| uv run python tts_gtts_make_mp3.py -o review.mp3 \
  --lang en --speed 1.5 --no-clean
```

#### 長時間講義用（ゆっくり・大音量）

```bash
./pptx2mp3.sh lecture.pptx lecture_audio.mp3
# その後、カスタマイズ
uv run python tts_gtts_make_mp3.py lecture_notes.txt -o lecture_final.mp3 \
  --speed 0.9 --gain 5.0 --silence-ms 500 --no-clean
```

## トラブルシューティング

### ffmpeg が見つからない

```bash
# Ubuntu/WSL
sudo apt update && sudo apt install ffmpeg

# macOS（多分）
brew install ffmpeg

# 確認
ffmpeg -version
```

### uv がインストールできない

```bash
# 代替方法: pip を使用
pip install python-pptx gtts pydub tqdm

# 実行時は python を直接使用
python notes_export.py slides.pptx -o notes.md
```

### 音声が生成されない

- 入力テキストが空でないか確認
- ネットワーク接続を確認（gTTSはオンライン必須）
- 言語設定が正しいか確認

## ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。詳細は [LICENSE](LICENSE) ファイルを参照してください。

## 謝辞

- [python-pptx](https://python-pptx.readthedocs.io/) - PowerPoint ファイル処理
- [gTTS](https://gtts.readthedocs.io/) - Google Text-to-Speech
- [pydub](https://github.com/jiaaro/pydub) - 音声ファイル処理
- [tqdm](https://github.com/tqdm/tqdm) - 進捗表示
- [uv](https://github.com/astral-sh/uv) - 高速パッケージ管理
