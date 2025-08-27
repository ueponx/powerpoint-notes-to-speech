#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tts_gtts_make_mp3.py - テキスト音声合成ツール

テキストファイルまたはMarkdownファイルをgTTS（Google Text-to-Speech）で
音声化し、MP3形式で出力するツールです。
長文対応、進捗表示、音声パラメータのカスタマイズが可能です。

Usage:
    # 基本的な使用方法
    python tts_gtts_make_mp3.py input.txt -o output.mp3
    
    # パイプライン処理
    cat input.txt | python tts_gtts_make_mp3.py -o output.mp3 --no-clean
    
    # パラメータのカスタマイズ
    python tts_gtts_make_mp3.py input.txt -o output.mp3 \\
        --chunk 1500 --speed 1.2 --gain 3.0 --silence-ms 250 --lang ja
        
    # Markdownファイルから直接音声化
    python tts_gtts_make_mp3.py notes.md -o notes.mp3
"""
import argparse
import sys
import warnings
import tempfile
from pathlib import Path
from typing import List, Optional, Tuple

# pydubの警告を抑制
warnings.filterwarnings("ignore", category=SyntaxWarning, module=r".*pydub\.utils")

try:
    from gtts import gTTS
    from pydub import AudioSegment, effects
    from tqdm import tqdm
except ImportError as e:
    missing_module = str(e).split("'")[1]
    print(f"ERROR: {missing_module} がインストールされていません。", file=sys.stderr)
    print("  uv add gtts pydub tqdm または pip install gtts pydub tqdm", file=sys.stderr)
    sys.exit(1)


DEFAULT_OUTPUT = "output.mp3"
DEFAULT_CHUNK_SIZE = 1800
DEFAULT_SPEED = 1.2
DEFAULT_GAIN = 3.0
DEFAULT_SILENCE_MS = 250
DEFAULT_LANGUAGE = "ja"

# gTTSがサポートする言語のリスト（主要なもの）
SUPPORTED_LANGUAGES = {
    "ja": "日本語",
    "en": "English",
    "zh": "中文",
    "ko": "한국어",
    "es": "Español",
    "fr": "Français",
    "de": "Deutsch",
    "it": "Italiano",
    "pt": "Português",
    "ru": "Русский",
    "ar": "العربية",
    "hi": "हिन्दी",
}


class TextChunker:
    """テキストを適切なサイズに分割するクラス"""
    
    def __init__(self, chunk_size: int = DEFAULT_CHUNK_SIZE):
        """
        初期化
        
        Args:
            chunk_size: 1チャンクの最大文字数
        """
        self.chunk_size = chunk_size
    
    def chunk_text(self, text: str) -> List[str]:
        """
        長文を適切なサイズに分割する（句点・改行で自然に区切る）
        
        Args:
            text: 分割対象のテキスト
            
        Returns:
            List[str]: 分割されたテキストのリスト
        """
        if not text:
            return []
        
        chunks = []
        i = 0
        n = len(text)
        
        while i < n:
            # チャンクの終了位置を決定
            end = min(i + self.chunk_size, n)
            
            # 文末でない場合、自然な区切り位置を探す
            if end != n:
                slice_text = text[i:end]
                
                # 優先順位: 句点 > 改行 > 読点 > スペース
                break_points = [
                    slice_text.rfind("。"),
                    slice_text.rfind(".\n"),
                    slice_text.rfind("\n"),
                    slice_text.rfind("、"),
                    slice_text.rfind(", "),
                    slice_text.rfind(" "),
                ]
                
                # 最も後ろにある区切り位置を選択
                last_break = max(break_points)
                if last_break > self.chunk_size * 0.3:  # 30%以上の位置なら採用
                    end = i + last_break + 1
            
            # チャンクを抽出
            chunk = text[i:end].strip()
            if chunk:
                chunks.append(chunk)
            
            i = end
        
        return chunks


class TTSProcessor:
    """音声合成処理を行うクラス"""
    
    def __init__(
        self,
        lang: str = DEFAULT_LANGUAGE,
        speed: float = DEFAULT_SPEED,
        gain: float = DEFAULT_GAIN,
        silence_ms: int = DEFAULT_SILENCE_MS
    ):
        """
        初期化
        
        Args:
            lang: 言語コード
            speed: 再生速度（1.0が標準）
            gain: 音量増幅（dB）
            silence_ms: チャンク間の無音（ミリ秒）
        """
        self.lang = lang
        self.speed = speed
        self.gain = gain
        self.silence_ms = silence_ms
    
    def synthesize_to_mp3(
        self,
        text: str,
        out_path: Path,
        chunk_size: int = DEFAULT_CHUNK_SIZE
    ) -> None:
        """
        テキストを音声合成してMP3ファイルを生成
        
        Args:
            text: 音声化するテキスト
            out_path: 出力先のパス
            chunk_size: チャンクサイズ
            
        Raises:
            ValueError: テキストが空の場合
            Exception: 音声合成に失敗した場合
        """
        # テキストを分割
        chunker = TextChunker(chunk_size)
        chunks = chunker.chunk_text(text)
        
        if not chunks:
            raise ValueError("入力テキストが空です。音声化できません。")
        
        # 一時ディレクトリで音声ファイルを生成
        with tempfile.TemporaryDirectory(prefix="gtts_parts_") as tmpdir:
            tmpdir_path = Path(tmpdir)
            part_paths = []
            
            # 各チャンクを音声化（進捗表示付き）
            for i, chunk in enumerate(tqdm(
                chunks, 
                desc=f"音声合成中 (gTTS/{self.lang})", 
                unit="chunk"
            ), 1):
                try:
                    tts = gTTS(chunk, lang=self.lang)
                    part_file = tmpdir_path / f"part_{i:04d}.mp3"
                    tts.save(str(part_file))
                    part_paths.append(part_file)
                except Exception as e:
                    print(f"\nWARNING: チャンク {i} の音声化に失敗: {e}", file=sys.stderr)
                    continue
            
            if not part_paths:
                raise Exception("すべてのチャンクの音声化に失敗しました")
            
            # 音声ファイルを結合
            self._merge_audio_files(part_paths, out_path)
    
    def _merge_audio_files(
        self,
        part_paths: List[Path],
        out_path: Path
    ) -> None:
        """
        複数の音声ファイルを結合して出力
        
        Args:
            part_paths: 結合する音声ファイルのパスリスト
            out_path: 出力先のパス
        """
        # 無音を準備
        silence = AudioSegment.silent(duration=self.silence_ms)
        merged = AudioSegment.silent(duration=0)
        
        # ファイルを結合（進捗表示付き）
        for path in tqdm(
            part_paths, 
            desc="音声結合中 (pydub)", 
            unit="file"
        ):
            try:
                segment = AudioSegment.from_mp3(path)
                merged += segment + silence
            except Exception as e:
                print(f"\nWARNING: {path.name} の読み込みに失敗: {e}", file=sys.stderr)
                continue
        
        # 音量調整
        if self.gain != 0:
            merged = merged + self.gain
        
        # 速度調整
        if abs(self.speed - 1.0) > 1e-3:
            merged = effects.speedup(merged, playback_speed=self.speed)
        
        # 出力ディレクトリを作成
        out_path.parent.mkdir(parents=True, exist_ok=True)
        
        # MP3形式で出力
        merged.export(out_path, format="mp3", bitrate="192k")


def maybe_clean_markdown(text: str, no_clean: bool) -> str:
    """
    必要に応じてMarkdown記法をクリーニングする
    
    Args:
        text: 処理対象のテキスト
        no_clean: クリーニングをスキップするかどうか
        
    Returns:
        str: 処理済みのテキスト
    """
    if no_clean:
        return text
    
    try:
        # clean_markdown.py のインポートを試みる
        from clean_markdown import load_and_clean_markdown
        
        # 一時ファイルを使用してクリーニング
        with tempfile.NamedTemporaryFile(
            "w+", 
            encoding="utf-8", 
            delete=True, 
            suffix=".md"
        ) as tf:
            tf.write(text)
            tf.flush()
            return load_and_clean_markdown(tf.name)
    
    except ImportError:
        # clean_markdown.py が利用できない場合の簡易クリーニング
        import re
        
        # コードブロックを削除
        text = re.sub(r"```.*?```", "", text, flags=re.DOTALL)
        # インラインコードを削除
        text = re.sub(r"`([^`]*)`", r"\1", text)
        # 画像を削除
        text = re.sub(r"!\[([^\]]*)\]\([^)]+\)", r"\1", text)
        # リンクをテキストのみに
        text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
        # 見出し記号を削除
        text = re.sub(r"^#{1,6}\s*", "", text, flags=re.MULTILINE)
        
        return text.strip()


def read_input(src: str) -> str:
    """
    入力ソースからテキストを読み取る
    
    Args:
        src: 入力ソースのパス（"-"の場合は標準入力）
        
    Returns:
        str: 読み取ったテキスト
        
    Raises:
        FileNotFoundError: ファイルが存在しない場合
        RuntimeError: 標準入力が空の場合
    """
    if src == "-":
        data = sys.stdin.read()
        if not data:
            raise RuntimeError(
                "標準入力からデータが読み取れません。"
                "ファイルを指定するか、パイプでテキストを渡してください。"
            )
        return data
    else:
        path = Path(src)
        if not path.exists():
            raise FileNotFoundError(f"入力ファイルが見つかりません: {path}")
        return path.read_text(encoding="utf-8")


def parse_arguments() -> argparse.Namespace:
    """
    コマンドライン引数をパースする
    
    Returns:
        argparse.Namespace: パース結果
    """
    parser = argparse.ArgumentParser(
        description="テキストをgTTSで音声化してMP3出力（進捗表示付き）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    
    parser.add_argument(
        "src",
        nargs="?",
        default="-",
        help="入力ファイルのパス（'-' で標準入力、デフォルト: '-'）"
    )
    parser.add_argument(
        "-o", "--out",
        default=DEFAULT_OUTPUT,
        help=f"出力MP3ファイル名（デフォルト: {DEFAULT_OUTPUT}）"
    )
    parser.add_argument(
        "--chunk",
        type=int,
        default=DEFAULT_CHUNK_SIZE,
        help=f"分割サイズ（文字数、デフォルト: {DEFAULT_CHUNK_SIZE}）"
    )
    parser.add_argument(
        "--speed",
        type=float,
        default=DEFAULT_SPEED,
        help=f"再生速度（1.0が標準、デフォルト: {DEFAULT_SPEED}）"
    )
    parser.add_argument(
        "--gain",
        type=float,
        default=DEFAULT_GAIN,
        help=f"音量増幅[dB]（デフォルト: {DEFAULT_GAIN}）"
    )
    parser.add_argument(
        "--silence-ms",
        type=int,
        default=DEFAULT_SILENCE_MS,
        help=f"チャンク間の無音[ms]（デフォルト: {DEFAULT_SILENCE_MS}）"
    )
    parser.add_argument(
        "--lang",
        default=DEFAULT_LANGUAGE,
        choices=list(SUPPORTED_LANGUAGES.keys()),
        help=f"gTTSの言語（デフォルト: {DEFAULT_LANGUAGE}）"
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Markdownクリーニングをスキップ（入力が既にプレーンテキストの場合）"
    )
    parser.add_argument(
        "--list-languages",
        action="store_true",
        help="サポートされている言語を表示"
    )
    
    return parser.parse_args()


def list_supported_languages() -> None:
    """サポートされている言語のリストを表示"""
    print("サポートされている言語:")
    print("-" * 30)
    for code, name in SUPPORTED_LANGUAGES.items():
        print(f"  {code:5} : {name}")
    print("-" * 30)
    print("※ 他の言語コードも使用可能です（ISO 639-1準拠）")


def main() -> None:
    """メイン処理"""
    args = parse_arguments()
    
    # 言語リストの表示
    if args.list_languages:
        list_supported_languages()
        sys.exit(0)
    
    try:
        # 入力テキストを読み取り
        raw_text = read_input(args.src)
        
        # 必要に応じてクリーニング
        cleaned_text = maybe_clean_markdown(raw_text, no_clean=args.no_clean)
        
        if not cleaned_text:
            print("ERROR: クリーニング後のテキストが空です", file=sys.stderr)
            sys.exit(1)
        
        # 出力パスを準備
        out_path = Path(args.out)
        
        # TTSプロセッサを初期化
        processor = TTSProcessor(
            lang=args.lang,
            speed=args.speed,
            gain=args.gain,
            silence_ms=args.silence_ms
        )
        
        # 音声合成を実行
        print(f"音声合成を開始します...")
        print(f"  言語: {SUPPORTED_LANGUAGES.get(args.lang, args.lang)}")
        print(f"  速度: {args.speed}x")
        print(f"  音量: +{args.gain}dB")
        print(f"  チャンクサイズ: {args.chunk}文字")
        print("-" * 40)
        
        processor.synthesize_to_mp3(
            cleaned_text,
            out_path,
            chunk_size=args.chunk
        )
        
        # 完了メッセージ
        file_size_mb = out_path.stat().st_size / (1024 * 1024)
        print("-" * 40)
        print(f"完了: MP3ファイルを生成しました")
        print(f"  出力: {out_path}")
        print(f"  サイズ: {file_size_mb:.1f} MB")
    
    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except RuntimeError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n処理を中断しました", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: 予期しないエラーが発生: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()