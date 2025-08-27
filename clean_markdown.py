#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
clean_markdown.py - Markdownクリーニングツール

Markdown記法を除去してプレーンテキストに変換するツールです。
音声合成の前処理として、読み上げに適さない記号や構造を削除します。

Usage:
    # ファイル入出力
    python clean_markdown.py input.md -o output.txt
    
    # 標準入出力（パイプ処理）
    cat input.md | python clean_markdown.py - -o -
    
    # 標準入力から読み取り、ファイルに出力
    cat input.md | python clean_markdown.py - -o output.txt
"""
import argparse
import re
import sys
import tempfile
from pathlib import Path
from typing import Optional, Union, List, Pattern


class MarkdownCleaner:
    """Markdown記法をクリーニングするクラス"""
    
    def __init__(self):
        """クリーナーを初期化し、正規表現パターンをコンパイル"""
        # 各種パターンを事前コンパイル（パフォーマンス向上）
        self.patterns: List[tuple[Pattern, str]] = [
            # コードブロック ``` ... ``` を削除
            (re.compile(r"```.*?```", re.DOTALL), ""),
            
            # インデントされたコードブロック（4スペース）を削除
            (re.compile(r"^    .*$", re.MULTILINE), ""),
            
            # インラインコード `code` を中身だけに
            (re.compile(r"`([^`]*)`"), r"\1"),
            
            # 画像 ![alt](url) をaltテキストのみに
            (re.compile(r"!\[([^\]]*)\]\([^)]+\)"), r"\1"),
            
            # リンク [text](url) をテキストのみに
            (re.compile(r"\[([^\]]+)\]\([^)]+\)"), r"\1"),
            
            # 参照スタイルのリンク [text][id] をテキストのみに
            (re.compile(r"\[([^\]]+)\]\[[^\]]+\]"), r"\1"),
            
            # 脚注定義 [^1]: ... を削除
            (re.compile(r"^\[\^[^\]]+\]:.*$", re.MULTILINE), ""),
            
            # 脚注参照 [^1] を削除
            (re.compile(r"\[\^[^\]]+\]"), ""),
            
            # 見出し # を削除
            (re.compile(r"^#{1,6}\s*", re.MULTILINE), ""),
            
            # 水平線 --- や *** を削除
            (re.compile(r"^[\-\*]{3,}$", re.MULTILINE), ""),
            
            # 引用 > を削除
            (re.compile(r"^>\s*", re.MULTILINE), ""),
            
            # リスト記号 *, -, + を削除
            (re.compile(r"^[\*\-\+]\s+", re.MULTILINE), ""),
            
            # 番号付きリスト 1. を削除
            (re.compile(r"^\d+\.\s+", re.MULTILINE), ""),
            
            # テーブルの区切り |---|---|を削除
            (re.compile(r"^\|[\s\-\|:]+\|$", re.MULTILINE), ""),
            
            # テーブルのパイプ | を削除
            (re.compile(r"\|"), " "),
            
            # 太字 **text** または __text__ を通常テキストに
            (re.compile(r"\*\*([^\*]+)\*\*"), r"\1"),
            (re.compile(r"__([^_]+)__"), r"\1"),
            
            # 斜体 *text* または _text_ を通常テキストに
            (re.compile(r"\*([^\*]+)\*"), r"\1"),
            (re.compile(r"_([^_]+)_"), r"\1"),
            
            # 打ち消し線 ~~text~~ を通常テキストに
            (re.compile(r"~~([^~]+)~~"), r"\1"),
            
            # HTMLタグを削除
            (re.compile(r"<[^>]+>"), ""),
            
            # 複数の空行を1つに
            (re.compile(r"\n{3,}"), "\n\n"),
        ]
    
    def clean(self, text: str) -> str:
        """
        Markdown記法をクリーニングする
        
        Args:
            text: クリーニング対象のテキスト
            
        Returns:
            str: クリーニング済みのプレーンテキスト
        """
        if not text:
            return ""
        
        # 各パターンを順番に適用
        for pattern, replacement in self.patterns:
            text = pattern.sub(replacement, text)
        
        # 行頭・行末の空白を削除
        lines = text.split('\n')
        lines = [line.strip() for line in lines]
        text = '\n'.join(lines)
        
        # 最終的な空白の調整
        text = text.strip()
        
        return text


def load_and_clean_markdown(md_path: str) -> str:
    """
    Markdownファイルを読み込んでクリーニングする（後方互換性のため）
    
    Args:
        md_path: Markdownファイルのパス
        
    Returns:
        str: クリーニング済みのプレーンテキスト
    """
    cleaner = MarkdownCleaner()
    text = Path(md_path).read_text(encoding="utf-8")
    return cleaner.clean(text)


def read_input(src: str) -> str:
    """
    入力ソースからテキストを読み取る
    
    Args:
        src: 入力ソースのパス（"-"の場合は標準入力）
        
    Returns:
        str: 読み取ったテキスト
        
    Raises:
        FileNotFoundError: ファイルが存在しない場合
        RuntimeError: 標準入力からデータが読み取れない場合
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


def write_output(content: str, dst: str) -> None:
    """
    出力先にテキストを書き込む
    
    Args:
        content: 出力するテキスト
        dst: 出力先のパス（"-"の場合は標準出力）
    """
    if dst == "-" or not dst:
        sys.stdout.write(content)
    else:
        Path(dst).write_text(content, encoding="utf-8")


def parse_arguments() -> argparse.Namespace:
    """
    コマンドライン引数をパースする
    
    Returns:
        argparse.Namespace: パース結果
    """
    parser = argparse.ArgumentParser(
        description="Markdown記法をクリーニングしてプレーンテキストに変換",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    
    parser.add_argument(
        "src",
        help="入力Markdownファイルのパス（'-' で標準入力）"
    )
    parser.add_argument(
        "-o", "--out",
        default="-",
        help="出力先のパス（'-' で標準出力、デフォルト: '-'）"
    )
    parser.add_argument(
        "--preserve-links",
        action="store_true",
        help="リンクURLを保持する（デフォルトは削除）"
    )
    parser.add_argument(
        "--preserve-emphasis",
        action="store_true",
        help="強調記法（太字・斜体）を保持する（デフォルトは削除）"
    )
    
    return parser.parse_args()


def main() -> None:
    """メイン処理"""
    args = parse_arguments()
    
    try:
        # 入力テキストを読み取り
        if args.src == "-":
            # 標準入力の場合は一時ファイルに保存
            raw_text = sys.stdin.read()
            if not raw_text:
                print("ERROR: 標準入力が空です", file=sys.stderr)
                sys.exit(1)
            
            # MarkdownCleanerを使用
            cleaner = MarkdownCleaner()
            
            # オプションに応じてパターンを調整
            if args.preserve_links:
                # リンク関連のパターンを除外
                cleaner.patterns = [
                    p for p in cleaner.patterns 
                    if not any(x in str(p[0].pattern) for x in [r"\[", "http"])
                ]
            
            if args.preserve_emphasis:
                # 強調記法関連のパターンを除外
                cleaner.patterns = [
                    p for p in cleaner.patterns 
                    if not any(x in str(p[0].pattern) for x in [r"\*", r"_", r"~~"])
                ]
            
            cleaned = cleaner.clean(raw_text)
        else:
            # ファイル入力の場合
            text = read_input(args.src)
            cleaner = MarkdownCleaner()
            
            # オプションに応じてパターンを調整
            if args.preserve_links:
                cleaner.patterns = [
                    p for p in cleaner.patterns 
                    if not any(x in str(p[0].pattern) for x in [r"\[", "http"])
                ]
            
            if args.preserve_emphasis:
                cleaner.patterns = [
                    p for p in cleaner.patterns 
                    if not any(x in str(p[0].pattern) for x in [r"\*", r"_", r"~~"])
                ]
            
            cleaned = cleaner.clean(text)
        
        # 出力
        write_output(cleaned, args.out)
        
        # 完了メッセージ（ファイル出力の場合）
        if args.out != "-" and args.out:
            print(f"完了: クリーニングを完了しました: {args.out}", file=sys.stderr)
    
    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except RuntimeError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: 予期しないエラーが発生: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()