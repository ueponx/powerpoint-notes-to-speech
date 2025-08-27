#!/usr/bin/env bash
# -*- coding: utf-8 -*-
#
# pptx2mp3.sh - PowerPointノート音声化ワンストップスクリプト
#
# PowerPointファイル(.pptx)から発表者ノートを抽出し、
# テキスト整形を経て音声ファイル(MP3)を生成します。
#
# Usage:
#   ./pptx2mp3.sh <input.pptx> <output.mp3> [options]
#
# Examples:
#   ./pptx2mp3.sh presentation.pptx audio.mp3
#   ./pptx2mp3.sh slides.pptx notes.mp3 --speed 1.5 --lang en
#   ./pptx2mp3.sh lecture.pptx lecture.mp3 --keep-temp

set -euo pipefail

# カラー出力の定義
readonly RED='\033[0;31m'
readonly GREEN='\033[0;32m'
readonly YELLOW='\033[1;33m'
readonly BLUE='\033[0;34m'
readonly NC='\033[0m' # No Color

# デフォルト設定
readonly DEFAULT_SPEED="1.2"
readonly DEFAULT_GAIN="3.0"
readonly DEFAULT_SILENCE="250"
readonly DEFAULT_LANG="ja"
readonly DEFAULT_CHUNK="1800"

# スクリプトのディレクトリを取得
readonly SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# =====================================================================
# ヘルプ表示
# =====================================================================
show_help() {
    cat << EOF
PowerPoint Notes to Speech Converter

Usage: $(basename "$0") <input.pptx> <output.mp3> [options]

Arguments:
    input.pptx     Input PowerPoint file
    output.mp3     Output MP3 audio file

Options:
    --speed VALUE      Playback speed (default: ${DEFAULT_SPEED})
    --gain VALUE       Volume gain in dB (default: ${DEFAULT_GAIN})
    --silence VALUE    Silence between chunks in ms (default: ${DEFAULT_SILENCE})
    --lang CODE        Language code (default: ${DEFAULT_LANG})
                       Available: ja, en, zh, ko, es, fr, de, it, pt, ru, ar, hi
    --chunk SIZE       Chunk size in characters (default: ${DEFAULT_CHUNK})
    --keep-temp        Keep intermediate files (markdown and text)
    --verbose          Show detailed output
    --help, -h         Show this help message

Examples:
    # Basic conversion (Japanese)
    $(basename "$0") presentation.pptx output.mp3

    # English presentation with faster speed
    $(basename "$0") slides.pptx notes.mp3 --lang en --speed 1.5

    # Keep intermediate files for review
    $(basename "$0") lecture.pptx audio.mp3 --keep-temp --verbose

EOF
}

# =====================================================================
# エラーハンドリング
# =====================================================================
error_exit() {
    echo -e "${RED}ERROR: $1${NC}" >&2
    exit 1
}

warn_msg() {
    echo -e "${YELLOW}WARNING: $1${NC}" >&2
}

info_msg() {
    echo -e "${BLUE}INFO: $1${NC}"
}

success_msg() {
    echo -e "${GREEN}SUCCESS: $1${NC}"
}

# =====================================================================
# 依存関係チェック
# =====================================================================
check_dependencies() {
    local missing_deps=()
    
    # Python チェック
    if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
        missing_deps+=("python3")
    fi
    
    # uv チェック（オプション）
    if ! command -v uv &> /dev/null; then
        warn_msg "uv is not installed. Using python directly."
        USE_UV=false
    else
        USE_UV=true
    fi
    
    # ffmpeg チェック
    if ! command -v ffmpeg &> /dev/null; then
        missing_deps+=("ffmpeg")
    fi
    
    # 必要なPythonモジュールチェック
    if $USE_UV; then
        if ! uv run python -c "import pptx, gtts, pydub, tqdm" 2>/dev/null; then
            warn_msg "Some Python modules are missing. Installing from pyproject.toml..."
            uv sync || error_exit "Failed to install Python modules"
        fi
    else
        if ! python3 -c "import pptx, gtts, pydub, tqdm" 2>/dev/null; then
            error_exit "Required Python modules are missing. Please install: pip install -r requirements.txt"
        fi
    fi
    
    # 不足している依存関係があれば報告
    if [ ${#missing_deps[@]} -gt 0 ]; then
        error_exit "Missing dependencies: ${missing_deps[*]}\nPlease install them first."
    fi
}

# =====================================================================
# Pythonコマンドのラッパー
# =====================================================================
run_python() {
    if $USE_UV; then
        uv run python "$@"
    else
        python3 "$@" || python "$@"
    fi
}

# =====================================================================
# メイン処理
# =====================================================================
main() {
    # 引数解析用の変数
    local INPUT=""
    local OUTPUT=""
    local SPEED="$DEFAULT_SPEED"
    local GAIN="$DEFAULT_GAIN"
    local SILENCE="$DEFAULT_SILENCE"
    local LANG="$DEFAULT_LANG"
    local CHUNK="$DEFAULT_CHUNK"
    local KEEP_TEMP=false
    local VERBOSE=false
    
    # 引数がない場合はヘルプを表示
    if [ $# -eq 0 ]; then
        show_help
        exit 0
    fi
    
    # 引数を解析
    while [ $# -gt 0 ]; do
        case "$1" in
            --help|-h)
                show_help
                exit 0
                ;;
            --speed)
                SPEED="$2"
                shift 2
                ;;
            --gain)
                GAIN="$2"
                shift 2
                ;;
            --silence)
                SILENCE="$2"
                shift 2
                ;;
            --lang)
                LANG="$2"
                shift 2
                ;;
            --chunk)
                CHUNK="$2"
                shift 2
                ;;
            --keep-temp)
                KEEP_TEMP=true
                shift
                ;;
            --verbose)
                VERBOSE=true
                shift
                ;;
            -*)
                error_exit "Unknown option: $1\nUse --help for usage information."
                ;;
            *)
                if [ -z "$INPUT" ]; then
                    INPUT="$1"
                elif [ -z "$OUTPUT" ]; then
                    OUTPUT="$1"
                else
                    error_exit "Too many arguments\nUse --help for usage information."
                fi
                shift
                ;;
        esac
    done
    
    # 必須引数チェック
    [ -z "$INPUT" ] && error_exit "Input file is required"
    [ -z "$OUTPUT" ] && error_exit "Output file is required"
    
    # 入力ファイルの存在確認
    [ ! -f "$INPUT" ] && error_exit "Input file not found: $INPUT"
    
    # 拡張子チェック
    if [[ ! "$INPUT" =~ \.pptx$ ]]; then
        warn_msg "Input file doesn't have .pptx extension"
    fi
    
    if [[ ! "$OUTPUT" =~ \.mp3$ ]]; then
        warn_msg "Output file doesn't have .mp3 extension"
    fi
    
    # 依存関係チェック
    info_msg "Checking dependencies..."
    check_dependencies
    
    # 中間ファイル名の生成
    local BASE_NAME="${OUTPUT%.mp3}"
    local MD_FILE="${BASE_NAME}_notes.md"
    local TXT_FILE="${BASE_NAME}_clean.txt"
    
    # 進捗表示
    echo "========================================"
    echo "PowerPoint Notes to Speech Converter"
    echo "========================================"
    echo "Input:    $INPUT"
    echo "Output:   $OUTPUT"
    echo "Language: $LANG"
    echo "Speed:    ${SPEED}x"
    echo "Gain:     +${GAIN}dB"
    echo "========================================"
    echo
    
    # ステップ1: ノート抽出
    info_msg "Step 1/3: Extracting notes from PowerPoint..."
    if $VERBOSE; then
        run_python "${SCRIPT_DIR}/notes_export.py" "$INPUT" -o "$MD_FILE" \
            --format md --md-separator="---" --md-pagebreak
    else
        run_python "${SCRIPT_DIR}/notes_export.py" "$INPUT" -o "$MD_FILE" \
            --format md --md-separator="---" --md-pagebreak 2>/dev/null
    fi
    
    # ステップ2: Markdownクリーニング
    info_msg "Step 2/3: Cleaning markdown..."
    if $VERBOSE; then
        run_python "${SCRIPT_DIR}/clean_markdown.py" "$MD_FILE" -o "$TXT_FILE"
    else
        run_python "${SCRIPT_DIR}/clean_markdown.py" "$MD_FILE" -o "$TXT_FILE" 2>/dev/null
    fi
    
    # ステップ3: 音声合成
    info_msg "Step 3/3: Converting to speech..."
    run_python "${SCRIPT_DIR}/tts_gtts_make_mp3.py" "$TXT_FILE" -o "$OUTPUT" \
        --no-clean \
        --speed "$SPEED" \
        --gain "$GAIN" \
        --silence-ms "$SILENCE" \
        --lang "$LANG" \
        --chunk "$CHUNK"
    
    # 中間ファイルの削除（オプション）
    if ! $KEEP_TEMP; then
        rm -f "$MD_FILE" "$TXT_FILE"
        info_msg "Temporary files removed"
    else
        info_msg "Intermediate files kept:"
        echo "  - Markdown: $MD_FILE"
        echo "  - Text:     $TXT_FILE"
    fi
    
    # 完了メッセージ
    echo
    echo "========================================"
    success_msg "Conversion completed successfully!"
    echo "Output saved to: $OUTPUT"
    
    # ファイルサイズの表示
    if [ -f "$OUTPUT" ]; then
        local FILE_SIZE=$(du -h "$OUTPUT" | cut -f1)
        echo "File size: $FILE_SIZE"
    fi
    echo "========================================"
}

# エラー時のクリーンアップ
trap 'echo -e "\n${RED}Process interrupted${NC}" >&2; exit 1' INT TERM

# メイン処理を実行
main "$@"