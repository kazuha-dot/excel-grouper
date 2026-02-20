import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

EXCEL_EXTS = {".xlsx", ".xlsm", ".xls"}

# デフォルト設定（必要なら config.json を編集）
DEFAULT_CONFIG = {
    "delimiter": "_",                 # 区切り文字（例: "_", "-", " "）
    "use_regex": False,               # Trueなら regex_pattern を使う
    "regex_pattern": r"^(.+?)[ _-]",  # group(1) がフォルダ名になる想定
    "mode": "move",                   # "move" or "copy"
    "skip_if_no_prefix": False        # Trueなら抽出できないファイルはスキップ
}

CONFIG_FILENAME = "config.json"
LOG_FILENAME = "excel_grouper.log"


def get_app_dir() -> Path:
    """exe/py が置かれているフォルダ（処理対象フォルダ）を返す。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent.resolve()
    return Path(__file__).parent.resolve()


def load_config(app_dir: Path) -> dict:
    """config.json があれば読み込んでDEFAULTにマージ。なければDEFAULTを返す。"""
    cfg_path = app_dir / CONFIG_FILENAME
    cfg = DEFAULT_CONFIG.copy()

    if cfg_path.exists():
        try:
            # WindowsでBOM付きになることがあるので utf-8-sig が安全
            user_cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
            if isinstance(user_cfg, dict):
                cfg.update({k: v for k, v in user_cfg.items() if k in cfg})
        except Exception:
            # 壊れてても落とさずデフォルトで動く
            pass

    # 正規化
    cfg["mode"] = str(cfg.get("mode", "move")).lower()
    if cfg["mode"] not in ("move", "copy"):
        cfg["mode"] = "move"

    cfg["delimiter"] = str(cfg.get("delimiter", "_"))
    cfg["use_regex"] = bool(cfg.get("use_regex", False))
    cfg["regex_pattern"] = str(cfg.get("regex_pattern", DEFAULT_CONFIG["regex_pattern"]))
    cfg["skip_if_no_prefix"] = bool(cfg.get("skip_if_no_prefix", False))

    return cfg


def write_default_config_if_missing(app_dir: Path) -> None:
    """初回ユーザーのために、config.json がなければ雛形を作る。"""
    cfg_path = app_dir / CONFIG_FILENAME
    if not cfg_path.exists():
        cfg_path.write_text(
            json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )


def log_line(app_dir: Path, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_path = app_dir / LOG_FILENAME
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"[{ts}] {msg}\n")


def extract_prefix(filename: str, cfg: dict) -> str | None:
    """
    フォルダ名(prefix)を抽出
    - use_regex=True: regex_patternの group(1) を使用
    - use_regex=False: delimiterでsplit（delimiterが無い場合はstem全体）
    """
    base = Path(filename).stem

    if cfg["use_regex"]:
        try:
            m = re.search(cfg["regex_pattern"], base)
            if m and m.group(1).strip():
                return m.group(1).strip()
            return None
        except Exception:
            return None

    delim = cfg["delimiter"]
    if delim and delim in base:
        prefix = base.split(delim, 1)[0].strip()
        return prefix if prefix else None

    return base.strip() if base.strip() else None


def safe_copy_or_move(src: Path, dst_dir: Path, mode: str) -> Path:
    """同名衝突を避けつつ copy/move する。"""
    dst_dir.mkdir(exist_ok=True)
    dst = dst_dir / src.name

    i = 1
    while dst.exists():
        dst = dst_dir / f"{src.stem}({i}){src.suffix}"
        i += 1

    if mode == "copy":
        shutil.copy2(str(src), str(dst))
    else:
        shutil.move(str(src), str(dst))

    return dst


def main() -> int:
    app_dir = get_app_dir()

    # config雛形を生成（コマンド不要で区切り文字オプションを実現）
    write_default_config_if_missing(app_dir)
    cfg = load_config(app_dir)

    log_line(
        app_dir,
        f"START | mode={cfg['mode']} use_regex={cfg['use_regex']} delimiter='{cfg['delimiter']}'"
    )

    processed = 0
    skipped = 0
    errors = 0

    for file in app_dir.iterdir():
        try:
            if not file.is_file():
                continue
            if file.suffix.lower() not in EXCEL_EXTS:
                continue
            if file.name in (CONFIG_FILENAME, LOG_FILENAME):
                continue

            prefix = extract_prefix(file.name, cfg)

            if not prefix:
                if cfg["skip_if_no_prefix"]:
                    skipped += 1
                    log_line(app_dir, f"SKIP(no prefix) | {file.name}")
                    continue
                prefix = file.stem.strip() or "UNGROUPED"

            dst_dir = app_dir / prefix
            safe_copy_or_move(file, dst_dir, cfg["mode"])
            processed += 1
            log_line(app_dir, f"DONE | {file.name} -> {dst_dir.name}/")

        except Exception as e:
            errors += 1
            log_line(app_dir, f"ERROR | {file.name} | {type(e).__name__}: {e}")

    log_line(app_dir, f"END | processed={processed} skipped={skipped} errors={errors}")

    print("\n===== 実行結果 =====")
    print(f"対象フォルダ: {app_dir}")
    print(f"処理モード: {cfg['mode']}")
    if cfg["use_regex"]:
        print(f"抽出方法: 正規表現  pattern={cfg['regex_pattern']}")
    else:
        print(f"抽出方法: 区切り文字  delimiter='{cfg['delimiter']}'")
    print(f"処理: {processed}")
    print(f"スキップ: {skipped}")
    print(f"エラー: {errors}")
    print(f"ログ: {LOG_FILENAME}")
    print("====================")

    input("\nEnterキーで終了します...")
    return 0 if errors == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())