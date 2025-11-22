import json
from pathlib import Path


def main():
    root = Path(__file__).resolve().parent
    live_data_dir = root / "live_data"
    if not live_data_dir.exists():
        raise SystemExit("live_data 目录不存在")

    files = sorted(
        [
            path.name
            for path in live_data_dir.iterdir()
            if path.is_file() and path.suffix.lower() == ".xlsx"
        ],
        reverse=True,
    )

    output = {"files": files}
    output_path = live_data_dir / "index.json"
    output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"写入 {output_path}，共 {len(files)} 个文件")


if __name__ == "__main__":
    main()
