from pathlib import Path
import sys


def latest_zip(folder: Path) -> Path | None:
    zips = list(folder.glob("*.zip"))
    if not zips:
        return None
    return max(zips, key=lambda p: p.stat().st_mtime)


def main(argv: list[str]) -> int:
    folder = Path(argv[1]) if len(argv) > 1 else Path.cwd()
    if not folder.exists():
        print(f"Folder does not exist: {folder}", file=sys.stderr)
        return 2
    last = latest_zip(folder)
    if last:
        print(last.resolve())
        return 0
    else:
        print("No .zip found", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
