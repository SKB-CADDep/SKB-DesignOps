"""Стабильный launcher: читает version.json и запускает нужную версию Baza_Znaniy_2.exe."""

import json
import subprocess
import sys
from pathlib import Path

APP_EXE_NAME = "Baza_Znaniy_2.exe"
VERSIONS_DIR_NAME = "versions"
VERSION_CONFIG_NAME = "version.json"


def get_app_root():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def show_error(message):
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
                None,
                message,
                "База знаний — Launcher",
                0x10,
            )
            return
        except Exception:
            pass

    print(message, file=sys.stderr)


def load_current_version(app_root):
    config_path = app_root / VERSION_CONFIG_NAME
    if not config_path.exists():
        return None

    try:
        with config_path.open(encoding="utf-8-sig") as config_file:
            data = json.load(config_file)
    except (OSError, ValueError) as exc:
        show_error(f"Не удалось прочитать {config_path}.\n\n{exc}")
        return None

    if not isinstance(data, dict):
        show_error(f"Некорректный формат файла {config_path}.")
        return None

    version = str(data.get("current", "")).strip()
    return version or None


def find_app_exe(app_root, version):
    candidates = []

    if version:
        candidates.append(app_root / VERSIONS_DIR_NAME / version / APP_EXE_NAME)

    candidates.append(app_root / APP_EXE_NAME)

    if version:
        versions_dir = app_root / VERSIONS_DIR_NAME
        if versions_dir.is_dir():
            for child in sorted(versions_dir.iterdir(), reverse=True):
                if child.is_dir():
                    candidates.append(child / APP_EXE_NAME)

    seen = set()
    for candidate in candidates:
        normalized = candidate.resolve()
        if normalized in seen:
            continue
        seen.add(normalized)
        if candidate.exists():
            return candidate

    return None


def main():
    app_root = get_app_root()
    version = load_current_version(app_root)
    app_exe = find_app_exe(app_root, version)

    if app_exe is None:
        if version:
            expected = app_root / VERSIONS_DIR_NAME / version / APP_EXE_NAME
            show_error(
                "Не найден файл программы.\n\n"
                f"Ожидался путь:\n{expected}\n\n"
                "Проверьте version.json и папку versions."
            )
        else:
            show_error(
                "Не найден Baza_Znaniy_2.exe.\n\n"
                f"Создайте {VERSION_CONFIG_NAME} и папку versions, "
                "либо положите Baza_Znaniy_2.exe рядом с launcher."
            )
        sys.exit(1)

    result = subprocess.call(
        [str(app_exe), *sys.argv[1:]],
        cwd=str(app_root),
    )
    sys.exit(result)


if __name__ == "__main__":
    main()
