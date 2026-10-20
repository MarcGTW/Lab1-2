"""Единая точка запуска: GUI по умолчанию, --console для лабораторной №1."""
import argparse


def main():
    parser = argparse.ArgumentParser(description="Лабораторные 1–2: генератор билетов")
    parser.add_argument("--console", action="store_true", help="консольный режим лабораторной №1")
    arguments = parser.parse_args()
    if arguments.console:
        from main import main as run
    else:
        from gui import main as run
    return run()


if __name__ == "__main__":
    raise SystemExit(main())
