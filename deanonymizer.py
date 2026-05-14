#!/usr/bin/env python3
"""
Anonymus — деанонимайзер: восстанавливает оригинальные значения из токенов.
Работает с текстовыми файлами (.txt, .md).
"""

import json
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

ROOT = Path(__file__).resolve().parent
KEYS_DIR = ROOT / "workspace" / "keys"


def deanonymize(text, mapping):
    replaced = 0
    for token, value in mapping.items():
        if token in text:
            text = text.replace(token, value)
            replaced += 1
    return text, replaced


def main():
    root = Tk()
    root.withdraw()

    text_path = filedialog.askopenfilename(
        title="Выберите файл с заглушками",
        filetypes=[
            ("Текстовые файлы", "*.txt *.md"),
            ("Все файлы", "*.*")
        ]
    )
    if not text_path:
        return

    keys_path = filedialog.askopenfilename(
        title="Выберите файл ключей (.json)",
        initialdir=str(KEYS_DIR),
        filetypes=[("JSON файлы", "*.json")]
    )
    if not keys_path:
        return

    text_path = Path(text_path)
    keys_path = Path(keys_path)

    try:
        text = text_path.read_text(encoding='utf-8')
        mapping = json.loads(keys_path.read_text(encoding='utf-8'))
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")
        return

    result, replaced = deanonymize(text, mapping)
    total = len(mapping)

    out_path = text_path.parent / f"{text_path.stem}_deanon{text_path.suffix}"
    out_path.write_text(result, encoding='utf-8')

    messagebox.showinfo(
        "Деанонимизация завершена",
        f"Восстановлено: {replaced} из {total} токенов.\n\nФайл сохранён:\n{out_path}"
    )


if __name__ == "__main__":
    main()
