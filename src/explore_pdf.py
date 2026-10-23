from pathlib import Path
import pdfplumber

PROJECT_DIR = Path(__file__).parent.parent
INPUT_DIR = PROJECT_DIR / "input"
PDF = sorted(INPUT_DIR.glob("*.pdf"))[0]

with pdfplumber.open(PDF) as pdf:
    page = pdf.pages[3]  # СТРАНИЦА 4
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2.0)
    
    sorted_w = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    current = [sorted_w[0]]
    current_top = sorted_w[0]["top"]
    for w in sorted_w[1:]:
        if abs(w["top"] - current_top) <= 3.0:
            current.append(w)
        else:
            lines.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_top = w["top"]
    if current:
        lines.append(sorted(current, key=lambda x: x["x0"]))

    print("=== СТРОКА 1 (ШАПКА) ===")
    for w in lines[1]:
        print(f"{w['text']!r} (x={int(w['x0'])})", end=" | ")
    print("\n\n=== СТРОКА 2 (ДАННЫЕ) ===")
    for w in lines[2]:
        print(f"{w['text']!r} (x={int(w['x0'])})", end=" | ")
    print()