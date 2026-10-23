"""Генерация icon.ico из icon.png (стандартная библиотека, без Pillow).

PNG должен быть квадратным RGBA/RGB с альфа-каналом — код поддерживает
Truecolor PNG (ctypes 2/6) и grayscale+alpha (ctypes 4), что покрывает
файлы, сохранённые в любом редакторе. Иконка = BMP-заголовок + пиксели + AND-маска.
"""
import struct
import zlib
from pathlib import Path


def _decode_png(data: bytes):
    """Возвращает (width, height, rgba_rows:list[list[bytes]])."""
    assert data[:8] == b"\x89PNG\r\n\x1a\n", "не PNG"
    pos, idat, w, h, bitdepth, ctype = 8, b"", 0, 0, 8, 6
    palette, trns = None, None
    while pos < len(data):
        ln = struct.unpack(">I", data[pos:pos + 4])[0]
        typ = data[pos + 4:pos + 8]
        chunk = data[pos + 8:pos + 8 + ln]
        if typ == b"IHDR":
            w, h, bitdepth, ctype = struct.unpack(">IIBB", chunk[:10])
        elif typ == b"IDAT":
            idat += chunk
        elif typ == b"PLTE":
            palette = chunk
        elif typ == b"tRNS":
            trns = chunk
        pos += 12 + ln

    raw = zlib.decompress(idat)
    ch = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype]
    bpp = max(1, ch * bitdepth // 8)
    stride = (w * ch * bitdepth + 7) // 8
    prev = bytearray(stride)
    rows = []
    i = 0
    for _ in range(h):
        ftype = raw[i]; i += 1
        line = bytearray(raw[i:i + stride]); i += stride
        if ftype == 1:  # Sub
            for x in range(bpp, stride):
                line[x] = (line[x] + line[x - bpp]) & 0xFF
        elif ftype == 2:  # Up
            for x in range(stride):
                line[x] = (line[x] + prev[x]) & 0xFF
        elif ftype == 3:  # Average
            for x in range(stride):
                a = line[x - bpp] if x >= bpp else 0
                line[x] = (line[x] + (a + prev[x]) // 2) & 0xFF
        elif ftype == 4:  # Paeth
            for x in range(stride):
                a = line[x - bpp] if x >= bpp else 0
                b = prev[x]
                c = prev[x - bpp] if x >= bpp else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[x] = (line[x] + pr) & 0xFF
        prev = line
        px = []
        for x in range(w):
            o = x * bpp
            if ctype == 6:
                px.append(bytes(line[o:o + 4]))
            elif ctype == 2:
                px.append(bytes((line[o], line[o + 1], line[o + 2], 255)))
            elif ctype == 0:
                g = line[o]
                px.append(bytes((g, g, g, 255)))
            elif ctype == 4:
                g, al = line[o], line[o + 1]
                px.append(bytes((g, g, g, al)))
            else:  # indexed
                idx = line[o]
                r, g, b = palette[idx * 3:idx * 3 + 3]
                al = trns[idx] if trns and idx < len(trns) else 255
                px.append(bytes((r, g, b, al)))
        rows.append(px)
    return w, h, rows


def make_ico(png_path: Path, ico_path: Path, size: int = 256) -> Path:
    w, h, rows = _decode_png(png_path.read_bytes())
    s = min(w, h, size)
    x0, y0 = (w - s) // 2, (h - s) // 2
    picked = [rows[y0 + y][x0:x0 + s] for y in range(y0, y0 + s)]

    # XOR: BGRA снизу вверх; AND-маска: прозрачные пиксели = 1
    xor = bytearray()
    for row in reversed(picked):
        for p in row:
            xor += bytes((p[2], p[1], p[0], p[3]))
    and_stride = (s + 31) // 32 * 4
    andm = bytearray(and_stride * s)
    for yi, row in enumerate(reversed(picked)):
        for xi, p in enumerate(row):
            if p[3] < 128:
                andm[yi * and_stride + xi // 8] |= 0x80 >> (xi % 8)

    bmp = struct.pack("<IiiHHIIiiII", 40, s, s * 2, 1, 32, 0,
                      len(xor) + len(andm), 0, 0, 0, 0)
    img = bmp + bytes(xor) + bytes(andm)
    ico = (struct.pack("<HHH", 0, 1, 1)
           + struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0, 1, 32, len(img), 22)
           + img)
    ico_path.write_bytes(ico)
    return ico_path


if __name__ == "__main__":
    here = Path(__file__).resolve().parent.parent / "assets"
    print(make_ico(here / "icon.png", here / "icon.ico"))
