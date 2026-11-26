#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Извлечение рисунков и формул из исходного PDF статьи в work/fig/.

Все вырезки делаются при 400 dpi, обрезаются по фактическим границам
чернил (с полями 12 px) и проверяются на обрезку: если в двух крайних
пиксельных рядах вырезанной области остаются чернила, значит исходный
прямоугольник захватил содержимое не полностью и требуется правка.

Запуск из любого каталога:
    python3 work/crop_final.py

Вход:  1706.03762v7.pdf в корне репозитория
Выход: work/fig/*.png  (5 рисунков + 5 формул)
"""
import os
import sys

import numpy as np
import pymupdf
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PDF = os.path.join(ROOT, '1706.03762v7.pdf')
OUTDIR = os.path.join(HERE, 'fig')

ZOOM = 400 / 72.0          # 400 dpi
PAD_PX = 12                # поля вокруг чернил, пикселей
INK_THRESHOLD = 247        # порог «есть чернила» по яркости

# имя -> (индекс страницы с нуля, прямоугольник в пунктах PDF)
CROPS = {
    'fig1':  (2,  (178, 62, 430, 400)),      # архитектура Трансформера
    'fig2':  (3,  (140, 62, 480, 272)),      # Scaled Dot-Product / Multi-Head
    'fig3':  (12, (104, 94, 512, 304)),      # карты внимания, пример 1
    'fig4':  (13, (106, 168, 512, 606)),     # карты внимания, пример 2
    'fig5':  (14, (106, 182, 508, 596)),     # карты внимания, пример 3
    'eq1':   (3,  (196, 458, 482, 496)),     # Attention(Q, K, V), формула (1)
    'eq_mh': (4,  (178, 136, 430, 182)),     # MultiHead + head_i
    'eq2':   (4,  (214, 556, 482, 586)),     # FFN(x), формула (2)
    'eq_pe': (5,  (220, 330, 396, 372)),     # позиционное кодирование
    'eq3':   (6,  (150, 606, 482, 640)),     # скорость обучения, формула (3)
}


def crop(doc, name, page_index, rect_pt):
    page = doc[page_index]
    rect = pymupdf.Rect(*rect_pt)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(ZOOM, ZOOM), clip=rect,
                          colorspace=pymupdf.csRGB)
    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    gray = np.array(img.convert('L'))

    ys, xs = np.where(gray < INK_THRESHOLD)
    if len(xs) == 0:
        raise SystemExit('%s: в области нет содержимого %s' % (name, rect_pt))

    # чернила вплотную к границе области означают, что вырезка неполная
    edge = np.concatenate([gray[:2, :].ravel(), gray[-2:, :].ravel(),
                           gray[:, :2].ravel(), gray[:, -2:].ravel()])
    truncated = bool((edge < INK_THRESHOLD).any())

    h, w = gray.shape
    box = (max(int(xs.min()) - PAD_PX, 0), max(int(ys.min()) - PAD_PX, 0),
           min(int(xs.max()) + PAD_PX, w - 1), min(int(ys.max()) + PAD_PX, h - 1))
    out = img.crop(box)
    path = os.path.join(OUTDIR, name + '.png')
    out.save(path)

    # в готовой вырезке по периметру не должно быть чернил
    ring = np.concatenate([np.array(out.convert('L'))[0, :],
                           np.array(out.convert('L'))[-1, :],
                           np.array(out.convert('L'))[:, 0],
                           np.array(out.convert('L'))[:, -1]])
    ring_ink = float((ring < INK_THRESHOLD).mean())

    print('%-6s %4dx%-4d ratio=%.2f truncated=%-5s ring_ink=%.4f'
          % (name, out.width, out.height, out.width / out.height,
             truncated, ring_ink))
    if truncated:
        raise SystemExit('%s: область обрезает содержимое, расширьте прямоугольник' % name)


def main():
    if not os.path.exists(PDF):
        raise SystemExit('не найден исходный PDF: %s' % PDF)
    os.makedirs(OUTDIR, exist_ok=True)
    doc = pymupdf.open(PDF)
    for name, (page_index, rect_pt) in CROPS.items():
        crop(doc, name, page_index, rect_pt)
    print('\nготово: %s' % OUTDIR)


if __name__ == '__main__':
    main()
