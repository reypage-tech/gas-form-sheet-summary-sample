# 見本画面のデモ PNG を生成する（架空データ）。フォントは /usr/share/fonts から取得。
import csv
import glob
import os
from PIL import Image, ImageDraw, ImageFont

W, H = 1280, 620


def find_font():
    patterns = [
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/**/NotoSansJP-Regular.otf',
        '/usr/share/fonts/**/IPAGothic.ttf',
        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    ]
    for p in patterns:
        hits = glob.glob(p, recursive=True)
        if hits:
            return hits[0]
    raise SystemExit('Japanese-capable font not found under /usr/share/fonts')


def find_font_bold(regular_path):
    bold = regular_path.replace('Regular', 'Bold').replace('Gothic', 'Gothic-Bold')
    if os.path.isfile(bold):
        return bold
    return regular_path


F = find_font()
FB = find_font_bold(F)
f = ImageFont.truetype(F, 15)
fb = ImageFont.truetype(FB, 15)
fc = ImageFont.truetype(FB, 22)
ft = ImageFont.truetype(FB, 18)


def rd(name):
    return list(csv.reader(open('sample_data/' + name, encoding='utf-8')))


responses = rd('responses.csv')
settings = rd('settings.csv')
list_preview = [
    ['タイムスタンプ', '氏名', '所属', '種別', '件名', '内容', '受付番号', '状態', '担当', '期限', '期限切れ'],
    ['2026-10-01 09:15', '架空 太郎', '総務部', '社内備品', 'コピー用紙追加', '500枚', 'R-2026-0001', '未対応', '佐藤 京子', '2026-10-06', '※'],
    ['2026-10-02 11:05', '架空 次郎', '開発部', 'アンケート回答', '満足度調査', 'Q3', 'R-2026-0003', '未対応', '中村 翔', '2026-10-09', '※'],
    ['2026-10-15 08:40', '架空 十八子', '人事部', '社内報投稿', '採用レポート', '新卒', 'R-2026-0018', '対応中', '木村 拓也', '2026-10-21', ''],
    ['2026-10-16 16:50', '架空 十九郎', '総務部', '備品修理', 'エアコン', '冷え不足', 'R-2026-0020', '未対応', '林 真理', '2026-10-28', ''],
]

TABS = ['フォームの回答 1', '設定', '一覧']


def base(active, title):
    im = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 40], fill='#f1f3f4')
    d.text((14, 10), '申込・アンケート集計（見本）', font=ft, fill='#202124')
    d.text((360, 12), 'ファイル  編集  表示  挿入  拡張機能  申込集計', font=f, fill='#3c4043')
    d.rectangle([0, H - 80, W, H - 50], fill='#f1f3f4')
    x = 8
    for t in TABS:
        w = d.textlength(t, font=f) + 20
        if t == active:
            d.rectangle([x, H - 80, x + w, H - 50], fill='white')
            d.text((x + 10, H - 74), t, font=fb, fill='#188038')
        else:
            d.text((x + 10, H - 74), t, font=f, fill='#5f6368')
        x += w + 4
    d.rectangle([0, H - 50, W, H], fill='#202124')
    cap = '見本データ / 再現画面'
    d.text((W // 2 - d.textlength(cap, font=fc) // 2, H - 40), cap, font=fc, fill='#fbbc04')
    d.text((14, 48), title, font=fb, fill='#1a73e8')
    return im, d


def grid(d, rows, y0=78, colw=None):
    if not rows:
        return
    n = len(rows[0])
    cw = colw or [min(120, (W - 20) // n)] * n
    for i, r in enumerate(rows):
        y = y0 + i * 28
        x = 10
        fill = '#e8f0fe' if i == 0 else 'white'
        d.rectangle([10, y, 10 + sum(cw), y + 28], fill=fill, outline='#dadce0')
        for j, c in enumerate(r):
            d.line([x, y, x, y + 28], fill='#dadce0')
            cell_font = fb if i == 0 else f
            txt = str(c)
            max_w = cw[j] - 10
            while txt and d.textlength(txt, font=cell_font) > max_w and len(txt) > 1:
                txt = txt[:-1]
            if len(txt) < len(str(c)) and len(txt) > 0:
                txt = txt[:-1] + '…'
            d.text((x + 5, y + 6), txt, font=cell_font, fill='#202124')
            x += cw[j]


os.makedirs('demo', exist_ok=True)

im, d = base('フォームの回答 1', '1. フォーム送信で受付番号・状態・担当・期限を付与')
grid(d, responses[:8], colw=[132, 72, 68, 108, 112, 148, 98, 62, 88, 92])
im.save('demo/frame_1.png')

im, d = base('設定', '2. 種別ごとの期限日数と担当（土日は数えない）')
grid(d, settings, colw=[200, 100, 160])
im.save('demo/frame_2.png')

im, d = base('一覧', '3. メニュー「一覧を作り直す」: 状態・担当・期限順、期限切れに※')
d.rectangle([18, 46, 318, 72], fill='white', outline='#9aa0a6')
d.text((28, 50), '申込集計 ▼', font=fb, fill='#202124')
d.rectangle([28, 74, 348, 198], fill='white', outline='#9aa0a6')
d.text((40, 82), '申込集計', font=fb, fill='#202124')
d.text((40, 108), '一覧を作り直す', font=f, fill='#202124')
d.text((40, 132), '期限切れの通知を送る', font=f, fill='#202124')
d.text((40, 156), '未付与行を一括補完', font=f, fill='#202124')
d.text((40, 180), '設定シートのひな型', font=f, fill='#202124')
grid(d, list_preview, y0=210, colw=[118, 72, 64, 96, 100, 88, 96, 58, 88, 88, 44])
im.save('demo/frame_3.png')

print('wrote demo/frame_1.png .. frame_3.png')
