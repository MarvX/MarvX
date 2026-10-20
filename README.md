# ساخت پاورپوینت فارسی (راست‌به‌چپ) با پایتون

این پروژه یک ارائه‌ی حرفه‌ای ۱۶:۹ را به‌صورت **کد** می‌سازد: ظاهر و چیدمان در
`build_deck.py` و متن‌ها در `content.py`. هر بار با تغییر متن، فایل `.pptx` از نو
ساخته می‌شود، پس نیازی به جابه‌جایی دستی کادرها در PowerPoint نیست.

خروجی آماده: `output/deck.pptx` (۱۲ اسلاید، نمونه‌ی پروپوزال محصول)

## راه‌اندازی

```bash
python3 -m pip install -r requirements.txt
# اگر محیط گواهی SSL شخصی دارد:
python3 -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
```

## ساخت فایل

```bash
python3 build_deck.py                        # -> output/deck.pptx
python3 build_deck.py -o output/v2.pptx      # نام/مسیر دلخواه
python3 build_deck.py --check                # ساخت + خلاصه‌ی محتوای هر اسلاید
python3 validate_deck.py output/deck.pptx    # کنترل چیدمان (بیرون‌زدگی/هم‌پوشانی)
```

## ویرایش محتوا

فقط `content.py` را تغییر دهید. هر اسلاید یک دیکشنری با کلید `type` است:

| type      | کاربرد            | کلیدها                                          |
|-----------|-------------------|-------------------------------------------------|
| `hero`    | اسلاید عنوان      | `kicker`, `title`, `subtitle`, `meta[]`         |
| `agenda`  | فهرست مطالب       | `kicker`, `title`, `items[]`                    |
| `section` | جداکننده‌ی بخش    | `index`, `title`, `subtitle`                    |
| `bullets` | بولت‌پوینت        | `kicker`, `title`, `items[]`, `note`, `size`    |
| `cards`   | کارت‌های ستونی     | `kicker`, `title`, `items[{title, body}]`       |
| `stats`   | اعداد کلیدی       | `kicker`, `title`, `items[{value, label}]`, `body` |
| `closing` | اسلاید پایانی     | `title`, `subtitle`, `meta[]`                   |

`kicker` همان برچسب کوچک بالای عنوان است و اختیاری‌ست. تعداد کارت‌ها و آمار
دلخواه است (۲ تا ۴ مورد خوب جواب می‌دهد) و به‌صورت خودکار از راست به چپ چیده
می‌شوند. `meta` با جداکننده‌ی `|` نمایش داده می‌شود.

## تغییر تم و فونت

در `build_deck.py`:

```python
THEME = {
    "font": "Tahoma",   # برای فونت وزیرمتن: "Vazirmatn"
    "ink": "0B1F33",    # پس‌زمینه‌ی تیره
    "primary": "0F766E",
    "accent": "E8A33D",
    ...
}
```

فونت باید روی سیستمی که فایل را باز می‌کند نصب باشد؛ `Tahoma` به‌طور پیش‌فرض
همراه Office است. اگر `Vazirmatn` را انتخاب کردید، فونت را نصب یا همراه فایل
ارائه کنید.

## نکته‌های فنی راست‌به‌چپ

- برای هر پاراگراف `rtl="1"` روی `a:pPr` تنظیم می‌شود و تراز راست اعمال می‌شود.
- فونت خط فارسی در OOXML از تگ `a:cs` خوانده می‌شود؛ `python-pptx` فقط
  `a:latin` را ست می‌کند، پس `_set_complex_script` مقدار `a:cs` را هم می‌گذارد.
- بولت‌ها با `a:buFont` + `a:buChar` و تورفتگی `marL/indent` ساخته می‌شوند.
- شماره‌ی اسلاید و سال با ارقام فارسی نمایش داده می‌شوند (`to_fa`).

## ساختار فایل‌ها

```
build_deck.py     موتور ساخت اسلایدها (چیدمان، تم، RTL)
content.py        متن‌ها و ترتیب اسلایدها
validate_deck.py  کنترل چیدمان خروجی
output/deck.pptx  خروجی نهایی
requirements.txt
```

## محدودیت‌ها

- متن به‌صورت خودکار در کادر جا نمی‌شود؛ متن‌های خیلی بلند را کوتاه کنید یا
  اندازه‌ی فونت را در `content.py` (کلید `size` برای `bullets`) کم کنید.
- در این محیط LibreOffice نصب نیست، بنابراین پیش‌نمایش تصویری گرفته نشده است؛
  صحت ساختار فایل و چیدمان با `--check` و `validate_deck.py` بررسی شده است.
