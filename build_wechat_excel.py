from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation
from PIL import Image, ImageDraw


BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "wechat_form_test.xlsx"
PHOTO_PATH = BASE_DIR / "sample_photo.jpg"


def create_sample_photo() -> None:
    img = Image.new("RGB", (960, 640), (224, 237, 245))
    draw = ImageDraw.Draw(img)
    draw.rectangle((120, 120, 840, 520), outline=(42, 102, 152), width=8)
    draw.text((360, 302), "SAMPLE PHOTO", fill=(26, 64, 104))
    img.save(PHOTO_PATH, "JPEG", quality=90)


def build_excel() -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "巡检登记测试"

    headers = [
        "问题地址",
        "问题路段",
        "截止时间(小时数)",
        "问题类别",
        "问题描述",
        "上传照片路径",
        "处置方式",
        "状态",
        "提交时间",
        "样例照片(预览)",
    ]
    ws.append(headers)

    photo_abs = str(PHOTO_PATH)
    rows = [
        [
            "测试地址-幸福路88号门前",
            "幸福路北段",
            24,
            "市政",
            "【test】发现路面坑槽，存在通行隐患，已上报待处置。",
            photo_abs,
            "项目部安排",
            "",
            "",
            "",
        ],
        [
            "测试地址-迎宾大道与东环路口",
            "迎宾大道",
            12,
            "排水",
            "【test】雨后积水，影响通行，需尽快疏通。",
            photo_abs,
            "本班组执行",
            "",
            "",
            "",
        ],
        [
            "测试地址-人民路地铁口",
            "人民路中段",
            48,
            "路灯",
            "【test】路灯不亮，夜间照明不足。",
            photo_abs,
            "指派其他班组",
            "",
            "",
            "",
        ],
    ]
    for row in rows:
        ws.append(row)

    widths = {
        "A": 30,
        "B": 18,
        "C": 16,
        "D": 12,
        "E": 48,
        "F": 56,
        "G": 16,
        "H": 10,
        "I": 20,
        "J": 18,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    category_dv = DataValidation(
        type="list", formula1='"市政,绿化,路灯,排水,保洁"', allow_blank=False
    )
    disposal_dv = DataValidation(
        type="list",
        formula1='"项目部安排,本班组执行,指派其他班组"',
        allow_blank=False,
    )
    ws.add_data_validation(category_dv)
    ws.add_data_validation(disposal_dv)
    category_dv.add("D2:D5000")
    disposal_dv.add("G2:G5000")

    preview = XLImage(str(PHOTO_PATH))
    preview.width = 92
    preview.height = 62
    ws.add_image(preview, "J2")

    wb.save(EXCEL_PATH)


if __name__ == "__main__":
    create_sample_photo()
    build_excel()
    print(f"Created: {EXCEL_PATH}")
    print(f"Created: {PHOTO_PATH}")
