"""用办公软件重新计算后核对实际单元格值，不使用历史公式缓存。"""
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

from config.config_loader import ConfigError


def recalculate_file(path, engine="Excel.Application"):
    """启动独立实例；关闭外链更新和宏，仅保存重算后的生成文件。"""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = book = None
    try:
        app = win32com.client.DispatchEx(engine)
        app.Visible = False
        app.DisplayAlerts = False
        app.AutomationSecurity = 3
        book = app.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=False)
        if book.ReadOnly:
            raise ConfigError("生成文件被占用，无法保存公式重算结果")
        app.CalculateFullRebuild()
        book.Save()
    finally:
        try:
            if book is not None:
                book.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


def check_sheet(sheet, tolerance="0.00", audit=None):
    """定位表头、合计、分录和 Check 记录，不依赖固定数据行数。"""
    def label(value):
        return re.sub(r"\s+", "", str(value)) if value is not None else ""
    def number(row, col):
        cell = sheet.cell(row, col)
        try:
            if cell.value is None or isinstance(cell.value, bool):
                raise ValueError
            value = Decimal(str(cell.value))
            if not value.is_finite():
                raise ValueError
            return value
        except (InvalidOperation, ValueError):
            raise ConfigError(f"核对无法取数：{sheet.title}!{cell.coordinate}={cell.value!r}，请确认公式重算成功")
    try:
        limit = Decimal(str(tolerance))
        if not limit.is_finite() or limit < 0:
            raise ValueError
    except (InvalidOperation, ValueError):
        raise ConfigError("核对容差必须为非负金额")
    def unique(cells, description):
        if len(cells) != 1:
            raise ConfigError(f"无法唯一定位{description}：找到 {len(cells)} 处")
        return cells[0]
    bu_header = unique([c for row in sheet for c in row if label(c.value) == "BU"], "BU 表头")
    bu_col = bu_header.column
    total = unique([sheet.cell(r, bu_col) for r in range(bu_header.row + 1, sheet.max_row + 1)
                    if label(sheet.cell(r, bu_col).value) == "合计"], "收入/税额合计行")
    income, tax = number(total.row, bu_col + 1), number(total.row, bu_col + 2)
    fixed = unique([sheet.cell(r, bu_col) for r in range(bu_header.row + 1, total.row)
                    if label(sheet.cell(r, bu_col).value) == "250108"], "固定 BU 250108 行")
    number(fixed.row, bu_col + 1)
    debit_header = unique([c for row in sheet for c in row if label(c.value) == "借项金额"], "借项金额表头")
    debit_rows = [r for r in range(debit_header.row + 1, sheet.max_row + 1)
                  if label(sheet.cell(r, debit_header.column - 1).value).endswith(".917013")]
    if not debit_rows:
        raise ConfigError("没有找到借项分录行")
    debit = sum((number(r, debit_header.column) for r in debit_rows), Decimal(0))
    check_header = unique([c for row in sheet for c in row
                           if label(c.value) == "分类帐类型1金额"], "Check 金额表头")
    check_rows = [r for r in range(check_header.row + 1, sheet.max_row + 1)
                  if label(sheet.cell(r, check_header.column - 1).value)]
    if not check_rows:
        raise ConfigError("Check 区没有科目明细")
    check_total_row = unique([sheet.cell(r, check_header.column) for r in
                             range(max(check_rows) + 1, sheet.max_row + 1)
                             if sheet.cell(r, check_header.column).value is not None], "Check 合计")
    check_total = number(check_total_row.row, check_total_row.column)
    check_details = sum((number(r, check_header.column) for r in check_rows), Decimal(0))
    def result(name, left, right, locations):
        delta = left - right
        return {"核对项": name, "左值": str(left), "右值": str(right),
                "差额": str(delta), "容差": str(limit), "通过": abs(delta) <= limit,
                "位置": locations}
    if audit is not None:
        original = sheet
        try:
            sheet = audit
            taxable, moh = number(11, 14), number(12, 14)
            difference = number(14, 14)
        finally:
            sheet = original
        return [result("税额合计=借项汇总", tax, debit,
                       {"税额合计": sheet.cell(total.row, bu_col+2).coordinate, "借项行": debit_rows}),
                result("收入合计=审核计税收入+MOH", income, taxable + moh,
                       {"审核计税收入": audit.title + "!N11", "MOH": audit.title + "!N12"}),
                result("审核差额=0", difference, Decimal(0), {"差额": audit.title + "!N14"}),
                result("Check合计=Check明细汇总", check_total, check_details, {"Check明细行": check_rows})]
    return [result("税额合计=借项汇总", tax, debit,
                   {"税额合计": sheet.cell(total.row, bu_col+2).coordinate, "借项行": debit_rows}),
            result("收入合计=Check取反", income, -check_total,
                   {"收入合计": sheet.cell(total.row, bu_col+1).coordinate,
                    "固定BU": fixed.coordinate, "Check合计": check_total_row.coordinate}),
            result("Check合计=Check明细汇总", check_total, check_details,
                   {"Check明细行": check_rows})]
