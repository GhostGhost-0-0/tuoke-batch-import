"""
Win32 文件选择对话框处理模块
供尺寸图、PDF文件等复用，查找对话框、填写文件名、点击打开
"""

import time
import ctypes
import uiautomation as auto
from .ui_utils import human_delay, mouse_click

user32 = ctypes.windll.user32

# 点击「选择…认证文件」后弹出的 Win32 对话框（#32770）的窗口标题；
# 泰国 / 菲律宾各两条，与下面 thPdfNames / phPdfNames 的前两项一致。
PDF_TRIGGER_DIALOG_NAMES_TH = (
    "选择Thai FDA cosmetic registration number on the product label认证文件",
    "选择Product label with the TIS certification mark认证文件",
)
PDF_TRIGGER_DIALOG_NAMES_PH = (
    "选择Food and Drug Administration (FDA) Certificate of Product Registration (CPR)认证文件",
    "选择Philippine Standard (PS) or Import Commodity Clearance (ICC) Certificate认证文件",
)


def resolve_site_pdf_region(site: str) -> str:
    """返回 'th' | 'ph' | ''，与 _find_file_dialog 判定一致。"""
    site_upper = (site or "").upper()
    if "泰" in (site or "") or "泰国" in (site or "") or "THAI" in site_upper or "TH" == site_upper:
        return "th"
    if "菲" in (site or "") or "菲律宾" in (site or "") or "PHILIP" in site_upper or "PH" == site_upper:
        return "ph"
    return ""


def category_pdf_dialog_title_hint(site: str, column_key: str):
    """
    类目多 PDF 列上传时，本次弹窗最可能对应的 #32770 标题（用于在 _find_file_dialog 中优先匹配）。
    :param column_key: FDApdf文件 | TISpdf文件 | PS/ICCpdf文件
    """
    region = resolve_site_pdf_region(site)
    if region == "th":
        names = PDF_TRIGGER_DIALOG_NAMES_TH
    elif region == "ph":
        names = PDF_TRIGGER_DIALOG_NAMES_PH
    else:
        return None
    if column_key == "FDApdf文件":
        return names[0]
    if column_key == "TISpdf文件":
        return names[1] if region == "th" else None
    if column_key == "PS/ICCpdf文件":
        return names[1] if region == "ph" else None
    return None


def _find_file_dialog(timeout_attempts=20, site: str = None, dialog_title_hint: str = None):
    """
    查找 Win32 文件选择对话框（ClassName=#32770，含文件名输入框 1148）
    :param site: 站点名称，泰国站用 thPdfNames，菲律宾站用 phPdfNames，尺寸图/其他用 sizeNames
    :param dialog_title_hint: 若提供，在本轮尝试中优先按该 Name 查找 #32770（如多列 PDF 对应不同长标题）
    :return: 对话框控件或 None
    """
    # 方法1: 直接通过 ClassName 和常见 Name 查找（按站点选择标题列表）
    sizeNames = ["请选择文件夹", "打开", "Open", "选择文件"]
    thPdfNames = list(PDF_TRIGGER_DIALOG_NAMES_TH) + ["打开", "Open", "选择文件"]
    phPdfNames = list(PDF_TRIGGER_DIALOG_NAMES_PH) + ["打开", "Open", "选择文件"]
    pdf_region = resolve_site_pdf_region(site)
    if pdf_region == "th":
        names_to_try = thPdfNames
    elif pdf_region == "ph":
        names_to_try = phPdfNames
    else:
        names_to_try = sizeNames
    if dialog_title_hint and str(dialog_title_hint).strip():
        h = str(dialog_title_hint).strip()
        names_to_try = [h] + [n for n in names_to_try if n != h]
    for attempt in range(timeout_attempts):
        for name in names_to_try:
            try:
                d = auto.WindowControl(ClassName="#32770", Name=name)
                if d.Exists(maxSearchSeconds=0.3):
                    try:
                        if d.EditControl(AutomationId="1148").Exists(maxSearchSeconds=0.3) or d.ComboBoxControl(AutomationId="1148").Exists(maxSearchSeconds=0.3):
                            return d
                    except:
                        pass
            except:
                pass
        time.sleep(0.5)
    
    # 方法2: 遍历所有窗口，找任一带 1148 输入框的 #32770（尺寸图/PDF 对话框结构相同）
    for attempt in range(timeout_attempts):
        try:
            root = auto.GetRootControl()
            for w in root.GetChildren():
                try:
                    if (w.ControlType == auto.ControlType.WindowControl and
                            (w.ClassName or "") == "#32770"):
                        try:
                            if w.EditControl(AutomationId="1148").Exists(maxSearchSeconds=0.3) or w.ComboBoxControl(AutomationId="1148").Exists(maxSearchSeconds=0.3):
                                return w
                        except:
                            pass
                except:
                    continue
        except:
            pass
        time.sleep(0.5)
    
    # 方法3: 查找模态对话框
    for attempt in range(10):
        try:
            root = auto.GetRootControl()
            for w in root.GetChildren():
                try:
                    if w.ControlType == auto.ControlType.WindowControl:
                        try:
                            wp = w.GetWindowPattern()
                            if wp and wp.IsModal:
                                w_name = (w.Name or "").strip()
                                w_class = w.ClassName or ""
                                if w_class == "#32770" and w_name and w_name not in ("批量导入", "类目修改"):
                                    try:
                                        if w.EditControl(AutomationId="1148").Exists(maxSearchSeconds=0.3) or w.ComboBoxControl(AutomationId="1148").Exists(maxSearchSeconds=0.3):
                                            return w
                                    except:
                                        pass
                        except:
                            pass
                except:
                    continue
        except:
            pass
        time.sleep(0.5)
    
    return None


def _fill_filename(dialog, file_name: str):
    """在文件对话框中填写文件名"""
    file_name_edit = None
    try:
        file_name_edit = dialog.EditControl(AutomationId="1148")
        if not file_name_edit.Exists(maxSearchSeconds=1):
            file_name_edit = dialog.ComboBoxControl(AutomationId="1148")
        if not file_name_edit or not file_name_edit.Exists(maxSearchSeconds=0.5):
            file_name_edit = None
    except:
        file_name_edit = None
    
    if not file_name_edit:
        # 遍历查找 EditControl
        all_edits = []
        def find_edits(ctrl, depth=0):
            if depth > 5:
                return
            try:
                if ctrl.ControlType == auto.ControlType.EditControl:
                    all_edits.append(ctrl)
            except:
                pass
            try:
                child = ctrl.GetFirstChildControl()
                while child:
                    find_edits(child, depth + 1)
                    child = child.GetNextSiblingControl()
            except:
                pass
        try:
            find_edits(dialog)
            if all_edits:
                file_name_edit = all_edits[0]
        except:
            pass
    
    if file_name_edit and file_name_edit.Exists(maxSearchSeconds=1):
        try:
            vp = file_name_edit.GetValuePattern()
            if vp:
                vp.SetValue("")
                time.sleep(0.2)
                vp.SetValue(str(file_name))
                time.sleep(0.3)
                return True
        except:
            pass
        try:
            file_name_edit.SetFocus()
            time.sleep(0.3)
            user32.keybd_event(0x11, 0, 0, 0)
            user32.keybd_event(0x41, 0, 0, 0)
            time.sleep(0.1)
            user32.keybd_event(0x41, 0, 2, 0)
            user32.keybd_event(0x11, 0, 2, 0)
            time.sleep(0.2)
            user32.keybd_event(0x2E, 0, 0, 0)
            time.sleep(0.05)
            user32.keybd_event(0x2E, 0, 2, 0)
            time.sleep(0.2)
            auto.SendKeys(str(file_name))
            time.sleep(0.5)
            return True
        except:
            pass
        try:
            dialog.SetFocus()
            time.sleep(0.3)
            user32.keybd_event(0x11, 0, 0, 0)
            user32.keybd_event(0x41, 0, 0, 0)
            time.sleep(0.1)
            user32.keybd_event(0x41, 0, 2, 0)
            user32.keybd_event(0x11, 0, 2, 0)
            time.sleep(0.2)
            user32.keybd_event(0x2E, 0, 0, 0)
            time.sleep(0.05)
            user32.keybd_event(0x2E, 0, 2, 0)
            time.sleep(0.2)
            auto.SendKeys(str(file_name))
            time.sleep(0.5)
            return True
        except:
            pass
    else:
        try:
            dialog.SetFocus()
            time.sleep(0.3)
            user32.keybd_event(0x11, 0, 0, 0)
            user32.keybd_event(0x41, 0, 0, 0)
            time.sleep(0.1)
            user32.keybd_event(0x41, 0, 2, 0)
            user32.keybd_event(0x11, 0, 2, 0)
            time.sleep(0.2)
            user32.keybd_event(0x2E, 0, 0, 0)
            time.sleep(0.05)
            user32.keybd_event(0x2E, 0, 2, 0)
            time.sleep(0.2)
            auto.SendKeys(str(file_name))
            time.sleep(0.5)
            return True
        except:
            pass
    return False


def _click_open_button(dialog):
    """点击文件对话框的"打开"按钮"""
    open_button = None
    for nm in ["打开(&O)", "打开(O)", "打开", "Open"]:
        try:
            btn = dialog.ButtonControl(Name=nm)
            if btn.Exists(maxSearchSeconds=1):
                open_button = btn
                break
        except:
            pass
    
    if not open_button:
        try:
            open_button = dialog.ButtonControl(AutomationId="1")
            if not open_button.Exists(maxSearchSeconds=0.5):
                open_button = None
        except:
            open_button = None
    
    if not open_button:
        def find_open_button(ctrl, depth=0):
            if depth > 5:
                return None
            try:
                if ctrl.ControlType == auto.ControlType.ButtonControl:
                    btn_name = (ctrl.Name or "").strip()
                    if "打开" in btn_name or "Open" in btn_name:
                        return ctrl
            except:
                pass
            try:
                child = ctrl.GetFirstChildControl()
                while child:
                    result = find_open_button(child, depth + 1)
                    if result:
                        return result
                    child = child.GetNextSiblingControl()
            except:
                pass
            return None
        try:
            open_button = find_open_button(dialog)
        except:
            pass
    
    if open_button and open_button.Exists(maxSearchSeconds=1):
        try:
            open_button.Click()
            human_delay(1.0, 1.5)
            return True
        except:
            pass
        try:
            invoke_pattern = open_button.GetInvokePattern()
            if invoke_pattern:
                invoke_pattern.Invoke()
                human_delay(1.0, 1.5)
                return True
        except:
            pass
        try:
            rect = open_button.BoundingRectangle
            if rect.width() > 0 and rect.height() > 0:
                cx = rect.left + rect.width() // 2
                cy = rect.top + rect.height() // 2
                if mouse_click(cx, cy):
                    human_delay(1.0, 1.5)
                    return True
        except:
            pass
    
    try:
        dialog.SetFocus()
        time.sleep(0.3)
        auto.SendKeys("{ENTER}")
        human_delay(1.0, 1.5)
        return True
    except:
        pass
    return False


def handle_file_select_dialog(
    file_name: str,
    wait_after_click: float = 1.5,
    site: str = None,
    dialog_title_hint: str = None,
) -> bool:
    """
    查找 Win32 文件选择对话框、填写文件名、点击打开
    供尺寸图、PDF 文件等复用
    :param file_name: 要选择的文件名
    :param wait_after_click: 点击打开按钮前的等待时间（秒），用于等待对话框出现
    :param site: 站点名称，泰国站查 thPdfNames，菲律宾站查 phPdfNames，尺寸图/其他查 sizeNames
    :param dialog_title_hint: 优先匹配的 #32770 窗口标题（与 PDF_TRIGGER_DIALOG_NAMES_* 中长标题一致）
    :return: 是否成功
    """
    if wait_after_click > 0:
        time.sleep(wait_after_click)
    
    print("  🔍 等待文件选择对话框出现...")
    dialog = _find_file_dialog(site=site, dialog_title_hint=dialog_title_hint)
    if not dialog:
        print("  ⚠️ 未找到文件选择对话框")
        return False
    
    print(f"  ✓ 找到文件选择对话框: {(dialog.Name or '')[:50]}")
    print(f"  📝 准备输入文件名: {file_name}")
    
    if _fill_filename(dialog, file_name):
        print(f"  ✓ 已输入文件名: {file_name}")
    else:
        print(f"  ⚠️ 输入文件名可能失败，继续尝试点击打开")
    
    if _click_open_button(dialog):
        print("  ✓ 已点击'打开'按钮")
        human_delay(0.5, 0.8)
        return True
    
    print("  ⚠️ 点击打开按钮失败")
    return False
