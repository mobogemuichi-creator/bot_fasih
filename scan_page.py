import uiautomator2 as u2
import os
import subprocess

LDPLAYER_ADB = r"C:\LDPlayer\LDPlayer9\adb.exe"

def hubungkan_adb():
    for port in ["5555", "5555"]:
        if os.path.exists(LDPLAYER_ADB):
            subprocess.run([LDPLAYER_ADB, "connect", f"127.0.0.1:{port}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            d = u2.connect(f"127.0.0.1:{port}")
            d.info
            return d
        except Exception:
            continue
    return None

import re
from PIL import Image

def ambil_warna_koordinat(img, bounds_str):
    """
    Mengambil warna RGB & HEX piksel tengah elemen dari string bounds '[x1,y1][x2,y2]'
    """
    try:
        match = re.search(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds_str)
        if match:
            x1, y1, x2, y2 = map(int, match.groups())
            cx = (x1 + x2) // 2
            cy = (y1 + y2) // 2
            
            if 0 <= cx < img.size[0] and 0 <= cy < img.size[1]:
                pixel = img.getpixel((cx, cy))
                r, g, b = pixel[:3]
                hex_code = f"#{r:02x}{g:02x}{b:02x}"
                return f"RGB=({r},{g},{b}) | HEX={hex_code} di ({cx},{cy})"
    except Exception:
        pass
    return None

def main():
    d = hubungkan_adb()
    if not d:
        print("[ERROR] Gagal terhubung ke emulator.")
        return
        
    # Dump XML
    xml_data = d.dump_hierarchy()
    with open("dump.xml", "w", encoding="utf-8") as f:
        f.write(xml_data)
    print("[SUKSES] XML disimpan ke dump.xml")
    
    # Ambil Screenshot untuk analisis warna
    try:
        screenshot_img = d.screenshot()
    except Exception:
        screenshot_img = None
    
    # Cek khusus status Checkbox 'Reject' dan checkbox lainnya
    print("\n--- Status Checkbox di Layar ---")
    checkboxes = [
        ("Open", "id.go.bpsfasih:id/open_cb_bottomSheetFilterAssignment"),
        ("Pernah dibuka", "id.go.bpsfasih:id/pernahDibuka_cb_bottomSheetFilterAssignment"),
        ("Submit", "id.go.bpsfasih:id/submit_cb_bottomSheetFilterAssignment"),
        ("Approve", "id.go.bpsfasih:id/approve_cb_bottomSheetFilterAssignment"),
        ("Reject", "id.go.bpsfasih:id/reject_cb_bottomSheetFilterAssignment"),
    ]
    for label, res_id in checkboxes:
        cb_el = d(resourceId=res_id)
        if not cb_el.exists():
            cb_el = d(text=label)
        if cb_el.exists():
            is_checked = cb_el.info.get("checked", False)
            status_txt = "DICENTANG [V]" if is_checked else "BELUM DICENTANG [X]"
            
            warna_str = ""
            if screenshot_img and hasattr(cb_el, 'info'):
                bounds = cb_el.info.get('bounds', {})
                if bounds:
                    cx = (bounds.get('left', 0) + bounds.get('right', 0)) // 2
                    cy = (bounds.get('top', 0) + bounds.get('bottom', 0)) // 2
                    pixel = screenshot_img.getpixel((cx, cy))
                    r, g, b = pixel[:3]
                    warna_str = f" | Warna: RGB=({r},{g},{b}) HEX=#{r:02x}{g:02x}{b:02x}"
                    
            print(f"  - Checkbox '{label}' ({res_id}): {status_txt} (checked={is_checked}){warna_str}")
        else:
            print(f"  - Checkbox '{label}': Tidak ditemukan di layar.")

    # Cetak rangkuman teks & content-desc elemen beserta atribut checked & warna jika ada
    print("\n--- Rangkuman Elemen Teks / Content-Desc di Layar ---")
    elements = d.xpath("//*").all()
    found = 0
    for el in elements:
        txt = el.text.strip() if el.text else ""
        desc = el.attrib.get('content-desc', '').strip() if el.attrib.get('content-desc') else ""
        
        display_label = txt if txt else desc
        if desc and txt and desc != txt:
            display_label = f"Text: '{txt}' | Content-Desc: '{desc}'"
        elif txt:
            display_label = f"Text: '{txt}'"
        elif desc:
            display_label = f"Content-Desc: '{desc}'"

        if display_label:
            found += 1
            resource_id = el.attrib.get('resource-id', '')
            cls = el.attrib.get('class', '')
            checked_attr = el.attrib.get('checked', '')
            checkable_attr = el.attrib.get('checkable', '')
            bounds_str = el.attrib.get('bounds', '')

            extra_info = ""
            if checkable_attr == 'true' or checked_attr:
                extra_info += f" | Checked: {checked_attr}"

            if screenshot_img and bounds_str:
                warna = ambil_warna_koordinat(screenshot_img, bounds_str)
                if warna:
                    extra_info += f" | Warna: {warna}"

            print(f"[{found}] {display_label} | ID: '{resource_id}' | Class: '{cls}'{extra_info}")

    # Pengecekan khusus elemen Plus (+) / Minus (-)
    print("\n--- Deteksi Elemen Plus (+) / Minus (-) ---")
    plus_minus_els = d.xpath("//*[contains(@text, '+') or contains(@text, '-') or contains(@content-desc, '+') or contains(@content-desc, '-')]").all()
    if plus_minus_els:
        for idx, pm_el in enumerate(plus_minus_els, start=1):
            pm_txt = pm_el.text.strip() if pm_el.text else ""
            pm_desc = pm_el.attrib.get('content-desc', '').strip() if pm_el.attrib.get('content-desc') else ""
            pm_bounds = pm_el.attrib.get('bounds', '')
            print(f"  [{idx}] Text: '{pm_txt}' | Content-Desc: '{pm_desc}' | Bounds: {pm_bounds}")
    else:
        print("  - Tidak ada elemen yang mengandung simbol '+' atau '-' di atribut text/content-desc.")

if __name__ == "__main__":
    main()
