import os
import shutil
import winreg

def install_fonts():
    local_app_data = os.environ.get("LOCALAPPDATA")
    fonts_dir = os.path.join(local_app_data, "Microsoft", "Windows", "Fonts")
    os.makedirs(fonts_dir, exist_ok=True)
    
    temp_dir = os.environ.get("TEMP")
    src_dirs = [
        os.path.join(temp_dir, "pashto_fonts", "all_fonts"),
        os.path.join(temp_dir, "pashto_fonts_2", "fonts", "ps"),
    ]
    
    reg_key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_key_path, 0, winreg.KEY_SET_VALUE)
    
    count = 0
    installed_fonts = []
    
    for s_dir in src_dirs:
        if not os.path.exists(s_dir):
            continue
        for root, dirs, files in os.walk(s_dir):
            for file in files:
                if file.lower().endswith((".ttf", ".otf")):
                    src_file = os.path.join(root, file)
                    dest_file = os.path.join(fonts_dir, file)
                    if not os.path.exists(dest_file):
                        shutil.copy2(src_file, dest_file)
                    
                    font_name = os.path.splitext(file)[0]
                    ext_tag = "OpenType" if file.lower().endswith(".otf") else "TrueType"
                    reg_val_name = f"{font_name} ({ext_tag})"
                    
                    try:
                        winreg.SetValueEx(reg_key, reg_val_name, 0, winreg.REG_SZ, dest_file)
                        count += 1
                        installed_fonts.append(file)
                    except Exception as e:
                        pass
                        
    winreg.CloseKey(reg_key)
    print(f"Successfully installed and registered {count} Pashto fonts in Windows registry and fonts directory.")
    print("Key Pashto fonts installed include:")
    for f in [
        "Bahij Nazanin-Regular.ttf", "Bahij Nazanin-Bold.ttf",
        "Bahij Zar-Regular.ttf", "Bahij Zar-Bold.ttf",
        "Bahij Titr-Bold.ttf", "AwamiNastaliq-Regular.ttf",
        "Scheherazade-Regular.ttf", "Khoshal Baba Medium.otf",
        "Hamza Baba.ttf", "Pashto Nastaliq.ttf", "Pashto Destaar.ttf",
        "AGT Pakhtoonkhwa Bold.otf", "AGT Peshawer.otf"
    ]:
        if f in installed_fonts or os.path.exists(os.path.join(fonts_dir, f)):
            print(f" - {f}")

if __name__ == "__main__":
    install_fonts()
