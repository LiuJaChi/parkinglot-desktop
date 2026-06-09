# -*- coding: utf-8 -*-
"""
啟動器 - 南方莊園車輛通行記錄導入管理系統
"""

import tkinter as tk
import sys
import os

# 添加當前目錄到路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from Qphoto import VehicleImporterGUI


def main():
    """主程序入口"""
    try:
        root = tk.Tk()
        app = VehicleImporterGUI(root)
        root.mainloop()
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print("錯誤信息:")
        print(error_msg)
        
        # 顯示錯誤對話框
        try:
            import tkinter.messagebox as messagebox
            messagebox.showerror("應用啟動失敗", f"無法啟動應用:\n{str(e)}")
        except:
            pass


if __name__ == "__main__":
    main()
