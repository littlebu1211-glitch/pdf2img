import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import sys
import subprocess

# 設定編碼支援
if sys.platform.startswith('win'):
    import locale
    locale.setlocale(locale.LC_ALL, 'Chinese_Taiwan.utf8')

# 動態檢查和載入套件
def check_and_install_packages():
    required_packages = {
        'fitz': 'PyMuPDF',
        'PIL': 'Pillow'
    }
    
    missing_packages = []
    
    for module, package in required_packages.items():
        try:
            if module == 'fitz':
                import fitz
            elif module == 'PIL':
                from PIL import Image
        except ImportError:
            missing_packages.append(package)
    
    return missing_packages

# 檢查套件
missing = check_and_install_packages()
if missing:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "缺少套件", 
        f"請先安裝必要套件:\n\n" +
        "\n".join([f"pip install {pkg}" for pkg in missing]) +
        f"\n\n缺少套件: {', '.join(missing)}"
    )
    sys.exit(1)

# 成功載入套件
import fitz  # PyMuPDF
from PIL import Image
import io  # 新增io模組用於JPG轉換

class PDFtoPNGConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF轉圖片轉換器 (BU@Claude)")  # 更新標題
        self.root.geometry("600x800")
        self.root.resizable(False, False)
        
        # 設定字型，優先使用系統預設字型
        try:
            # Windows繁體中文系統常見字型
            font_options = ['Microsoft JhengHei UI', 'Microsoft JhengHei', 'PMingLiU', '新細明體']
            self.system_font = None
            
            for font in font_options:
                try:
                    test_font = (font, 12)
                    # 測試字型是否可用
                    test_label = tk.Label(self.root, text="測試", font=test_font)
                    test_label.destroy()
                    self.system_font = font
                    break
                except:
                    continue
            
            if not self.system_font:
                self.system_font = 'TkDefaultFont'
                
        except Exception as e:
            self.system_font = 'TkDefaultFont'
            print(f"字型設定警告: {e}")
        
        self.selected_file = None
        self.dpi_var = tk.StringVar(value="300")
        self.format_var = tk.StringVar(value="PNG")  # 新增格式選項變數
        
        try:
            self.setup_ui()
        except Exception as e:
            messagebox.showerror("初始化錯誤", f"程式初始化失敗:\n{str(e)}")
            sys.exit(1)
        
    def setup_ui(self):
        try:
            # 主框架
            main_frame = ttk.Frame(self.root, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 標題
            title_label = ttk.Label(main_frame, text="PDF轉圖片轉換器", 
                                   font=(self.system_font, 18, 'bold'))
            title_label.pack(pady=(0, 30))
            
            # 檔案選擇區域
            file_frame = ttk.LabelFrame(main_frame, text="選擇PDF檔案", padding="15")
            file_frame.pack(fill=tk.X, pady=(0, 20))
            
            self.file_label = ttk.Label(file_frame, text="尚未選擇檔案", 
                                       foreground="gray", wraplength=500,
                                       font=(self.system_font, 10))
            self.file_label.pack(pady=(0, 10))
            
            select_button = ttk.Button(file_frame, text="選擇PDF檔案", 
                                      command=self.select_file, width=20)
            select_button.pack()
            
            # DPI選擇區域
            dpi_frame = ttk.LabelFrame(main_frame, text="選擇解析度", padding="15")
            dpi_frame.pack(fill=tk.X, pady=(0, 20))
            
            dpi_options = [
                ("150 DPI", "150"),
                ("300 DPI", "300"),
                ("600 DPI", "600")
            ]
            
            for text, value in dpi_options:
                radio = ttk.Radiobutton(dpi_frame, text=text, variable=self.dpi_var, 
                                       value=value)
                radio.pack(anchor=tk.W, pady=2)
            
            # 檔案格式選擇區域
            format_frame = ttk.LabelFrame(main_frame, text="選擇輸出格式", padding="15")
            format_frame.pack(fill=tk.X, pady=(0, 20))
            
            format_options = [
                ("PNG格式 (無壓縮，檔案較大)", "PNG"),
                ("JPG格式 (高品質壓縮，檔案較小)", "JPG")
            ]
            
            for text, value in format_options:
                radio = ttk.Radiobutton(format_frame, text=text, variable=self.format_var, 
                                       value=value)
                radio.pack(anchor=tk.W, pady=2)
            
            # 轉換按鈕
            convert_button = ttk.Button(main_frame, text="開始轉換", 
                                       command=self.start_conversion)
            convert_button.pack(pady=20)
            
            # 進度區域
            progress_frame = ttk.LabelFrame(main_frame, text="轉換進度", padding="15")
            progress_frame.pack(fill=tk.X, pady=(0, 20))
            
            self.progress_var = tk.StringVar(value="請選擇PDF檔案開始轉換")
            self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var,
                                           wraplength=500, font=(self.system_font, 10))
            self.progress_label.pack(pady=(0, 10))
            
            self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
            self.progress_bar.pack(fill=tk.X)
            
            # 狀態區域
            status_frame = ttk.Frame(main_frame)
            status_frame.pack(fill=tk.X)
            
            self.status_var = tk.StringVar(value="準備就緒")
            status_label = ttk.Label(status_frame, textvariable=self.status_var,
                                    font=(self.system_font, 9))
            status_label.pack(side=tk.LEFT)
            
        except Exception as e:
            messagebox.showerror("介面建立錯誤", f"無法建立使用者介面:\n{str(e)}")
            raise
        
    def select_file(self):
        """選擇PDF檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇PDF檔案",
            filetypes=[("PDF檔案", "*.pdf")],
            initialdir=os.path.expanduser("~")
        )
        
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"已選擇: {filename}", foreground="black")
            self.status_var.set("請選擇PDF檔案開始轉換")
    
    def start_conversion(self):
        """開始轉換流程"""
        if not self.selected_file:
            messagebox.showerror("錯誤", "請先選擇PDF檔案")
            return
        
        # 選擇輸出目錄
        output_dir = filedialog.askdirectory(
            title="選擇輸出目錄",
            initialdir=os.path.dirname(self.selected_file)
        )
        
        if not output_dir:
            return
        
        # 在新線程中執行轉換，避免UI卡死
        thread = threading.Thread(target=self.convert_pdf_to_images, 
                                args=(self.selected_file, output_dir))
        thread.daemon = True
        thread.start()
    
    def convert_pdf_to_images(self, pdf_path, output_dir):
        """轉換PDF為圖片檔"""
        try:
            # 更新狀態
            self.status_var.set("正在處理PDF檔案...")
            self.progress_var.set("正在讀取PDF檔案...")
            
            # 開啟PDF檔案
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            
            # 建立輸出資料夾
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_folder = os.path.join(output_dir, pdf_name)
            
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            # 設定DPI
            dpi = int(self.dpi_var.get())
            zoom = dpi / 72.0  # PDF預設72 DPI
            mat = fitz.Matrix(zoom, zoom)
            
            # 取得選擇的格式
            selected_format = self.format_var.get()
            file_extension = selected_format.lower()
            
            self.progress_bar.config(maximum=total_pages)
            
            # 轉換每一頁
            for page_num in range(total_pages):
                self.progress_var.set(f"正在轉換第 {page_num + 1} 頁，共 {total_pages} 頁...")
                self.progress_bar.config(value=page_num)
                self.root.update_idletasks()
                
                # 獲取頁面
                page = doc[page_num]
                pix = page.get_pixmap(matrix=mat)
                
                # 生成檔名 (流水號，從001開始)
                output_filename = f"{page_num + 1:03d}.{file_extension}"
                output_path = os.path.join(output_folder, output_filename)
                
                # 根據格式儲存檔案
                if selected_format == "PNG":
                    # 儲存為PNG
                    pix.save(output_path)
                elif selected_format == "JPG":
                    # 轉換為PIL Image並儲存為高品質JPG
                    img_data = pix.tobytes("ppm")
                    img = Image.open(io.BytesIO(img_data))
                    # 如果是RGBA模式，轉換為RGB（JPG不支援透明度）
                    if img.mode == 'RGBA':
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        background.paste(img, mask=img.split()[-1])  # 使用alpha通道作為遮罩
                        img = background
                    img.save(output_path, "JPEG", quality=99, optimize=True)
            
            # 完成
            self.progress_bar.config(value=total_pages)
            self.progress_var.set(f"轉換完成！共轉換 {total_pages} 頁")
            self.status_var.set("轉換完成")
            
            doc.close()
            
            # 顯示完成對話框
            self.root.after(0, lambda: self.show_completion_dialog(output_folder, total_pages))
            
        except Exception as e:
            error_msg = f"轉換過程中發生錯誤: {str(e)}"
            self.status_var.set("轉換失敗")
            self.progress_var.set(error_msg)
            self.root.after(0, lambda: messagebox.showerror("錯誤", error_msg))
    
    def show_completion_dialog(self, output_folder, page_count):
        """顯示完成對話框"""
        try:
            result = messagebox.askyesno(
                "轉換完成", 
                f"PDF轉換完成！\n\n"
                f"轉換頁數: {page_count} 頁\n"
                f"輸出位置: {output_folder}\n"
                f"解析度: {self.dpi_var.get()} DPI\n"
                f"檔案格式: {self.format_var.get()}\n\n"
                f"是否要開啟輸出資料夾？"
            )
            
            if result:
                # 開啟輸出資料夾，支援不同作業系統
                try:
                    if sys.platform.startswith('win'):
                        os.startfile(output_folder)
                    elif sys.platform.startswith('darwin'):  # macOS
                        subprocess.call(['open', output_folder])
                    else:  # Linux
                        subprocess.call(['xdg-open', output_folder])
                except Exception as e:
                    messagebox.showinfo("提示", f"請手動開啟資料夾:\n{output_folder}")
        except Exception as e:
            messagebox.showinfo("轉換完成", f"轉換完成！輸出位置:\n{output_folder}")

def main():
    try:
        # 設定控制台編碼（Windows系統）
        if sys.platform.startswith('win'):
            try:
                # 設定控制台編碼為UTF-8
                import codecs
                sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
                sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
            except Exception:
                pass  # 如果失敗就忽略，使用預設編碼
        
        # 建立主視窗
        root = tk.Tk()
        root.title("PDF轉圖片轉換器")  # 更新標題
        
        # 確保在主執行緒中處理GUI
        try:
            app = PDFtoPNGConverter(root)
        except Exception as e:
            messagebox.showerror("應用程式錯誤", f"無法啟動應用程式:\n{str(e)}")
            return
        
        # 設定視窗關閉事件
        def on_closing():
            try:
                root.quit()
                root.destroy()
            except Exception:
                pass
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # 啟動應用程式
        root.mainloop()
        
    except Exception as e:
        # 如果連建立視窗都失敗，就顯示命令列錯誤
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("系統錯誤", f"程式無法啟動:\n{str(e)}")
        except Exception:
            print(f"程式啟動失敗: {str(e)}")
            input("按Enter鍵退出...")
        sys.exit(1)

if __name__ == "__main__":
    main()