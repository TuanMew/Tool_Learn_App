import tkinter as tk
from tkinter import font
import pandas as pd
import pystray, random
import os, sys
from PIL import Image
import signal

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class VocabularySlideshow:
    def __init__(self, root, excel_file):
        self.root = root
        self.root.title("Chinese Vocabulary")
        self.root.overrideredirect(True)

        # Tạo Frame làm viền
        self.border_frame = tk.Frame(self.root, bg="#FFFEEA", borderwidth=12, relief="sunken")
        self.border_frame.pack(fill="both", expand=True)

        self.drag_start_x = 0
        self.drag_start_y = 0
        self.has_dragged = False
        self.border_frame.bind("<Button-1>", self.start_drag)
        self.border_frame.bind("<B1-Motion>", self.on_drag)
        self.border_frame.bind("<ButtonRelease-1>", self.on_click_release)

        # Font chữ hỗ trợ tiếng Trung
        try:
            self.chinese_font = font.Font(family="SimSun", size=50)
        except:
            self.chinese_font = font.Font(family="Arial", size=20)  # Fallback font
        self.pinyin_font = font.Font(family="Segoe UI", size=20)
        self.normal_font = font.Font(family="Arial", size=18)
        self.example_font = font.Font(family="SimSun", size=35)

        # Tạo các nhãn để hiển thị thông tin
        self.hanzi_label = tk.Label(self.border_frame, font=self.chinese_font, bg="#FFFEEA")
        self.pinyin_label = tk.Label(self.border_frame, font=self.pinyin_font, bg="#FFFEEA")
        self.meaning_label = tk.Label(self.border_frame, font=self.normal_font, bg="#FFFEEA")
        self.example_label = tk.Label(self.border_frame, font=self.example_font, bg="#FFFEEA")
        self.pinyin_example_label = tk.Label(self.border_frame, font=self.pinyin_font, bg="#FFFEEA")
        self.translation_label = tk.Label(self.border_frame, font=self.normal_font, bg="#FFFEEA")

        # Đặt vị trí các nhãn và căn giữa
        self.hanzi_label.pack()
        self.pinyin_label.pack()
        self.meaning_label.pack()

        # Đọc từ vựng từ file Excel
        try:
            self.vocab_list = self.load_vocabulary(excel_file)
        except Exception as e:
            self.vocab_list = []
            self.hanzi_label.config(text=f"Lỗi khi đọc file Excel: {str(e)}")
            print(f"lỗi khi đọc file excel: {str(e)}")

        # Tạo biến để duyệt không lặp trong 1 vòng
        if self.vocab_list:
            self.deck = list(range(len(self.vocab_list)))
            random.shuffle(self.deck)
            self.deck_pos = 0
        else:
            self.deck = []
            self.deck_pos = 0

        # Tạo biểu tượng trong system tray
        try:
            icon_path = resource_path("china.ico")  # Đặt file icon.ico trong cùng thư mục
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                # Biểu tượng mặc định nếu không có file icon.ico
                image = Image.new('RGB', (64, 64), color='blue')
        except Exception as e:
            print(f"Không thể tải biểu tượng system tray: {str(e)}")
            image = Image.new('RGB', (64, 64), color='blue')
        menu = (
            pystray.MenuItem("Show", self.show_app, default=True),
            pystray.MenuItem("Hide", self.hide_app),
            pystray.MenuItem("Exit", self.quit_app)
        )
        self.icon = pystray.Icon("Chinese Vocabulary", image, "Chinese Vocabulary", menu)
        self.icon.run_detached()

        # Biến để theo dõi slide hiện tại và trạng thái slideshow
        self.slide_visible = False
        if self.vocab_list:
            self.run_slideshow()
        

    def start_drag(self, event):
        # Lưu tọa độ chuột khi bắt đầu kéo
        self.has_dragged = False
        self.drag_start_x = event.x_root - self.root.winfo_x()
        self.drag_start_y = event.y_root - self.root.winfo_y()

    def on_drag(self, event):
        # Cập nhật vị trí cửa sổ khi kéo
        self.has_dragged = True
        x = event.x_root - self.drag_start_x
        y = event.y_root - self.drag_start_y
        self.root.geometry(f"+{x}+{y}")

    def on_click_release(self, event):
        if not self.has_dragged:
            self.show_full_slide()

    #hiển thị cửa sổ và thoát chương trình
    def show_app(self, *args, **kwargs):
        def _show():
            self.root.deiconify()
            self.root.attributes('-topmost', True)
            self.root.update()
            self.root.attributes('-topmost', False)
        self.root.after(0, _show)
    def hide_app(self, *args, **kwargs):
        self.root.after(0, self.root.withdraw)
    def safe_shutdown(self):
        try:
            if hasattr(self, "icon") and self.icon:
                self.icon.stop()
        except Exception:
            pass
        try:
            self.root.quit()
            self.root.destroy()
        finally:
            sys.exit(0)
    def quit_app(self, *args, **kwargs):
        self.root.after(0, self.safe_shutdown)

    # Đọc file Excel
    def load_vocabulary(self, excel_file):
        df = pd.read_excel(excel_file, engine='openpyxl', dtype=str)
        df = df.fillna("N/A")

        # Đảm bảo các cột đúng tên
        expected_columns = ['Chữ hán', 'Phiên âm', 'Nghĩa', 'Ví dụ', 'Phiên âm ví dụ', 'Dịch']
        for col in expected_columns:
            if col not in df.columns:
                df[col] = "N/A"  # Giá trị mặc định nếu cột không tồn tại
        
        # Lọc bỏ các dòng không có chữ hán
        df = df[df['Chữ hán'].astype(str).str.strip() != '']
        # Loại bỏ khoảng trắng thừa
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        return df.to_dict('records')

    def fade_in_label(self, label, text, target_color, steps=30, delay=50):
        label.config(text=text)
        from_colors = (180, 180, 180)   # Màu xám nhạt ban đầu
        to_rgb = tuple(c // 256 for c in self.root.winfo_rgb(target_color))

        r_step = (to_rgb[0] - from_colors[0]) // steps
        g_step = (to_rgb[1] - from_colors[1]) // steps
        b_step = (to_rgb[2] - from_colors[2]) // steps

        def _step(i=0, r=from_colors[0], g=from_colors[1], b=from_colors[2]):
            color = f'#{r:02x}{g:02x}{b:02x}'
            label.config(fg=color)
            if i < steps:
                self.root.after(delay, _step, i + 1, r + r_step, g + g_step, b + b_step)
            else:
                label.config(fg=target_color)
        _step()

    def show_full_slide(self, event=None):
        if self.slide_visible:
            return
        for label in [self.example_label, self.pinyin_example_label, self.translation_label]:
            if not label.winfo_ismapped():
                label.pack()

        self.root.update_idletasks()
        width = self.border_frame.winfo_reqwidth()
        height = self.border_frame.winfo_reqheight()
        self.root.geometry(f"{width + 10}x{height}")
        self.slide_visible = True

    def draw_next_index(self):
        if not self.deck:
            return None
        idx = self.deck[self.deck_pos]
        self.deck_pos += 1
        if self.deck_pos >= len(self.deck):
            random.shuffle(self.deck)
            self.deck_pos = 0
        return idx

    def update_slide(self):
        if not self.vocab_list:
            return
        
        # Cập nhật nội dung các nhãn
        idx = self.draw_next_index()
        if idx is None:
            return

        #foget label skip hanzi_label, pinyin_label, meaning_label
        self.slide_visible = False
        for label in [self.example_label, self.pinyin_example_label, self.translation_label]:
            label.pack_forget()

        vocab = self.vocab_list[idx]
        self.fade_in_label(self.hanzi_label,f"{vocab['Chữ hán']}","#CA0000")
        self.fade_in_label(self.pinyin_label,f"{vocab['Phiên âm']}","#595656")
        self.fade_in_label(self.meaning_label,f"{vocab['Nghĩa']}","black")
        self.fade_in_label(self.example_label,f"{vocab['Ví dụ']}","#190098")
        self.fade_in_label(self.pinyin_example_label,f"{vocab['Phiên âm ví dụ']}","#595656")
        self.fade_in_label(self.translation_label,f"{vocab['Dịch']}","black")

        self.root.update_idletasks()
        width = self.border_frame.winfo_reqwidth()
        height = self.border_frame.winfo_reqheight()
        self.root.geometry(f"{width + 10}x{height}")

    def continue_app(self):
        self.root.deiconify()
        self.run_slideshow()

    def hide_and_restore_app(self):
        self.root.withdraw()
        self.root.after(5000, self.continue_app)

    def run_slideshow(self):
        if self.vocab_list:
            self.update_slide()
            self.show_app()
            self.root.after(15000, self.hide_and_restore_app)

def main():
    excel_file = resource_path("vocabulary.xlsx")
    root = tk.Tk()
    app = VocabularySlideshow(root, excel_file)

    root.bind("<Escape>", lambda e: app.quit_app())
    def handle_sigint(signum, frame):
        root.after(0, app.safe_shutdown)
    signal.signal(signal.SIGINT, handle_sigint)
    if sys.platform.startswith("win") and hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, handle_sigint)
    root.mainloop()
    
if __name__ == "__main__":
    main()