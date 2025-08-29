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
            pystray.MenuItem("Hide", self.root.withdraw),
            pystray.MenuItem("Exit", self.quit_app)
        )
        self.icon = pystray.Icon("Chinese Vocabulary", image, "Chinese Vocabulary", menu)
        self.icon.run_detached()

        # Biến để theo dõi slide hiện tại và trạng thái slideshow
        self.current_index = 0
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

    def stop_drag(self, event):
        # Kết thúc kéo, không cần xử lý thêm
        pass

    def on_click_release(self, event):
        if not self.has_dragged:
            self.show_full_slide()

    #hiển thị cửa sổ và thoát chương trình
    def show_app(self):
        self.root.deiconify()
        self.root.attributes('-topmost', True)
        self.root.update()
        self.root.attributes('-topmost', False)
    def quit_app(self):
        self.icon.stop()
        self.root.quit()

    def load_vocabulary(self, excel_file):
        # Đọc file Excel
        df = pd.read_excel(excel_file, engine='openpyxl')
        # Chuyển đổi dữ liệu thành danh sách từ điển
        vocab_list = df.to_dict('records')
        # Đảm bảo các cột đúng tên
        expected_columns = ['Chữ hán', 'Phiên âm', 'Nghĩa', 'Ví dụ', 'Phiên âm ví dụ', 'Dịch']
        for vocab in vocab_list:
            for col in expected_columns:
                if col not in vocab:
                    vocab[col] = "N/A"  # Thêm giá trị mặc định nếu thiếu cột
        return vocab_list

    def fade_in_label(self, label, text, target_color, steps=30, delay=50):
        label.config(text=text)
        # Tạo các bước màu dần từ màu xám sang target
        from_colors = [(180, 180, 180)]  # Màu xám nhạt ban đầu
        to_rgb = self.root.winfo_rgb(target_color)  # (65535, 0, 0)
        to_rgb = tuple(c // 256 for c in to_rgb)

        r_step = (to_rgb[0] - from_colors[0][0]) // steps
        g_step = (to_rgb[1] - from_colors[0][1]) // steps
        b_step = (to_rgb[2] - from_colors[0][2]) // steps

        def _step(i=0, r=from_colors[0][0], g=from_colors[0][1], b=from_colors[0][2]):
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
        for label in [self.meaning_label, self.example_label, self.pinyin_example_label, self.translation_label]:
            if not label.winfo_ismapped():
                label.pack()
            self.root.update_idletasks()
            width = self.border_frame.winfo_reqwidth()
            height = self.border_frame.winfo_reqheight()
            self.root.geometry(f"{width + 10}x{height}")
            self.slide_visible = True

    def update_slide(self):
        if not self.vocab_list:
            return
        # Cập nhật nội dung các nhãn
        previous_index = getattr(self, 'current_index', None)
        new_index = previous_index
        while new_index == previous_index:
            new_index = random.randint(0, len(self.vocab_list) - 1)

        #foget label skip hanzi_label, pinyin_label, meaning_label
        self.slide_visible = False
        for label in [self.example_label, self.pinyin_example_label, self.translation_label]:
            label.pack_forget()

        self.current_index = new_index
        vocab = self.vocab_list[self.current_index]
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
    
    def on_escape(self, event=None):
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

def main():
    excel_file = resource_path("vocabulary.xlsx")
    root = tk.Tk()
    app = VocabularySlideshow(root, excel_file)

    # Nhấn ESC cũng thoát
    root.bind("<Escape>", app.on_escape)

    # Bắt Ctrl+C từ terminal
    def handle_sigint(signum, frame):
        # gọi thoát trên GUI thread để an toàn
        root.after(0, app.on_escape)

    signal.signal(signal.SIGINT, handle_sigint)
    # (Tùy chọn) Windows: Ctrl+Break
    if sys.platform.startswith("win") and hasattr(signal, "SIGBREAK"):
        signal.signal(signal.SIGBREAK, handle_sigint)
        
    root.mainloop()
if __name__ == "__main__":
    main()