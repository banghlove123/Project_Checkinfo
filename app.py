import tkinter as tk
import platform
import psutil
import os
import subprocess
import wmi
from tkinter import simpledialog, messagebox

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("System Monitor")
        self.geometry("800x600")

        self.label_font = ("Helvetica", 12, "bold")
        self.text_font = ("Courier New", 10)

        self.create_log_frame()
        self.create_keyboard_frame()
        self.create_info_frame()

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.bind("<Key>", self.on_key_press)
        self.load_masver_log()

    def create_log_frame(self):
        log_frame = tk.LabelFrame(self, text="Log", padx=10, pady=5)
        log_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.log_text = tk.Text(log_frame, height=10, width=80, font=self.text_font, state=tk.DISABLED)
        self.log_text.pack(side=tk.LEFT, fill=tk.Y)
        self.log_scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=self.log_scrollbar.set)

    def create_keyboard_frame(self):
        keyboard_frame = tk.LabelFrame(self, text="Keyboard Input", padx=10, pady=5)
        keyboard_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.keyboard_text = tk.Text(keyboard_frame, height=5, width=80, font=self.text_font, state=tk.DISABLED)
        self.keyboard_text.pack(side=tk.LEFT, fill=tk.Y)
        self.keyboard_scrollbar = tk.Scrollbar(keyboard_frame, command=self.keyboard_text.yview)
        self.keyboard_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.keyboard_text.config(yscrollcommand=self.keyboard_scrollbar.set)

    def create_info_frame(self):
        info_frame = tk.LabelFrame(self, text="System Information", padx=10, pady=5)
        info_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.create_info_labels(info_frame)

    def create_info_labels(self, info_frame):
        info_data = [
            ("Activation Status:", self.get_activation_status, 0),
            ("Disk C:", lambda: self.get_disk_space("C:\\"), 1),
            ("Disk D:", lambda: self.get_disk_space("D:\\"), 2),
            ("Windows Version:", lambda: platform.platform(), 3),
            ("RAM Size:", self.get_ram_size, 4),
            ("Serial Number:", self.get_serial_number, 5),
        ]

        for text, func, row in info_data:
            label = tk.Label(info_frame, text=text, font=self.label_font)
            label.grid(row=row, column=0, sticky="w")
            value_label = tk.Label(info_frame, text=func(), font=self.text_font)
            value_label.grid(row=row, column=1, sticky="w")
            if text == "Serial Number:":
                self.serial_number_label = value_label
                self.add_serial_button = tk.Button(info_frame, text="Add Serial", font=self.label_font, command=self.get_user_input)
                self.add_serial_button.grid(row=row, column=2, sticky="w")

    def load_masver_log(self):
        file_path = r"C:\Windows\MASVER.txt"
        try:
            with open(file_path, "r") as f:
                lines = f.readlines()[:9]
                self.log_text.config(state=tk.NORMAL)
                self.log_text.delete("1.0", tk.END)
                self.log_text.insert(tk.END, "".join(lines))
                self.log_text.config(state=tk.DISABLED)
        except FileNotFoundError:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete("1.0", tk.END)
            self.log_text.insert(tk.END, "File not found.")
            self.log_text.config(state=tk.DISABLED)
        except Exception as e:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete("1.0", tk.END)
            self.log_text.insert(tk.END, f"Error: {e}")
            self.log_text.config(state=tk.DISABLED)

    def on_key_press(self, event):
        key = event.char if event.char != "" else event.keysym
        self.keyboard_text.config(state=tk.NORMAL)
        self.keyboard_text.insert(tk.END, f"{key} ")
        self.keyboard_text.see(tk.END)
        self.keyboard_text.config(state=tk.DISABLED)

    def get_activation_status(self):
        try:
            result = subprocess.run(["cscript", "//Nologo", "C:\\Windows\\System32\\slmgr.vbs", "/xpr"],
                                    capture_output=True, text=True, encoding="utf-8")
            output = result.stdout.strip()
            if "permanently activated" in output.lower():
                return "Windows Activated"
            elif "will expire" in output.lower():
                return "Windows is in Trial Mode"
            else:
                return "Windows Not Activated"
        except Exception as e:
            return f"Error: {e}"

    def get_disk_space(self, drive):
        try:
            disk_usage = psutil.disk_usage(drive)
            return f"Total: {disk_usage.total / (1024**3):.2f} GB, Used: {disk_usage.used / (1024**3):.2f} GB"
        except FileNotFoundError:
            return "Drive not found"
        except Exception as e:
            return f"Error: {e}"

    def get_ram_size(self):
        try:
            ram = psutil.virtual_memory()
            return f"{ram.total / (1024**3):.2f} GB"
        except Exception as e:
            return f"Error: {e}"

    def get_serial_number(self):
        return self.serial_number_text if hasattr(self, 'serial_number_text') else "Not Set"

    def get_user_input(self):
        user_input = simpledialog.askstring("ป้อนข้อมูล", "กรุณาป้อนข้อมูล Serial:")
        if user_input:
            self.serial_number_text = user_input
            self.serial_number_label.config(text=user_input)
            self.add_serial_button.grid_forget() # ซ่อนปุ่ม "Add Serial"
        else:
            messagebox.showinfo("ข้อมูล", "ไม่ได้ป้อนข้อมูล")

if __name__ == "__main__":
    app = App()
    app.mainloop()
