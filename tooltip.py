import tkinter as tk

AUTO_CLOSE = 3000

class Tooltip:
    def __init__(self, widget, default: str):
        self.default = default
        self.widget = widget
        self.tip_window = None
        self.after_id = None

        widget.bind('<Enter>', self.schedule_tip)
        widget.bind('<Leave>', self.cancel_tip)

    def schedule_tip(self, event):
        self.after_id = self.widget.after(500, self.show)

    def show(self):
        if self.tip_window:
            return

        if self.widget.cget('text') == self.default:
            return

        self.tip_window = tk.Toplevel(self.widget)

        x, y = self.widget.master.winfo_pointerxy()
        x += 12
        y -= 25

        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f'+{x}+{y}')

        tip = tk.Label(
            self.tip_window,
            text=self.widget.cget('text'),
            bg='#FFFFE0',
            fg='black',
            relief='solid',
            borderwidth=0,
            padx=5,
            pady=2
        )

        tip.pack()
        self.tip_window.after(AUTO_CLOSE, self.hide)

    def hide(self):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def cancel_tip(self, event):
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

        self.hide()
