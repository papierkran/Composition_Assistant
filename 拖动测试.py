import tkinter as tk

class DraggableWidget:
    def __init__(self, widget, grid_size=50, snap_threshold=20):
        self.widget = widget
        self.grid_size = grid_size
        self.snap_threshold = snap_threshold
        self.start_x = 0
        self.start_y = 0

        # 绑定事件
        widget.bind("<Button-1>", self.on_click)
        widget.bind("<B1-Motion>", self.on_drag)
        widget.bind("<ButtonRelease-1>", self.on_release)

    def on_click(self, event):
        # 记录点击时的相对位置
        self.start_x = event.x
        self.start_y = event.y

    def on_drag(self, event):
        # 计算新位置
        x = self.widget.winfo_x() + event.x - self.start_x
        y = self.widget.winfo_y() + event.y - self.start_y
        self.widget.place(x=x, y=y)

    def on_release(self, event):
        # 获取当前位置
        x = self.widget.winfo_x()
        y = self.widget.winfo_y()

        # 计算最近的网格点
        grid_x = round(x / self.grid_size) * self.grid_size
        grid_y = round(y / self.grid_size) * self.grid_size

        # 判断是否需要吸附
        if abs(grid_x - x) <= self.snap_threshold:
            x = grid_x
        if abs(grid_y - y) <= self.snap_threshold:
            y = grid_y

        self.widget.place(x=x, y=y)

# 创建主窗口
root = tk.Tk()
root.geometry("400x300")
root.title("Tkinter 拖动吸附示例")

# 创建可拖动的标签
label = tk.Label(root, text="拖我", bg="skyblue", width=10, height=2)
label.place(x=100, y=100)

# 启用拖动吸附
DraggableWidget(label, grid_size=50, snap_threshold=15)

root.mainloop()
