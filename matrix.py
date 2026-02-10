import sys
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# 1. 设置中文字体与矢量图基础配置
plt.rcParams['font.sans-serif'] = ['SimHei']  # 正常显示中文
plt.rcParams['axes.unicode_minus'] = False    # 正常显示负号
# 关键设置：导出矢量图时将文本转为路径，防止在没有字体的电脑上打开 PDF/SVG 时乱码
plt.rcParams['pdf.fonttype'] = 42
plt.rcParams['svg.fonttype'] = 'path'

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QFileDialog, QMessageBox, QGroupBox)
from PyQt5.QtCore import Qt

class MatrixPlotter(QWidget):
    def __init__(self):
        super().__init__()
        self.fileName = None
        self.initUI()

    def initUI(self):
        # 窗口基础设置
        self.setWindowTitle('功效矩阵分析工具 (支持矢量导出)')
        self.setGeometry(100, 100, 1000, 850)

        # 顶部控制面板
        self.ctrl_group = QGroupBox("控制面板")
        ctrl_layout = QHBoxLayout()
        ctrl_layout.setContentsMargins(20, 40, 20, 20) 
        ctrl_layout.setSpacing(15)

        self.btn_import = QPushButton("📂 导入 Excel")
        self.btn_plot = QPushButton("📊 生成矩阵图")
        self.btn_plot_t = QPushButton("🔄 生成转置图")
        self.btn_save = QPushButton("💾 导出图片/矢量图")

        # 初始状态
        self.btn_plot.setEnabled(False)
        self.btn_plot_t.setEnabled(False)
        self.btn_save.setEnabled(False)

        ctrl_layout.addWidget(self.btn_import)
        ctrl_layout.addWidget(self.btn_plot)
        ctrl_layout.addWidget(self.btn_plot_t)
        ctrl_layout.addWidget(self.btn_save)
        self.ctrl_group.setLayout(ctrl_layout)

        # 图表显示区
        self.figure = Figure(dpi=100)
        self.canvas = FigureCanvas(self.figure)

        # 底部状态栏
        self.status_lab = QLabel("提示：请先导入 Excel 数据。")
        self.status_lab.setStyleSheet("color: #555; font-size: 13px; padding: 5px;")

        # 总布局
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 10, 15, 10)
        layout.addWidget(self.ctrl_group)
        layout.addWidget(self.canvas, 1)
        layout.addWidget(self.status_lab)
        self.setLayout(layout)

        # 信号绑定
        self.btn_import.clicked.connect(self.import_file)
        self.btn_plot.clicked.connect(lambda: self.process_data(transpose=False))
        self.btn_plot_t.clicked.connect(lambda: self.process_data(transpose=True))
        self.btn_save.clicked.connect(self.save_image)

        self.apply_styles()

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget { font-family: "Microsoft YaHei"; background-color: #fcfcfc; }
            QGroupBox { 
                font-weight: bold; font-size: 15px;
                border: 2px solid #bbb; border-radius: 10px; 
                margin-top: 20px; background-color: white; 
            }
            QGroupBox::title {
                subcontrol-origin: margin; subcontrol-position: top left;
                left: 20px; padding: 0 5px; top: 2px;
            }
            QPushButton { 
                background-color: #3498db; color: white; border-radius: 6px; 
                padding: 10px; font-weight: bold; height: 28px; font-size: 14px;
            }
            QPushButton:hover { background-color: #2980b9; }
            QPushButton:disabled { background-color: #ddd; color: #888; }
            QPushButton#btn_save { background-color: #27ae60; }
        """)
        self.btn_save.setObjectName("btn_save")

    def import_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "选择数据", "", "Excel (*.xlsx *.xls)")
        if fname:
            self.fileName = fname
            self.status_lab.setText(f"当前文件：{fname}")
            self.btn_plot.setEnabled(True)
            self.btn_plot_t.setEnabled(True)

    def process_data(self, transpose=False):
        try:
            # 自动跳过空行，读取Excel
            df = pd.read_excel(self.fileName, index_col=0)
            if transpose: df = df.T
            self.draw_plot(df)
            self.btn_save.setEnabled(True)
            self.status_lab.setText("绘图成功，可以导出矢量图了")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"数据解析失败: {e}")

    def draw_plot(self, df):
        self.figure.clear()
        ax = self.figure.add_subplot(111)

        # 数据清洗：确保数值类型
        num_df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
        cols, rows = df.columns.tolist(), df.index.tolist()
        x_indices = list(range(1, len(cols) + 1))
        
        # 调色盘
        colors = ['#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F', 
                  '#EDC948', '#B07AA1', '#FF9DA7', '#9C755F', '#BAB0AC']

        # 1. 绘制气泡
        for i in range(len(rows)):
            row_vals = num_df.iloc[i].values
            sizes = [abs(float(v)) * 180 for v in row_vals] 
            y_coords = [i + 1] * len(x_indices)
            ax.scatter(x_indices, y_coords, s=sizes, c=colors[i % len(colors)], 
                       alpha=0.6, edgecolors='white', zorder=2)

        # 2. 绘制带背景遮挡的数值文字
        offset = 0.15 
        for i in range(len(rows)):
            for j, val in enumerate(df.iloc[i].values):
                ax.text(x_indices[j] + offset, (i + 1) + offset, str(val),
                        ha='left', va='bottom',
                        fontsize=10, fontweight='bold', color='black',
                        zorder=10,
                        bbox=dict(
                            facecolor='white',
                            alpha=0.7,
                            edgecolor='none',
                            boxstyle='round,pad=0.2'
                        ))

        # 3. 细节美化
        ax.set_xticks(x_indices)
        ax.set_xticklabels(cols, rotation=30, ha='right')
        ax.set_yticks(range(1, len(rows) + 1))
        ax.set_yticklabels(rows)
        
        # 设置坐标轴范围，防止文字溢出
        ax.set_xlim(0.4, len(cols) + 1.0)
        ax.set_ylim(0.4, len(rows) + 1.0)
        
        ax.set_title("功效矩阵分析图", fontsize=16, fontweight='bold', pad=25)
        ax.grid(True, linestyle='--', alpha=0.3, zorder=0)

        self.figure.tight_layout()
        self.canvas.draw()

    def save_image(self):
        # 定义支持的格式，包括矢量图 PDF, SVG 和位图 PNG, JPG
        file_filter = (
            "PDF 矢量图 (*.pdf);;"
            "SVG 矢量图 (*.svg);;"
            "PNG 高清图片 (*.png);;"
            "JPG 压缩图片 (*.jpg)"
        )
        path, selected_filter = QFileDialog.getSaveFileName(self, "保存图表", "矩阵分析结果", file_filter)
        
        if path:
            try:
                # 检查是否为 SVG 并特殊处理
                if path.lower().endswith('.svg'):
                    plt.rcParams['svg.fonttype'] = 'path'
                
                # 保存文件，dpi 对位图有效，对矢量图主要影响嵌入的某些效果
                self.figure.savefig(path, dpi=300, bbox_inches='tight')
                QMessageBox.information(self, "成功", f"文件已成功保存至：\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存文件失败: {e}")

if __name__ == '__main__':
    # 适配高分屏
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    window = MatrixPlotter()
    window.show()
    sys.exit(app.exec_())
