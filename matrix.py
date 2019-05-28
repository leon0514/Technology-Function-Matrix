import sys
# import PyQt5.sip
import matplotlib.pyplot as plt  # 导入Python的绘图扩展包matplotlib，并重新命名为plt
import pandas as pd  # 导入python的数据处理扩展包pandas，并重命名为pd，该包用于读写excel文件
# import PyQt5.sip
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
from PyQt5.QtWidgets import QFileDialog


class Matrix(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Matrix')
        self.setFixedSize(280, 150)
        self.initUI()

    def initUI(self):
        self.layout = QVBoxLayout()
        self.layout1 = QVBoxLayout()
        self.layout2 = QHBoxLayout()

        self.lab = QLabel("生成功效矩阵图小工具", self)
        self.btn1 = QPushButton("导入excel", self)
        self.btn1.clicked.connect(self.getfile)
        self.btn2 = QPushButton("生成图", self)
        self.btn2.clicked.connect(self.getMatrix)
        self.btn3 = QPushButton("生成转置图", self)
        self.btn3.clicked.connect(self.getT)
        self.layout.addLayout(self.layout1)
        self.layout.addLayout(self.layout2)
        self.layout1.addWidget(self.lab)
        self.layout2.addWidget(self.btn1)
        self.layout2.addWidget(self.btn2)
        self.layout2.addWidget(self.btn3)
        self.setLayout(self.layout)

    def getfile(self):
        global fileName
        global filetype
        fileName, filetype = QFileDialog.getOpenFileName(self,
                                                         "选取文件",
                                                         "./",
                                                         "Text Files (*.xls?)")  # 设置文件扩展名过滤,注意用双分号间隔

    def getMatrix(self):
        try:
            data1 = pd.read_excel(fileName)  # 从excel文件中读取数据，并保存到data1变量中
        except:
            return
        color = ['b', 'g', 'r', 'c', 'm', 'y', 'gold', 'k', 'navy', 'purple', 'aqua', 'saddle', 'brown', 'brown',
                 'tomato', 'lightskyblue']
        xname = data1.columns  # 获取数据表的列，作为x轴数据
        xname.tolist()  # 转化成Python列表，方便绘图
        xz = list(range(1, len(xname) + 1))

        yname = data1.T.columns
        yz = list(range(1, len(yname) + 1))

        try:
            y_index = []  # 声明一个空列表
            for i in yname:  # 循环访问列元素的每个索引
                y_index.append(i)  # 将列元素的索引加入列表
                y = [[i for y_data in range(len(xname))] for i in range(1, len(y_index) + 1)]  # 从列表数据中生成y轴坐标矩阵

            for i in range(len(y_index)):  # 画圈
                s = data1.ix[i].values
                s = [i * 10 for i in s]
                plt.scatter(xz, y[i], s=s, c=color[i % 15], alpha=0.5)

            for i in range(len(yname)):  # 画数字
                s = data1.ix[i].values
                for m, n in zip(xz, y[i]):
                    plt.text(m, n, s[m - 1], family='fantasy', fontsize=12,
                             style='italic', color='mediumvioletred')

            plt.xticks(xz, xname, fontproperties='SimHei', fontsize=14)
            plt.yticks(yz, yname, fontproperties='SimHei', fontsize=14)
            plt.grid(True)
            plt.show()  # 显示绘制的图像
        except:
            return

    def getT(self):
        try:
            data1 = pd.read_excel(fileName).T  # 从excel文件中读取数据，并保存到data1变量中
        except:
            return
        color = ['b', 'g', 'r', 'c', 'm', 'y', 'gold', 'k', 'navy', 'purple', 'aqua', 'saddle', 'brown', 'brown',
                 'tomato','lightskyblue']
        xname = data1.columns  # 获取数据表的列，作为x轴数据
        xname.tolist()  # 转化成Python列表，方便绘图
        xz = list(range(1, len(xname) + 1))

        yname = data1.T.columns
        yz = list(range(1, len(yname) + 1))

        try:
            y_index = []  # 声明一个空列表
            for i in yname:  # 循环访问列元素的每个索引
                y_index.append(i)  # 将列元素的索引加入列表
                y = [[i for y_data in range(len(xname))] for i in range(1, len(y_index) + 1)]  # 从列表数据中生成y轴坐标矩阵

            for i in range(len(y_index)):  # 画圈
                s = data1.ix[i].values
                s = [i * 10 for i in s]
                plt.scatter(xz, y[i], s=s, c=color[i % 15], alpha=0.5)

            for i in range(len(yname)):  # 画数字
                s = data1.ix[i].values
                for m, n in zip(xz, y[i]):
                    plt.text(m, n, s[m - 1], family='fantasy', fontsize=12,
                             style='italic', color='mediumvioletred')

            plt.xticks(xz, xname, fontproperties='SimHei', fontsize=14)
            plt.yticks(yz, yname, fontproperties='SimHei', fontsize=14)
            plt.grid(True)
            plt.show()  # 显示绘制的图像
        except:
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Matrix()
    ex.show()
    sys.exit(app.exec_())
