# ExcelFontFix
## 介绍

***ExcelFontFix***: 修复Excel当中Unicode字符所导致的中文错误字体

<p align="center"> 
  <a href="https://github.com/SwimmingLiu/ExcelFontFix/blob/master/README.md"> English</a> &nbsp; | &nbsp; 简体中文</a>
 </p>


![ExcelFontFix-Screen](https://oss.swimmingliu.cn/screenshot_ExcelFontFix.gif)

## 准备工作 

### 1. 创建虚拟环境

创建内置Python 3.9的conda虚拟环境, 然后激活该环境.

```shell
conda create -n excelfontfix python=3.9
conda activate excelfontfix
```

### 2. 安装依赖包

切换到YOLOSHOW程序所在的路径

```shell
cd {YOLOSHOW程序所在的路径}
```

安装程序所需要的依赖包

```shell
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 3. 运行程序

``` python
python main.py
```

## 框架

[![Python](https://img.shields.io/badge/python-3776ab?style=for-the-badge&logo=python&logoColor=ffd343)](https://www.python.org/)[![Static Badge](https://img.shields.io/badge/Pyside6-test?style=for-the-badge&logo=qt&logoColor=white)](https://doc.qt.io/qtforpython-6/PySide6/QtWidgets/index.html)

## 参考

[PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets)
