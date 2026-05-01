# DataKiller — Desktop ADC Signal & Waveform Analysis Tool

[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)]()

**DataKiller** is a lightweight desktop signal processing and analysis software built with Python + Tkinter + Matplotlib. It is specially designed for ADC sampling data analysis, time-domain waveform observation, noise evaluation and FFT frequency-domain analysis. With a complete graphical interface, users can finish data editing, signal preprocessing, mathematical analysis and report export without coding.

![Screenshot](screenshot.png) <!-- Replace with actual screenshot path -->

---

## ✨ Key Features

### 📊 Data Management & Generation
- Simulate test signals: single/multi-frequency sine wave, Gaussian white noise, noisy sine wave
- Data import: support TXT and XLSX format files
- Table editing: add/delete rows, double-click cell modification, paste from clipboard
- Data export: export raw data, selected region data and analysis results to Excel

### 🔧 Signal Preprocessing
- ADC raw data conversion: convert sampling code to actual voltage (bit depth & reference voltage configurable)
- Moving average filter: customizable sliding window size for noise reduction
- Manual region selection: intercept specified waveform segments for local analysis

### 📈 Time-Domain Analysis
- Mouse interactive plot: drag to pan, scroll wheel to zoom
- Big data downsampling: ensure smooth rendering for massive sampling points
- Linear fitting: least-square fitting, output fitting formula and correlation coefficient R
- Noise calculation: automatic calculation of Noise Vpp, 3σ and 6σ noise

### 🌊 FFT Frequency-Domain Analysis
- Rich FFT configuration: linear / dB amplitude mode
- Multiple window functions: Rectangular, Hanning, Hamming, Blackman
- Sampling frequency: auto calculation or manual input
- Standard single-sided spectrum with amplitude correction

### ⚡ User Experience Optimization
- Multi-thread file loading to prevent UI freezing
- Built-in Matplotlib toolbar for screenshot and fine adjustment
- Clear mode switching and status management
- Simple and intuitive GUI, easy for beginners to use

---

## 🛠️ Tech Stack

| Category          | Libraries / Frameworks           |
| ----------------- | -------------------------------- |
| GUI Framework     | tkinter / ttt                    |
| Data Visualization| matplotlib                       |
| Scientific Computing | numpy, scipy.fftpack         |
| File I/O & Export | openpyxl, xlsxwriter             |
| Concurrency       | threading, queue                 |

---

## 📥 Installation & Run

### Requirements
- Python 3.8 or higher

### Install Dependencies

```bash
pip install numpy scipy matplotlib openpyxl xlsxwriter
🎯 Application Scenarios
ADC sampling data calibration and performance testing

Circuit signal noise & waveform quality analysis

Electronic engineering, embedded hardware debugging

Teaching demonstration of signal processing and FFT algorithm

Laboratory experimental data analysis and report output

📄 Author Info
Author: Putishapohe

Location: Shanghai

Version: V8

📝 Update Log (V8)
Add detailed code comments

Optimize code structure & modular design

Optimize multi-thread import to avoid UI stuck

Upgrade FFT module with multiple window functions & dB mode

Support one-click test data generation

📌 License
Open-source for learning and communication. For commercial use, please contact the author.

https://img.shields.io/badge/License-MIT-yellow.svg
The code is licensed under MIT, but commercial use is kindly requested to contact the author.

DataKiller — 桌面级 ADC 信号与波形分析工具
https://img.shields.io/badge/License-MIT-yellow.svg
https://img.shields.io/badge/python-3.8+-blue.svg
https://img.shields.io/badge/platform-Windows%2520%257C%2520macOS%2520%257C%2520Linux-lightgrey

DataKiller 是一款轻量级桌面信号处理与分析软件，基于 Python + Tkinter + Matplotlib 构建。专为 ADC 采样数据分析、时域波形观察、噪声评估及 FFT 频域分析而设计。通过完整的图形界面，用户无需编写代码即可完成数据编辑、信号预处理、数学分析和报告导出。

https://screenshot.png <!-- 可替换为实际截图路径 -->

✨ 核心特性
📊 数据管理与生成
仿真信号生成：单频/多频正弦波、高斯白噪声、含噪正弦波

数据导入：支持 TXT 与 XLSX 格式文件

表格编辑：增删行、双击单元格修改、从剪贴板粘贴

数据导出：将原始数据、选定区域数据、分析结果导出至 Excel

🔧 信号预处理
ADC 原始数据转换：将采样码转换为实际电压（可配置位宽和参考电压）

移动平均滤波：可自定义滑动窗口大小，有效降噪

手动区域选择：截取指定波形片段进行局部分析

📈 时域分析
鼠标交互绘图：拖拽平移、滚轮缩放

大数据降采样：保障海量采样点下的流畅渲染

线性拟合：最小二乘法拟合，输出拟合公式与相关系数 R

噪声计算：自动计算噪声峰峰值 Vpp、3σ 和 6σ 噪声值

🌊 FFT 频域分析
丰富的 FFT 配置：线性幅值 / dB 幅值模式

多种窗函数：矩形窗、汉宁窗、海明窗、布莱克曼窗

采样频率：自动计算或手动输入

标准化单边频谱：支持幅值修正

⚡ 用户体验优化
多线程文件加载，杜绝界面卡死

内置 Matplotlib 工具栏，支持截图与精细调节

清晰的模式切换与状态管理

界面简洁直观，对初学者友好

🛠️ 技术栈
类别	库/框架
GUI 框架	tkinter / ttk
数据可视化	matplotlib
科学计算	numpy, scipy.fftpack
文件解析与导出	openpyxl, xlsxwriter
并发处理	threading, queue
📥 安装与运行
环境要求
Python 3.8 或更高版本

安装依赖
bash
pip install numpy scipy matplotlib openpyxl xlsxwriter
运行程序
bash
python main.py
💡 提示：建议在虚拟环境中运行，避免依赖冲突。

🎯 应用场景
ADC 采样数据校准与性能测试

电路信号噪声与波形质量分析

电子工程、嵌入式硬件调试

信号处理与 FFT 算法教学演示

实验室实验数据分析与报告输出

📄 作者信息
作者：Putishapohe

所在地：上海

版本：V8

📝 更新日志 (V8)
✅ 添加详细代码注释

✅ 优化代码结构与模块化设计

✅ 优化多线程导入，避免界面卡死

✅ 升级 FFT 模块，支持多种窗函数与 dB 模式

✅ 支持一键生成测试数据

📌 声明与许可
本项目仅用于学习交流。如需商业使用，请联系作者获得授权。

https://img.shields.io/badge/License-MIT-yellow.svg
代码开源协议为 MIT，但商业使用前建议与作者沟通。
