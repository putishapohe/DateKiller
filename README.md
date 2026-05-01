DataKiller
A Desktop ADC Signal & Waveform Analysis Tool
DataKiller is a lightweight desktop signal processing and analysis software built with Python + Tkinter + Matplotlib.
It is specially designed for ADC sampling data analysis, time-domain waveform observation, noise evaluation and FFT frequency-domain analysis.
With a complete graphical interface, users can finish data editing, signal preprocessing, mathematical analysis and report export without coding.
📌 Key Features
1. Data Management & Generation
Simulate test signals: single/multi-frequency sine wave, Gaussian white noise, noisy sine wave
Data import: support TXT and XLSX format files
Table editing: add/delete rows, double-click cell modification, paste from clipboard
Data export: export raw data, selected region data and analysis results to Excel
2. Signal Preprocessing
ADC raw data conversion: convert sampling code to actual voltage (bit depth & reference voltage configurable)
Moving average filter: customizable sliding window size for noise reduction
Manual region selection: intercept specified waveform segments for local analysis
3. Time-Domain Analysis
Mouse interactive plot: drag to pan, scroll wheel to zoom
Big data downsampling: ensure smooth rendering for massive sampling points
Linear fitting: least-square fitting, output fitting formula and correlation coefficient R
Noise calculation: automatic calculation of Noise Vpp, 3σ and 6σ noise
4. FFT Frequency-Domain Analysis
Rich FFT configuration: linear / dB amplitude mode
Multiple window functions: Rectangular, Hanning, Hamming, Blackman
Sampling frequency: auto calculation or manual input
Standard single-sided spectrum with amplitude correction
5. User Experience Optimization
Multi-thread file loading to prevent UI freezing
Built-in Matplotlib toolbar for screenshot and fine adjustment
Clear mode switching and status management
Simple and intuitive GUI, easy for beginners to use
🛠 Tech Stack
GUI Framework: tkinter / ttk
Data Visualization: matplotlib
Scientific Computing: numpy, scipy.fftpack
File Parsing & Export: openpyxl, xlsxwriter
Concurrent Processing: threading, queue
📥 Installation & Run
1. Install Dependencies
bash
运行
pip install numpy scipy matplotlib openpyxl xlsxwriter
2. Run Project
bash
运行
python main.py
🎯 Application Scenarios
ADC sampling data calibration and performance testing
Circuit signal noise & waveform quality analysis
Electronic engineering, embedded hardware debugging
Teaching demonstration of signal processing and FFT algorithm
Laboratory experimental data analysis and report output
📄 Author Info
Author: Putishapohe
Location: Shanghai
Version: V8 Update
📝 Update Log
V8 Update
Add detailed code comments
Optimize code structure & modular design
Optimize multi-thread import to avoid UI stuck
Upgrade FFT module with multiple window functions & dB mode
Support one-click test data generation
📌 License
Open-source for learning and communication.
For commercial use, please contact the author.
