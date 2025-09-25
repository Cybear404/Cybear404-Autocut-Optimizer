# 🚀 Cybear404 AutoCut Optimizer v0.1.0
**Proof-of-Concept Release**

## ✨ Features
- Modern GUI built with PySide6 / Qt and Cybear404 branding  
- Reads/writes Excel `.xlsx` via pandas + openpyxl  
- Optimizes cuts by stock length and material with bin-packing logic  
- Configurable saw kerf (fractional or decimal input)  
- Options to overwrite or copy the workbook  
- Optional reports:  
  - **Summary** (overview of cuts, waste, utilization)  
  - **Procurement** (bars required per material/length)  
  - **Validation** (checks feasibility, flags issues)  
  - **Waste Report** (per-bar breakdown)  
  - **Issues** (oversize or unplaceable cuts)  
- Built-in sample/template generator for quick testing  
- Settings persistence and “Open Output File” button  

---

## 🧪 Status
This is an **early proof-of-concept**.  
- ✅ Core optimizer works and generates output Excel sheets.  
- ⚠️ Light/dark theme toggle is included but **not fully working** yet.  
- ⚠️ Some edge cases (e.g., very large cuts, unusual Excel formats) may produce warnings or issues.  

---

## 📥 Installation
Clone the repo and install requirements:
```bash
git clone https://github.com/Cybear404/cybear404-autocut-optimizer.git
cd cybear404-autocut-optimizer
pip install -r requirements.txt
python app.py
```

---

## 📌 Roadmap
- Fix and polish theme toggle (light/dark mode)  
- Add packaged builds (.exe for Windows, .app for macOS)  
- More detailed error handling & user guidance  
- Explore cloud integration / API backend  

---

## 🙏 Acknowledgments
Developed by **Cybear404, LLC**  
Contributors: with assistance from **OpenAI’s ChatGPT**
