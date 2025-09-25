# Cybear404 AutoCut Optimizer

**Plan smarter cuts. Waste less material. Quote with confidence.**

Cybear404 AutoCut Optimizer helps shops, fabricators, and builders minimize
material waste when cutting rods, pipes, bars, or tubes. Instead of manually
figuring out how many pieces fit in each stock length, the app groups your cut
list by material and stock length, accounts for saw kerf, and proposes an
efficient cut plan—plus clear reports you can share with your team or customer.

---

### Why it’s helpful
- **Reduce scrap & cost** — Packs more cuts per bar by accounting for kerf and leftover space.  
- **Order the right amount** — Summaries and procurement totals show how many bars to buy per material/length.  
- **Quote faster** — Generate a clean plan and reports you can attach to estimates.  
- **Catch problems early** — Flags oversize requests and feasibility issues before you cut.  
- **Simple Excel workflow** — Works with a single sheet (`Cut Length`, `Max Length`, `Material Type`).  

---

### How it works (at a glance)
1. **Load** your Excel sheet or generate sample data from the menu.  
2. **Set kerf** (fraction like `1/8` or decimal like `0.125`) and pick optional reports.  
3. **Run optimizer** — The tool creates a “Grouped Cuts” sheet and (optionally) Summary, Validation, Waste Report, Procurement, and Issues sheets.  

---

### What you get
- **Grouped Cuts** — Per stock length + material, showing how pieces are assigned to each bar.  
- **Summary** — Total bars used, total cut length, kerf consumed, estimated waste, utilization.  
- **Procurement** — Bars to order by material and stock length.  
- **Validation** — Quick checks for feasibility.  
- **Waste Report** — Per-bar breakdown of used length, kerf, leftover, utilization.  
- **Issues** — Oversize or unplaceable cuts (if any).  

> **Note:** Light/dark theme toggle is included but still **experimental**.

---

## Intended Use
- Demonstration / proof-of-concept for construction, fabrication, and shop workflows  
- Free to download and use for testing/learning  
- Not certified for production-critical environments (use at your own risk)  

---

## Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/Cybear404/cybear404-autocut-optimizer.git
cd cybear404-autocut-optimizer
pip install -r requirements.txt
```

Run with:

```bash
python app.py
```

---

## Packaging

To build a standalone executable (Windows example):

```bash
pyinstaller --onefile --windowed app.py -n AutoCutOptimizer
```

On macOS or Linux, similar PyInstaller commands apply.  

---

## Branding / Logo

This app ships with a **neutral placeholder logo** (`assets/logo.png`).  
You’re free to replace it with your own branding.  

⚠️ Branding shown in screenshots (Cybear404 name/logo) is **for demonstration only**  
and is **not licensed for reuse**. See **BRANDING.md** for details.  

---

## License

This project is licensed under the **PolyForm Noncommercial 1.0.0** license.  

- ✅ Personal, educational, or internal company use is free.  
- ❌ You may **not** sell, offer as a paid service, or bundle into a commercial product.  
- For **commercial licensing**, please contact Cybear404, LLC (see COMMERCIAL-LICENSE.md).  

Branding and trademarks (Cybear404 name and logo) are **not open-licensed** and may not be used without permission.  
See **BRANDING.md**.

---

## Author

**Cybear404, LLC**  
Contributors: Developed with assistance from OpenAI’s ChatGPT  

Website: https://cybear404.com  
GitHub: https://github.com/Cybear404  

Questions or ideas? Open an issue or reach out through the website.
