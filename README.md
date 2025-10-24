# 🧩 Member Offset Application Tool for STAAD.Pro

<div align="center">

### 🚀 Automate Member Offset Assignments in STAAD.Pro  
**Powerful VBA script to analyze and automatically apply offsets** to structural members based on node coordinates and section properties.

*Developed by* **Engr. Lowrence Scott D. Gutierrez**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?style=flat&logo=linkedin)](https://www.linkedin.com/in/lsdg)

---

### 📊 Project Stats

![GitHub Views](https://komarev.com/ghpvc/?username=SC0L0W&label=Repository%20Views&color=0e75b6&style=flat)  
![GitHub Stars](https://img.shields.io/github/stars/SC0L0W/STAADToolAuto-Offset?style=flat&color=yellow)  
![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue?style=flat&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=flat)

---

Streamline your structural modeling workflow by automatically detecting member types and applying precise offsets based on node coordinates and section properties.

</div>

---

## ✨ What Makes This Tool Special

This **VBA macro** for **STAAD.Pro** intelligently reads member and node data, classifies members into beams and columns based on coordinate change patterns, and applies suitable offsets considering member support conditions and section details.  

It simplifies complex structural adjustments, reduces manual errors, and saves valuable engineering time — especially for projects with numerous members.

> **Note:**  
> - Currently supports **beams** only.  
> - **Columns** must be rectangular or circular.  
>   - Rectangular columns should have their depth aligned parallel to the global Z-axis.  
>   - See the “How to Use” section for orientation instructions.  

> **Important:**  
> - For steel sections to be recognized and processed, the file **`AISC_STEEL_DATABASE.xlsx`** must be in the same directory as your STAAD file.  
> - This Excel database is used to look up steel section properties.

---

## 🔮 Planned Improvements

- Seamlessly detect column orientation and support offsets  
- Apply offsets to steel columns  
- Improve accuracy and flexibility of offset calculations  

---

## 🧮 Special Notes

- Current version applies offsets **only to beams**.  
- For rectangular columns, ensure depth is aligned to **Z-axis**.  
- Section names should clearly indicate size/orientation (e.g., `Rect 0.40x0.25`).  
- Future updates will extend support to columns and more section types.

---

## 🚀 Key Features

| Category | Description |
|:--|:--|
| 🎯 **Automated Member Classification** | Detects beams/columns based on coordinate patterns; handles multi-axis members; supports bracing detection |
| 🧰 **Section Property Extraction** | Reads RECT, CIRC, STEEL, and custom section labels; retrieves STAAD section data; integrates Excel-based steel database |
| 🔧 **Intelligent Offset Calculation** | Applies offsets based on section depth and connection type; supports beams and columns; uses STAAD COM API |
| 🖥️ **Robust Error Handling** | Captures COM connection issues; reports process status; detailed debug logs |
| 📊 **Summary & Reporting** | Displays total members processed, offsets applied, and errors; ready for integration with automation workflows |

---

## ⚙️ Compatibility & Requirements

- **STAAD.Pro** (with VBA enabled)  
- **Windows OS**  
- **VBA Version:** 7.0+ (Excel VBA)  
- **External File:** `AISC_STEEL_DATABASE.xlsx` in the same directory as your `.std` file  

### 🔧 Prerequisites
- STAAD.Pro installed and accessible via COM automation  
- Basic understanding of STAAD modeling and VBA  
- Optional: external steel database for steel sections  

---

## 📝 Usage

#### **Method A: Direct Installation**

1. Open **STAAD.Pro**
2. Navigate to **Utilities** → **Tools** → **Customize**
3. Click the **Commands** tab
4. Click **New** to create a new command
5. **Configuration:**
   - **Name:** `STAADToolAuto-Offset`
   - **Command:** Browse to your `.bas` or `.vbs` file
   - **Icon:** Choose a recognizable icon (optional)
   - **Shortcut:** Assign a keyboard shortcut like `Ctrl+Shift+L` (optional)
6. Click **OK** to save
7. **Place the Steel Database**
   - Copy `AISC_STEEL_DATABASE.xlsx` to your STAAD project folder  

8. **Run the Macro**
   - STAAD → Tools → User Tools → select **STAADToolAuto-Offset**
   - The script will analyze members and apply offsets
   - Check debug window for logs and results


#### **Method B: Excel-Based Execution**

1. Open the included Excel file with the macro
2. Enable macros when prompted
3. Keep STAAD.Pro running with your model open
4. paste the STAADToolAuto-Offset.vbs code content then run.

---

## 🔍 How the Code Works

1. **Connect to STAAD.Pro** via COM automation.  
2. **Extract Member List** using `GetBeamList()` or sequential member IDs.  
3. **Analyze Each Member:**
   - Get start and end node IDs and coordinates.  
   - Read section label and dimensions.  
   - Determine member type (Beam/Column).  
   - Calculate offsets based on section properties.  
4. **Apply Offsets** using `CreateMemberOffsetSpec()` and `AssignMemberSpecToBeam()`.  
5. **Summarize Results** – reports processed members, applied offsets, and failures.


## 🧾 License & Credits

**License:** MIT License  

**Developed by:** Engr. **Lowrence Scott D. Gutierrez**  

**Special Thanks To:**
- STAAD.Pro Development Team  
- Structural Engineering Community  

---

> 💡 *Have suggestions or found a bug?*  
> Feel free to open an issue or contribute via pull request!

---
