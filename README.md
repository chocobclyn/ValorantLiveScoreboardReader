# VALORANT LIVE SCOREBOARD READER

This project contains two Python scripts (`ValScoreboard.py` and `ValWS.py`) for extracting and matching data from the in-game Valorant scoreboard. **Note: I will not offer support for the Python scripts as I'm not an expert, so proceed with this understanding.**

### Roadblock:
A current limitation of the project is with Excel recalculations. The extracted data is saved to an Excel file, but **Excel does not update or recalculate formulas while closed**. This affects the intended workflow as I am unable to force Excel to recalculate automatically. You may encounter issues if the file remains open while running the scripts, as Excel locks the file.

## Installation Instructions

### Prerequisites
Make sure the following dependencies are installed on your system:
1. **Python 3.12.6 or higher**
2. **NVIDIA CUDA Toolkit**: Download and install from the [NVIDIA CUDA website](https://developer.nvidia.com/cuda-downloads). This is required if you're running the project with GPU acceleration.
3. **cuDNN (CUDA Deep Neural Network)**: Download from the [NVIDIA cuDNN website](https://developer.nvidia.com/cudnn). Follow the instructions for copying the necessary cuDNN files to your CUDA directory.

### Python Dependencies

After setting up CUDA/cuDNN (if applicable), install the required Python libraries using pip. Run the following command in your terminal:

```bash
pip install paddlepaddle paddleocr opencv-python mss openpyxl numpy
```

## Project Files
### 1. ValScoreboard.py
Purpose: Extracts in-game scoreboard text (names, kills, assists, deaths, etc.) using OCR (PaddleOCR). The results are then saved to an Excel file (scoreboard.xlsx).
  1. Key Features:
  - Uses multiple regions of interest for OCR extraction.
  - Preprocesses images before running OCR to enhance accuracy.
  - Data is saved to the live sheet of scoreboard.xlsx.

### 2. ValWS.py
Purpose: Matches weapons and shields from the scoreboard using template matching. Extracts the player's loadout (weapon and shield type) and saves this data to the loadout sheet of scoreboard.xlsx.
  1. Key Features:
  - Processes image regions and uses pre-defined templates for matching weapons and shields
  - Each player’s loadout (weapon and shield) is recorded in the corresponding Excel file.

## Running the Project
### Step 1: Start the Scripts

To run the project sequentially (first ValScoreboard.py then ValWS.py), use the provided GUI interface (Tkinter). When the Start button is pressed, the scripts run at specified intervals, e.g., every 3, 5, or 10 seconds. The Stop button halts the execution.
### Step 2: Update Excel File

The extracted data is saved to an Excel file, but if the file is opened while running the scripts, Excel will lock the file, and updates will fail. Ensure the file is closed or explore alternatives like saving to CSV (although this doesn’t support formula recalculation).
