# Impedance Tube SAC Data Processor (MATLAB)

This MATLAB script provides a simple, user-friendly graphical interface (UI) specifically designed to process **Sound Absorption Coefficient (SAC)** data obtained from **Impedance Tube testing**. It allows researchers and engineers to easily read, truncate, average, and downsample multi-run frequency datasets exported in Excel (`.xlsx`) format.

## 🌟 Key Features

* **Interactive File Selection:** Choose your Excel file directly through a file explorer dialog.
* **Multi-Run Processing:** Select one or multiple sheets (e.g., Run 1, Run 2, Run 3) from your workbook to average them together for a more accurate final SAC curve.
* **Automated Header Detection:** Automatically scans rows to find the "Hz" unit block, determining exactly where the measurement data starts while ignoring the pre-test metadata.
* **Frequency Filtering:** Set custom frequency ranges (Start and End Hz) to crop the data to your specific range of interest (e.g., focusing only on 50 Hz - 5600 Hz).
* **Downsampling:** Automatically interpolates the data to new frequency steps based on your input.
* **Clean Export:** The final processed and averaged data is neatly saved as a `.csv` file in a dedicated `Processed_Data` folder.

## 📊 Expected Data Format

The script is built to handle standard Impedance Tube exports. 
* The source file should be an `.xlsx` workbook where each test run is on a separate sheet.
* The script bypasses the system metadata at the top of the file (such as sample frequency, resolution, transducer details, etc.).
* **Data Columns:** It expects the raw measurement data to be structured with **Frequency (Hz)** in Column A and the **SAC Ratio / Complex Value** in Column B. 

## 🚀 How to Use

1. Open and run the script in your MATLAB environment.
2. A dialog window will appear; select your target `.xlsx` file containing the Impedance Tube data.
3. Choose the specific sheets (runs) you want to process and average from the provided list.
4. Review the confirmation message detailing where your data starts.
5. Enter your desired downsampling step (leave blank to keep the original frequency steps) and your target frequency limits.
6. Click OK, and wait for the process to finish! A clean, averaged `.csv` file will be generated automatically.

## 🛠️ System Requirements
* MATLAB (utilizes modern functions like `readmatrix` and `sheetnames`).
* The frequency column header must contain the exact string "Hz" for the automatic detection to work correctly.

* Author: Amirul Azhajzul
* Institution: USM
* 2026
