# Interactive Excel Data Processor (MATLAB)

This MATLAB script provides a simple, user-friendly graphical interface (UI) to process Excel (`.xlsx`) data files. It is designed to easily read, truncate, average, and downsample frequency datasets.

## 🌟 Key Features

* **Interactive File Selection:** Choose your Excel file directly through a file explorer dialog.
* **Multi-Sheet Processing:** Select one or multiple sheets (runs) from your workbook to average them together.
* **Automated Header Detection:** Automatically scans rows to find the "Hz" cell, determining exactly where your data starts.
* **Frequency Filtering:** Set custom frequency ranges (Start and End Hz) to crop the data as needed.
* **Downsampling:** Automatically interpolates the data to new frequency steps based on your input.
* **Clean Export:** The final processed data is neatly saved as a `.csv` file in a dedicated `Processed_Data` folder.

## 🚀 How to Use

1. Open and run the script in your MATLAB environment.
2. A dialog window will appear; select your target `.xlsx` file.
3. Choose the sheets you want to process from the provided list.
4. Review the confirmation message detailing where your data starts.
5. Enter your desired downsampling step (leave blank to keep the original frequency steps) and frequency limits.
6. Click OK, and wait for the process to finish! A clean `.csv` file will be generated automatically.

## 🛠️ System Requirements
* MATLAB (utilizes modern functions like `readmatrix` and `sheetnames`).
* The target Excel data file must have Column 1 as Frequency and Column 2 as Data. The frequency column header must contain the exact string "Hz".
