最後更新時間: 2026/3/16

**Follow the steps below to complete the environment setup and run the program:**

1.Download Executable:
    Go to the Releases page and download the latest .exe file and relevant configuration files.。

2.Configure Directory:
    Place the downloaded executable in the same directory as the hospital_doctors.yaml file.

3.Run Program:
    Double-click the .exe file to launch the Graphical User Interface (GUI).

4.Data Storage:
    After the program starts, it will automatically create a data folder in the current directory. All data will be saved here in Excel format.

---

**Parameter Settings (hospital_doctors.yaml):**
Before running the program, ensure that hospital_doctors.yaml is correctly formatted. This file defines the initialization behavior:

    language:
        Set the interface language; supports "zh" (Traditional Chinese) or "en" (English).

    Hospital List:
        You can preset frequently used hospital names, URLs, and doctor names according to the format to enable quick switching within the GUI.

    lastBBox:
        The system automatically records the position and size (Bounding Box) of the last captured screen. These will be applied automatically next time, so there is no need to readjust.

---

**GUI Operation Instructions:**
This tool provides an intuitive five-page workflow:

    Page 1: Screen Capture Settings
        Operation: Use the sliders or input boxes on the right to adjust the coordinates and size of the capture area.

        Preview: The left window displays the currently selected area in real-time. Please ensure it covers the consultation number area.

    Page 2: Recognition and Schedule Settings
        File Naming: Confirm the filename for the Excel storage at the top.

        Recognition Test: Confirm whether the image recognition result on the left accurately reads the numbers.

        Parameter Settings: Set the total operating time and recognition interval on the right.

    Page 3: Operation Log
        Displays the program's running status and recognition records in real-time.

        Supports manual stop or automatic shutdown after the schedule ends.

    Page 4: Data Post-processing
        Noise Filtering: Check this to generate "clean" data by filtering out abnormal number jumps (this creates an additional sheet without modifying the original data).

        Parameter Settings: Set the threshold for the filtering algorithm.

        Visualization: Check this to automatically generate analytical charts based on the data.

        For detailed logic regarding the filtering algorithm, please refer to: [Link]

    Page 5: Result Display
        Displays the execution results of the filtering algorithm (success status or cause of error).

---
        
Technical Details:
    Operating System: Windows
    Programming Language: Python
    Recognition Engine: EasyOCR (Optical Character Recognition)
    Storage Format: Microsoft Excel (.xlsx)