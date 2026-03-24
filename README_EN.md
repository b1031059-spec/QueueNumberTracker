Last Updated: March 24, 2026

---

**Please follow the steps below to complete the environment setup and run the program:**

    1.Download Executable:
        Go to the Releases page and download the latest .exe file.

    2.Download Runtime Environment:
        Download _internal.zip and extract it.

    3.Configure Directory:
        Ensure that QueueNumberTracker_v1.0.exe and the extracted _internal folder are located in the same directory.

---

**Parameter Settings (hospital_doctors.yaml):**
The program relies on the hospital_doctors.yaml configuration file. Setting this up allows you to conveniently name and categorize saved data. However, the program will still function even without custom settings. To configure it, you can:

    Option A: Go directly into the _internal folder and modify the default hospital_doctors.yaml.

    Option B: Download the sample yaml file from the GitHub source code, modify it, and overwrite the hospital_doctors.yaml inside the _internal folder.

After configuration, please ensure the hospital_doctors.yaml format is correct. This file defines the following initialization behaviors:

    Language: Set the interface language; supports "zh" (Traditional Chinese) or "en" (English).

    Hospital List: Preset frequently used hospital names, URLs, and doctor names to enable quick switching within the GUI.

    lastBBox: The system automatically records the position and size (Bounding Box) of the last captured area, which will be applied automatically upon the next launch.

---

Running the Program
Important Precautions:

    Screen Sleep: This program acquires data via screen capturing. If you need to run it for a long duration, please adjust your computer's Auto-Sleep settings to prevent capture failure caused by the screen turning off.

    First Launch: Double-click QueueNumberTracker_v1.0.exe to launch. The program requires an internet connection during the first launch to download image recognition models (approx. 30–60 seconds). If the interface appears unresponsive during this time, please wait patiently.

---

GUI Operation Instructions:
This tool provides an intuitive five-page workflow:

    Page 1: Screen Capture Settings

        Operation: Use the sliders or input boxes on the right to adjust the capture area's coordinates and size.

        Preview: The left window displays the selected area in real-time. Ensure it covers the consultation number area.

    Page 2: Recognition & Schedule Settings

        File Naming: Confirm the Excel filename at the top.

        Recognition Test: Verify that the OCR result on the left accurately reads the numbers.

        Parameter Settings: Set the total operating duration and recognition interval on the right.

    Page 3: Operation Log

        Displays real-time running status and recognition records.

        Supports manual stop or automatic shutdown after the schedule completes.

        Data Storage: Upon launch, the program automatically creates a data folder in the current directory. All data will be saved here in Excel format.

    Page 4: Data Post-processing

        Noise Filtering: Check this to generate "clean" data by filtering out abnormal number jumps (creates an additional sheet without modifying original data).

        Threshold Settings: Configure the threshold for the filtering algorithm.

        Visualization: Check this to automatically generate analytical charts based on the data.

        Algorithm details: [Link (Pending Publication)]

    Page 5: Result Display

        Displays the execution results of the filtering algorithm (success status or specific error causes).

---

Technical Details:
    Operating System: Windows
    Programming Language: Python
    Recognition Engine: EasyOCR (Optical Character Recognition)
    Storage Format: Microsoft Excel (.xlsx)