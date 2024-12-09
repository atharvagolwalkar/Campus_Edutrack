# Campus EduTrack

  
Campus EduTrack is a data processing tool designed to automate the population of student data from a central database into a set of pre-designed Excel templates. The application allows for efficient processing of personal information and semester results for students, handling data for multiple students at a time.

The tool leverages **xlwings** for Excel automation, **Tkinter** for the GUI, and **PIL** for image handling, providing an easy-to-use interface to select files, define processing options, and review output in real-time.


## Features


-  **Batch Processing:** Populate personal information for multiple students at once using data from a central Excel database.

-  **Excel Automation:** Automate the filling of pre-designed Excel templates with student information for personal details and results for different semesters.

-  **Custom Operations:** Select from various operations like filling Page 1 and Page 2 of personal information or processing semester results (1-8).

-  **Interactive GUI:** Built using Tkinter and CustomTkinter for a user-friendly interface with custom buttons, entries, and output display.

-  **File Handling:** Load database files and templates, specify row ranges for processing, and save output files automatically.

-  **Error Handling:** Provides detailed error messages to troubleshoot issues during data processing.


## Technologies Used
  
-  **Python**: Core programming language for the tool.

-  **Tkinter**: For building the graphical user interface.

-  **CustomTkinter**: Enhanced styling and customization for Tkinter components.

-  **PIL (Python Imaging Library)**: For image handling in the GUI.

-  **xlwings**: For automating Excel tasks such as reading, writing, and copying data.

-  **re**: Regular expressions for formatting filenames.

 
## Installation

To use Campus EduTrack, you will need to install the following dependencies:

```bash
pip  install  xlwings
```
```bash
pip  install  tkinter
```
```bash
pip  install  Pillow
```
```bash
pip  install  customtkinter
```

Make sure you have **Microsoft Excel** installed, as xlwings uses Excel to perform the automation.

## How to Use

1.  **Run the Application**: Start the application by running the main Python script:
    
    ```bash
    python Campus_Edu.py
    ```
2.  **Browse for Database**: Use the `Browse Database` button to select the Excel file that contains student information.
    
3.  **Enter Row Range**: Define the starting and ending rows for the students you want to process.
    
4.  **Select Operation**: Choose the operation you want to perform:
    
    -   Fill Page 1 and Page 2 (Personal Information)
    -   Process Semester 1–8
    -   Process Extracurricular Activities
5.  **Process the Operation**: Click the `Process Selected Operation` button to execute the selected task. The progress will be displayed in the output text box.
    
6.  **Output Files**: The processed Excel files will be saved in the `output` folder.
    

## GUI Overview

-   **Browse Database**: Opens a file dialog to select the student database Excel file.
-   **File Path Entry**: Displays the path of the selected database file.
-   **Start/End Row**: Allows input of row numbers to specify the range of students to process.
-   **Select Operation**: Dropdown to choose the operation you want to perform.
-   **Process Button**: Starts the selected operation.
-   **Output Display**: Shows real-time status and messages related to the processing.

## Error Handling

If any issues occur during the processing, the application will display an error message in a dialog box and provide details in the output text widget.

Common errors include:

-   Missing or incorrect database file selection.
-   Incorrect row numbers.
-   Excel-related issues like copying worksheets.

## Future Enhancements

-   Add functionality to process additional student records automatically without manual row selection.
-   Implement data validation checks before processing.
-   Improve error logging to a file for debugging purposes.

## Contribution

If you want to contribute to this project, feel free to fork the repository, create new branches, and submit pull requests with your improvements.

## License

This project is licensed under the MIT License.