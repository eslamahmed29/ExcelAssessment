Excel Assessment Project
Overview
This project is a web-based application developed in C# using ASP.NET Core. The application allows users to upload an Excel sheet, perform calculations, and download the modified version of the sheet.

Features:
Excel Import: Allows users to upload an Excel file containing tax data.
Data Processing: Adds a new column called "Total Value before Taxing" and calculates values based on the provided tax information.
Automatic Calculations: Adds a summary row at the end, calculating the sum of "Total Value After Taxing" and "Total Value before Taxing".
Download Processed Excel: Provides an option to download the modified Excel sheet with the new calculations.
Technologies Used
.NET 6.0 / .NET Core 6.0: For building the web application.
ASP.NET Core MVC: Used for the web interface.
EPPlus: A library for reading from and writing to Excel files.
Bootstrap: For styling the web interface.
Getting Started
Prerequisites
.NET SDK 6.0: Download .NET SDK
Visual Studio 2022 or a compatible IDE with .NET 6.0 support.
Git: Download Git
Installation
Clone the Repository:

bash
Copy code
git clone https://github.com/yourusername/ExcelAssessmentProject.git
cd ExcelAssessmentProject
Open the Project:

Open the solution file (.sln) in Visual Studio.
Install Dependencies:

Run the following command in the terminal to install any missing dependencies:
bash
Copy code
dotnet restore
Running the Application
Run the Project:

In Visual Studio, set the project as the startup project.
Press F5 or use the command:
bash
Copy code
dotnet run
Access the Application:

Open your browser and navigate to https://localhost:5001 (or the specified port).
Upload an Excel File:

Use the interface to upload your "Taxes Sheet.xlsx" file.
The application will process the sheet and display the calculated totals.
Download the Modified Excel:

After processing, click the download button to get the modified Excel file.
Project Structure
Controllers: Contains the HomeController for managing file uploads and processing logic.
Models: Holds data models (if any are used).
Views: Contains Index and Result views for displaying the upload form and results.
wwwroot: Holds static files (e.g., uploads folder for storing Excel files temporarily).
Notes
Make sure to upload an Excel file in .xlsx format with the required columns ("Total Value After Taxing" and "Taxing Value").
The application checks if the uploaded Excel file has already been processed to avoid double calculations.
Example files can be found in the sample_data directory.
Known Issues
Ensure the column names in the Excel file match exactly as expected ("Total Value After Taxing", "Taxing Value").
The application assumes that the first sheet of the uploaded Excel file contains the data.
