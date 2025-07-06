# <p align="center">Excel and Twitter Data Processor</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

## Introduction

This project provides a Python application that combines the power of Excel data processing with real-time Twitter (X) data scraping. It exposes a FastAPI endpoint for interacting with Excel data, answering user queries using AI, and even modifying the spreadsheet. Additionally, it includes functionality for scraping and processing Twitter data to identify and analyze emergency-related information. This is designed for developers and data scientists who need to integrate Excel with web data and build AI-powered data processing solutions.

## Table of Contents

1.  [Key Features](#key-features)
2.  [Installation Guide](#installation-guide)
3.  [Usage](#usage)
4.  [API Reference](#api-reference)
5.  [Environment Variables](#environment-variables)
6.  [Project Structure](#project-structure)
7.  [Technologies Used](#technologies-used)
8.  [License](#license)

## Key Features

*   **Excel Interaction via API**:  Allows querying and modifying Excel spreadsheets using a FastAPI.
*   **AI-Powered Data Processing**: Employs AI agents to process user queries against Excel data and execute commands.
*   **Twitter (X) Data Scraping**: Scrapes tweets based on a search query using Selenium.
*   **Emergency Tweet Identification**: Analyzes scraped tweets to identify emergency-related information.
*   **Data Analysis with Pandas**: Uses Pandas for efficient data manipulation and analysis of both Excel and Twitter data.
*   **Health Check Endpoint:** Provides an endpoint to check API, Excel connection, and external dependencies.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

3.  **Set up environment variables:**
    Create a `.env` file in the root directory and add the following variables:

    ```
    EXCEL_FILE_PATH=/path/to/your/excel/file.xlsx
    GEMINI_API_KEY=your_gemini_api_key
    ```
   **Note:**: Replace the placeholder with actual values

4.  **Run the FastAPI server:**

    ```bash
    python main.py
    python app.py
    ```

## Usage

The application provides a FastAPI that can be used to interact with data.

*   **/process\_query**: Accepts a user query and returns an AI response based on Excel data.
*   **/autonomous\_action**: Processes a user query and executes autonomous actions in Excel.
*   **/execute\_command**: Executes a specific command in Excel.

The Twitter scraping functionality is integrated into `main.py`.  You can adjust the `SEARCH_QUERY` variable to customize the scraping.

## API Reference

*   **/process\_query**
    *   **Method**: POST
    *   **Request Body**:
        ```json
        {
            "query": "Your question about the Excel data"
        }
        ```
    *   **Response**:
        ```json
        {
            "response": "AI-generated answer based on Excel data"
        }
        ```
    *   **Error**: Returns a 500 status code with an error message if there is an exception during processing.

*   **/autonomous\_action**
    *   **Method**: POST
    *   **Request Body**:
        ```json
        {
            "query": "Instruction to modify Excel",
            "context": "Optional context for the query",
            "image_data": "Optional image data"
        }
        ```
    *   **Response**:
        ```json
        {
            "result": "Result of the autonomous action"
        }
        ```
    *   **Error**: Returns a 500 status code with an error message if there is an exception during processing.

*   **/execute\_command**
    *   **Method**: POST
    *   **Request Body**:
        ```json
        {
            "command": "Excel command to execute"
        }
        ```
    *   **Response**:
        ```json
        {
            "result": "Result of the command execution"
        }
        ```
    *   **Error**: Returns a 500 status code with an error message if there is an exception during processing.

*   **/health\_check**
    *   **Method**: GET
    *   **Response**:
        ```json
        {
            "status": "ok",
            "excel_connection": "connected",
            "gemini_api": "configured"
        }
        ```

*   **/connect\_to\_excel**
    *   **Method**: POST
    *   **Response**:
        ```json
        {
            "connection_status": "success"
        }
        ```
    *   **Error**: Returns a 500 status code with an error message if connection fails.

## Environment Variables

*   `EXCEL_FILE_PATH`: The path to the Excel file to be processed.
*   `GEMINI_API_KEY`: The API key for the Gemini AI model.

## Project Structure

```
.
├── app.py          # FastAPI application
├── main.py         # Web scraping and data processing logic
└── readme.md       # Project documentation
```

## Technologies Used

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

*   **Backend**: FastAPI (Python)
*   **Web Scraping**: Selenium
*   **Data Analysis**: Pandas
*   **AI Model**: Gemini (configured via API key)

## License

MIT License