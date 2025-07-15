# <p align="center">Excel Data Automation and Twitter Scraper API</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
</p>

## Introduction

This project provides an API to automate Excel tasks and scrape data from Twitter/X. It allows users to process queries, execute commands within Excel, and gather relevant information from social media for analysis and automation purposes. Target users include data analysts, automation engineers, and anyone needing to integrate Excel with web data.

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

*   **Excel Automation:** Execute commands and process queries directly within Excel via API calls.
*   **Twitter/X Scraping:** Scrape tweets based on specific search queries using Selenium.
*   **Data Processing:** Filter and process scraped data to identify relevant information.
*   **Web API:** Expose functionalities through a Flask-based web API.
*   **Health Check:** Verify the API and Excel connection status.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install dependencies:**

    ```bash
    pip install flask selenium pandas webdriver_manager
    ```

3.  **Set up environment variables:**
    Create a `.env` file in the root directory and define the following variables:

    ```
    EXCEL_AGENT_CONFIG_PATH=<path_to_excel_agent_config>  # Example: ./config/excel_agent_config.json
    SEARCH_QUERY="your_search_query"                     # Example: "emergency situation"
    CHROME_PROFILE_PATH="/path/to/chrome/profile"          # Example: /Users/username/chrome_profile
    CHROME_PROFILE_NAME="Profile 1"                      # Example: Default
    GEMINI_API_KEY="your_gemini_api_key"                   # Example: AIzaSy...
    ```

    **Note:**
    *   Replace the placeholder values with your actual configurations.
    *   `EXCEL_AGENT_CONFIG_PATH` refers to the configuration file for the `excel_agent` module which contains details such as workbook name and path.
    *   `SEARCH_QUERY` is the query to use for scraping Twitter/X.
    *   `CHROME_PROFILE_PATH` and `CHROME_PROFILE_NAME` point to the Chrome profile you want Selenium to use.
    *   `GEMINI_API_KEY` is the API key for the Gemini AI model.

4.  **Run the Flask server:**

    ```bash
    python app.py
    ```

## Usage

The application provides several API endpoints for interacting with Excel and scraping data.

*   **/process\_query**: Processes a user query and returns an AI-generated response.
*   **/autonomous\_action**: Processes a query and executes actions in Excel.
*   **/execute\_command**: Executes a specific command in Excel.
*   **/health\_check**: Checks the health of the API and Excel connection.
*   **/connect\_to\_excel**: Establishes a connection to Excel.

Example:
To process a query, send a POST request to `/process_query` with the following JSON payload:

```json
{
  "query": "What is the total sales for product X?",
  "context": "The sales data is in sheet 'Sales Data'",
  "image_data": "base64_encoded_image"
}
```

## API Reference

### `/process_query`

*   **Method:** POST
*   **Request Body:**
    ```json
    {
      "query": "string",
      "context": "string",
      "image_data": "string" // Base64 encoded image
    }
    ```
*   **Response:**
    ```json
    {
      "response": "string" // AI generated response
    }
    ```

### `/autonomous_action`

*   **Method:** POST
*   **Request Body:**
    ```json
    {
      "query": "string",
      "context": "string",
      "image_data": "string" // Base64 encoded image
    }
    ```
*   **Response:**
    ```json
    {
      "result": "string", // Result of the action
      "error": "string"   // Error message if any
    }
    ```

### `/execute_command`

*   **Method:** POST
*   **Request Body:**
    ```json
    {
      "command": "string" // Command to execute in Excel
    }
    ```
*   **Response:**
    ```json
    {
      "result": "string", // Result of the command
      "error": "string"   // Error message if any
    }
    ```

### `/health_check`

*   **Method:** GET
*   **Response:**
    ```json
    {
      "api_status": "string", // "running" or "down"
      "excel_connection": "boolean",
      "gemini_api_key_configured": "boolean"
    }
    ```

### `/connect_to_excel`

*   **Method:** GET
*   **Response:**
    ```json
    {
      "status": "string" // "connected" or "failed"
    }
    ```

## Environment Variables

| Variable                      | Description                                                                     |
| ----------------------------- | ------------------------------------------------------------------------------- |
| `EXCEL_AGENT_CONFIG_PATH`    | Path to the Excel agent configuration file.                                  |
| `SEARCH_QUERY`               | The search query used for scraping Twitter/X.                                  |
| `CHROME_PROFILE_PATH`          | Path to the Chrome profile directory.                                         |
| `CHROME_PROFILE_NAME`          | Name of the Chrome profile to use.                                            |
| `GEMINI_API_KEY`               | API key for the Gemini AI model.                                               |

## Project Structure

```
.
├── app.py      # Flask web application logic and API endpoints
├── main.py     # Web scraping and data processing functions
└── readme.md   # Project documentation
```

## Technologies Used

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas" alt="Pandas"></a>
  <a href="#"><img src="https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask"></a>

</p>

*   **Backend:** Flask, Python
*   **Web Scraping:** Selenium
*   **Data Processing:** Pandas
*   **WebDriver Management:** `webdriver_manager`

## License

MIT License