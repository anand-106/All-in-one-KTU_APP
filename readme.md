# <p align="center">Web Scraping and Excel Automation API</p>

<p align="center">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

This project provides a Flask API for web scraping Twitter (X) and automating Excel tasks. It scrapes tweets based on a search query, processes the data, and interacts with Excel using an `excel_agent`. The API allows users to process queries, execute commands, and connect to an Excel instance.

## Table of Contents

1. [Key Features](#key-features)
2. [Installation Guide](#installation-guide)
3. [Usage](#usage)
4. [Environment Variables](#environment-variables)
5. [Project Structure](#project-structure)
6. [Technologies Used](#technologies-used)
7. [License](#license)

## Key Features

- **Web Scraping:** Scrapes tweets from Twitter (X) based on a search query using Selenium.
- **Data Processing:** Processes scraped data using Pandas, filtering for emergency-related tweets and location mentions.
- **Excel Automation:** Interacts with Excel through an `excel_agent` to process queries and execute commands.
- **API Endpoints:** Provides Flask API endpoints for processing queries, executing commands, connecting to Excel, autonomous actions in Excel, and health checks.
- **Error Handling:** Implements robust error handling for web scraping, API calls, and Excel interactions.
- **Configurable Chrome Driver:** Uses a Chrome profile to persist session data, avoiding bot detection.

## Installation Guide

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Install dependencies:**

    ```bash
    pip install Flask selenium pandas
    # Install ChromeDriverManager (if not already installed and needed)
    # pip install webdriver_manager
    ```

3.  **Set up environment variables:**

    Create a `.env` file in the project root and add the necessary environment variables (see [Environment Variables](#environment-variables) section).

4.  **Run the Flask server:**

    ```bash
    python app.py
    ```

## Usage

The Flask API provides the following endpoints:

-   **`/` (GET):** Renders the main page (`index.html`).
-   **`/process_query` (POST):** Processes a user query using the `excel_agent`.
-   **`/autonomous_action` (POST):** Executes an autonomous action on excel.
-   **`/execute_command` (POST):** Executes a command using the `excel_agent`.
-   **`/connect_to_excel` (GET):** Connects to an Excel instance using the `excel_agent`.
-   **`/health_check` (GET):** Checks the health of the API and the Excel connection.

Example `process_query` request (using `curl`):

```bash
curl -X POST -H "Content-Type: application/json" -d '{"query": "Your Excel Query"}' http://localhost:5000/process_query
```

## Environment Variables

The following environment variables are required:

-   `EXCEL_FILE_PATH`: The path to the Excel file used by the `excel_agent`.  (Example)

## Project Structure

```
.
├── app.py       # Flask API implementation
├── main.py      # Web scraping and data processing functions
└── readme.md    # This README file
```

## Technologies Used

<p align="left">
  <a href="#"><img src="https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi" alt="FastAPI"></a>
  <a href="#"><img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"></a>
  <a href="#"><img src="https://img.shields.io/badge/Selenium-4DB33D?style=for-the-badge&logo=selenium" alt="Selenium"></a>
  <a href="#"><img src="https://img.shields.io/badge/pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas"></a>
</p>

-   **Backend:** Flask (Python)
-   **Web Scraping:** Selenium
-   **Data Processing:** Pandas
-   **Other:** ChromeDriverManager (potentially)

## License

MIT License