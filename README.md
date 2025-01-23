# Prometheus Query Generator

A Python tool for generating SQL queries from Excel data and exporting database state.

## Features

- Generate SQL queries from Excel input files
- Export current database state to Excel
- Support for multiple environments (SIT, UAT, PREPROD, PROD)
- Formatted Excel output with proper styling

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/prometheus-query-generator.git
cd prometheus-query-generator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the program:
```bash
python main.py
```

The program will:
1. Ask for an input Excel file
2. Generate SQL queries based on the Excel data
3. Save the queries to a file
4. Optionally export the current database state to Excel

## Project Structure

- `src/`
  - `excel/`: Excel reading and export functionality
  - `query/`: Query generation logic
  - `utils/`: Common utility functions
- `config/`: Configuration files and constants
- `main.py`: Main application entry point
- `requirements.txt`: Python dependencies

## Building

The project includes GitHub Actions workflows to automatically build an executable file.
You can find the latest build in the GitHub Actions artifacts.

## License

[MIT License](LICENSE)
