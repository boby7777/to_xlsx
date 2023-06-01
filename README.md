# XLS to XLSX Converter

This script converts XLS files to XLSX format using the Excel COM object.

## Prerequisites

- Python 3.x
- `win32com` module (`pywin32` package)

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/boby7777/to_xlsx.git
   ```

2. Install the required dependencies:

   ```bash
   pip install pywin32
   ```

## Usage

```bash
python to_xlsx.py input.xls output.xlsx
```

Replace `input.xls` with the path to your input XLS file and `output.xlsx` with the desired output XLSX file path.

If the output file already exists, the script will prompt you to confirm whether to replace it or abort the conversion.

## Contributing

Contributions are welcome! If you encounter any issues or have suggestions for improvements, please [open an issue](https://github.com/your-username/your-repo/issues) or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
