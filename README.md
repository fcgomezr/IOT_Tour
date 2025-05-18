# IOT_Tour

## Overview
IOT_Tour is an Input-Output Tables analysis tool for the Tourism sector. It calculates employment and occupation multipliers to quantify how tourism activities affect employment across different sectors of the economy.

## Features
- Process raw production and employment data from various sources
- Transform and correlate data between different economic classification systems
- Calculate employment and occupancy multipliers using the Leontief inverse method
- Generate standardized reports in Excel format for further analysis

## Installation

### Prerequisites
- Python 3.10 or 3.11
- Poetry package manager

### Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/fcgomezr/IOT_Tour.git
   cd IOT_Tour
   ```

2. Install dependencies using Poetry:
   ```bash
   poetry install
   ```

## Usage Guide

### Data Preparation
1. Ensure your raw data files are placed in the `RawData` directory:
   - Production matrix data (`Matriz_Prod_2023.xlsx`)
   - Employment and occupation data (`Ocupados y TETC para multiplicadores 2021-2024.xlsx`)

2. Verify that correlation files are in the `ProcessedData` directory:
   - Mapping between classification systems (`Correlativa 68 a 27.xlsx`)
   - Input-Output Tables in CSV format (`ProcessedData/2019/Leon.csv`)

### Running the Model
The model consists of two main scripts that should be run in sequence:

1. **Data Extraction** - Prepares the input data:
   ```bash
   poetry run python extract_data.py
   ```
   This script:
   - Extracts production values from the raw data
   - Correlates different economic classification systems
   - Prepares employment and occupation data

2. **Multiplier Calculation** - Performs the matrix operations:
   ```bash
   poetry run python Process.py
   ```
   This script:
   - Loads the prepared data
   - Applies the Leontief inverse method
   - Calculates employment and occupation multipliers
   - Exports results to Excel

### Output Files
The main output is an Excel file named `Multiplicadores de Empleo Turismo - 2.1.xlsx` containing:

- **Demanda Final**: Final demand impact on employment
- **Mult Ocup 2024**: Occupation multipliers by sector
- **Mult Empl 2024**: Employment multipliers by sector

### Interpreting Results
The multipliers in the output Excel file show:

1. **Direct Effects**: Jobs created in the sector receiving the initial tourism spending
2. **Indirect Effects**: Jobs created in sectors supplying goods/services to the tourism sector
3. **Induced Effects**: Jobs created due to increased household spending from new income

To interpret the results:
- Higher multiplier values indicate sectors with greater employment impact from tourism spending
- Compare multipliers across sectors to identify priority areas for tourism development
- Use the multipliers to estimate employment changes from projected tourism growth

## Advanced Usage

### Updating Input Data
To update the model with new data:

1. Replace the input files in the `RawData` directory with updated versions
2. Run the extraction and calculation scripts as described above
3. Compare results with previous outputs to analyze trends

### Customizing Analysis
For custom analyses:

1. Modify the extraction parameters in `extract_data.py` to focus on specific sectors
2. Adjust matrix operations in `Process.py` for different economic scenarios
3. Export results to custom Excel templates by modifying the output section of `Process.py`

## Documentation
For more detailed information about the code:
- See [Process_README.md](Process_README.md) for details on the matrix calculations and economic methodology

## Requirements
- TensorFlow
- Polars
- NumPy
- OpenPyXL
- FastExcel 
