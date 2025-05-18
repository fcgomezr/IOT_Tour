# Process.py - Employment and Occupation Multipliers Calculator

## Overview

This module calculates employment and occupation multipliers for the tourism sector using Input-Output Tables (IOT) analysis. It implements the Leontief inverse matrix method to determine how changes in tourism demand affect employment across different sectors of the economy.

## Functionality

Process.py performs the following key operations:

1. **Data Loading**: 
   - Imports employment, occupation, and production data from the `extract_data` module
   - Loads Input-Output Table data from CSV files in the ProcessedData directory

2. **Matrix Operations**:
   - Converts data to TensorFlow tensors for efficient matrix calculations
   - Creates diagonal matrices for employment and occupation ratios
   - Applies the Leontief inverse method to calculate multipliers
   - Calculates final multiplier matrices for employment and occupation

3. **Output Generation**:
   - Exports results to an Excel file with multiple sheets for different analyses
   - Formats data appropriately for economic analysis and reporting

## Data Flow

```
extract_data.py               ProcessedData/2019/Leon.csv
     │                                  │
     ▼                                  ▼
  Raw Data                       Input-Output Table
     │                                  │
     └──────────────► Process.py ◄──────┘
                           │
                           ▼
             Multiplicadores de Empleo Turismo.xlsx
```

## Economic Methodology

The code implements the Leontief input-output model, a fundamental tool in economic analysis:

### Leontief Inverse Method

The Leontief model is based on the equation:
```
x = (I - A)^(-1) * y
```
Where:
- x: total output vector
- I: identity matrix
- A: technical coefficients matrix
- y: final demand vector
- (I - A)^(-1): Leontief inverse matrix

In our implementation:
- The Leon matrix (loaded from Leon.csv) represents the technical coefficients
- The `nl = tf.linalg.inv(meloci)` operation calculates the Leontief inverse
- Employment and occupation matrices are derived by multiplying diagonal matrices of employment and occupation ratios with the Leontief inverse

### Employment and Occupation Multipliers

The multipliers show how changes in final demand (such as tourism spending) affect:
1. **Direct effects**: Jobs created in the sector receiving the initial spending
2. **Indirect effects**: Jobs created in sectors supplying goods/services to the tourism sector
3. **Induced effects**: Jobs created due to increased household spending from new income

## Code Explanation

### Data Loading and Preparation
```python
# Loading employment, occupation, and production data
emp = ed.empleos_df.select(ed.empleos_df.columns[-1]).to_series() 
oci = ed.ocupados_df.select(ed.ocupados_df.columns[-1]).to_series() 
prod = ed.production_df.select(ed.production_df.columns[-1]).to_series()

# Loading IOT data
df_leon = pl.read_csv(ruta_leon, separator=';', has_header=True, decimal_comma=True)
primera_columna = df_leon.columns[0]
matriz_leon = df_leon.drop(primera_columna).to_numpy()
```

### Matrix Calculations
```python
# Converting to TensorFlow tensors
g_tensor = tf.convert_to_tensor(prod, dtype=tf.float32)
leon_tensor = tf.convert_to_tensor(matriz_leon, dtype=tf.float32)
emp_tensor = tf.convert_to_tensor(emp, dtype=tf.float32)
oci_tensor = tf.convert_to_tensor(oci, dtype=tf.float32)

# Creating employment and occupation matrices
mod = tf.divide(emp_tensor, g_tensor)
mod2 = tf.divide(oci_tensor, g_tensor)
diag_oci = tf.linalg.diag(mod2)
diag_empl = tf.linalg.diag(mod) 

# Matrix multiplication with Leon tensor
mele = tf.matmul(diag_empl, leon_tensor)   # Leon matrix EMPE
meloci = tf.matmul(diag_oci, leon_tensor)  # Leon matrix OCUP   

# Leontief inverse calculation
nl = tf.linalg.inv(meloci)

# Final multiplier calculation
emp_col = tf.expand_dims(oci_tensor, axis=-1)
F = tf.matmul(nl, emp_col)
```

### Output Generation
The calculations produce three key outputs that are exported to an Excel file:

1. **F tensor (Demanda Final)**: Final demand impact on employment
2. **meloci (Mult Ocup 2024)**: Occupation multipliers by sector
3. **mele (Mult Empl 2024)**: Employment multipliers by sector

## Usage

To run the employment multiplier calculations:

1. Ensure that the `extract_data.py` module has been executed to prepare the input data
2. Verify that the Input-Output Table CSV files are present in the ProcessedData directory
3. Run the script using Python: `python Process.py`
4. The results will be written to "Multiplicadores de Empleo Turismo - 2.1.xlsx"

## Requirements

- TensorFlow
- Polars
- NumPy
- OpenPyXL
