# Utility MEP-MRP GUI

This script/executable creates a UI using tkinter. The UI allows users to create an excel file in the current working directory with their choice of **Gas** **Water** or **Electric** Meter Exchange/Meter Replacement outputs.

## Data

Single SQL folder containing all files needed to run the script

1. SQL/

   - Electric MEP MRP Queries.sql: Electric meter replacement and exchange
   - GAS MEP-MRP QUERIES.sql: Gas meter replacement and exchange
   - WATER MEP-MRP QUERIES.sql: Water meter replacement and exchange

## Data Contact

 **Razan Hussien - Initial Work/Author - <razan.hussien@kub.org>**
 **Bryce Cook - Contributor - <bryce.cook@kub.org>**

## Getting Started

### Steps Below

1. Clone this repository locally or download the .zip file and extract it in your project directory
2. Ensure you're using Python version 3.11 or lower (cx_Oracle dependency)
3. Create your virtual environment
4. Run the UI.py file
5. In order to build the .exe ensure you have pyinstaller installed

```shell
python -m venv venv
venv\Scripts\activate
python -m pip install -r requirements.txt
```

```shell
python UI.py 
```

### Building the .exe

```shell
pyinstaller UI.spec
```

- The line above creates build and dist folders, the executable should be visible in the dist/ folder

### Requirements

The requirements.txt file contains all necessary dependencies for running the script

## Main Packages

### Standard Library Imports

- **sys**: Provides access to some variables used or maintained by the Python interpreter.
- **os**: Provides operating system dependent functionality, such as file and directory manipulation.
- **tkinter**: GUI toolkit for Python applications.
- **pandas**: Data manipulation and analysis library.
  - DataFrame: Data structure for tabular data representation and operations.
  - Series: One-dimensional labeled array capable of holding data of various types.
- **sqlalchemy**: SQL toolkit and Object-Relational Mapping (ORM) library for Python.
  - create_engine: Function to create a SQLAlchemy engine for database connections.
- **dotenv**: Module for parsing .env files to load environment variables.
  - load_dotenv: Function to load environment variables from a .env file into the environment.

### Third-Party Library Imports

- **cx_Oracle**: Enables connection to Oracle databases, allowing execution of SQL and PL/SQL statements.
- **matplotlib**: Data visualization library for creating static, animated, and interactive visualizations in Python.
  - pyplot: Provides a MATLAB-like interface for creating plots and visualizations.
  - bar: Function to create bar charts in matplotlib.
  - annotate: Function for adding annotations to plots.
- **logging**: Module for flexible logging in Python applications.
  - basicConfig: Function to set up logging configuration.
  - DEBUG: Log level for detailed diagnostic information.
  - ERROR: Log level indicating errors that should be investigated.

## Contributing

Please contact **Razan Hussien** or **Bryce Cook** before attempting to make changes to this repository

## Authors

- **Razan Hussien** - [razanhussien](https://github.kub.org/RNN08578); <razan.hussien@kub.org>
