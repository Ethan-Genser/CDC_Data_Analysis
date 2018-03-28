# CDC_Data_Analysis

Visualizes datasets provided by the Center for Disease Control and Prevention with xlsx spreadsheets.

![alt text](https://github.com/GHC-0/CDC_Data_Analysis/blob/master/Info/Screenshot.png)

*Note: Any use of this data implies consent to abide by the Center for Disease Control and Prevention's term of the data use restrictions as stated [here](https://wonder.cdc.gov/ucd-icd10.html).*

### Prerequisites

Before you can run this program, ensure that yoou have the following software installed and functional:
* Python 3.5.2 or better
* xlsxwriter 1.0.2 or better

### Installing

To install this program, simply download the file called 'CDC_Data_Analysis.py' as well as some datasets. Datasets for the years 2008-2016 are provided in the folder named 'Datasets'. If you wish to use a dataset from a year that is not included in this repository, visit the CDC's official website [here](https://wonder.cdc.gov/ucd-icd10.html).

If you choose to use extra datasets not provided by the repository, be sure to change the `YEARS` constant in the `CDC_Data_Analysis.py` source file to include your chosen years.

## Running

To run this program, move all the datasets you wish to analyze to the same folder as `CDC_Data_Analysis.py`. Run the program from your python command terminal. The program will create an xlsx file on your desktop contain the collected data.

## Authors

* [**Ethan Genser**](https://github.com/Ethan-Genser) - *Creator*

## License

This project is licensed under the Apache License Version 2.0 - see the [LICENSE](Info/LICENSE) file for details.
