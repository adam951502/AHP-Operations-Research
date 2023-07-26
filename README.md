# **Analytical Hierarchy Process for Operations Research**

This project is an implementation of the Analytical Hierarchy Process (AHP), a framework for decision making that involves structuring a hierarchy from a given set of data. This project was developed by Yu-Sheng Tang at the University of Freiburg during the Winter Semester of 2019-20 for the Sustainable Systems Engineering program.

The project uses the AHP method to analyze operations research data from a provided Excel file. The data represents various factors to consider, including Levelized Cost of Electricity (LCOE), ability to respond to demand, efficiency, capacity factor, land use, environmental external costs, human health external costs, job creation, social acceptability, and external supply risk.

## **Prerequisites**

- Python
- Numpy
- xlrd

## **Usage**

1. Make sure you have the required dependencies installed in your Python environment.
2. Run the Python script **`Tang_Roland_AHP_Script.py`**. The script will load the data from the **`Tang_Roland_AHP_InputMatrices.xlsx`** Excel file.
3. The script computes the weight of each factor using the AHP method, checks the consistency ratio to ensure it is below 0.1, and then calculates the final score for each factor.
4. The final scores are printed to the console. These scores can be used to make decisions based on the given factors.

Please note that the script assumes that the Excel file is in the same directory as the Python script and that the file contains properly formatted data. If the consistency ratio is not below 0.1, you will need to adjust the data in the Excel file.