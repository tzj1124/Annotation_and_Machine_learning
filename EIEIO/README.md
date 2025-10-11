# User Guide 
1. Create a folder named "input-mzml" in the downloaded package before running the _identification.py_ script. This tool supports the mzML file format of the LC-MS raw data and only need to run the _identification.py_ to get the final output file after downloading this folder. For the traceability of the identification results, the temporary results obtained in each step were saved. 
2. It is recommended to convert the original data into mzml format in the MSConvert software by following the steps shown in the picture below.
![msconvert](https://github.com/user-attachments/assets/da3dfe84-0ca6-444d-b69f-3d2017f596be)
3. It is worth mentioning that the 0.01 Da tolerance in MS1 is the default value in our tool. This value can be modified in the step3.py file according to individual needs.
