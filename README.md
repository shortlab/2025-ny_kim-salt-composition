# 2025-ny_kim-salt-composition
Public data repository for Nayoung's paper on molten fluoride salt ICP-MS analysis protocols

**Raw data files** 
Current signal report shows the counts of elements and the instrument parameters. Detailed calibration summary contains the calibration curves of all elements. Quick batch report is a PDF version of results. Batch table is an excel version of results. Its inputs are extracted and used to calculate the final concentrations. 

**MATLAB script for data processing**
1st input file is the excel file with sampling and dilution information. The file name starts with “ICP #”. The format should be kept the same in order to apply the same code for data processing. 
2nd input file is the batch table with all measurements. The number of samples in the 1st input file must match the number of samples in the 2nd input file. When multiple samples from different investigations of parameters are analyzed within the same batch, the author has manually separated the two investigation raw data files. This may cause differences in the quick batch report PDF file and the batch table excel file in some cases. The name must contain either “Major” for Li, Na, and K or otherwise “Trace”. 
The code executes extraction or calculations for following values. 

	Dilution data with salt mass, solution volume, internal standard concentration, dilution factors
	Raw ICP data with concentration and RSD of each element for each sample 
	Concentration of replicates, average concentration, sample standard deviation, and propagated error values in mol. % 
	Concentration of replicates, average concentration, sample standard deviation, and propagated error values in wt. ppm. 
However, the current script automatically generates excel files for average and standard deviations regarding wt. ppm only. If desired, one can simply uncomment the line corresponding to the generation of excel files for other values as well. Otherwise, all result values are stored in the MATLAB space. 

**Processed data files** 
“ICP ##_CONC wt%_AVG_Major/Trace” consists of concentrations averaged from replicate measurements. “ICP ##_CONC wt%_STD_Major/Trace” consists of sample standard deviation values calculated from replicate measurements. “ICP ##_sigma_C_wt_avg_Major/Trace” is the errors propagated from the measurement of each variable in calculating the average wt. ppm values. 
MATLAB script for graphing data (scatter and bar plots) 
1st input file is the excel file containing average concentrations with a name “ICP ##_CONC wt%_AVG_Major/Trace”. The user can choose which elements to display in the graph. 2nd and 3rd input files required are sample standard deviation and propagated error files, respectively. Graph settings can be customized as desired. 
