# ID_Jarvis_Bango
This Repository is build for the Data Blending for 64 Models divided into GT and MT channels at various granularity which includes at National level , Channel level , Account/Region level and Basepack level . The Repository contains the .r scripts for the data preprocessing which will be used for the Predictive modelling and the data visualization .

# R packages used 
The code uses heavily the dplyr , tidyr package created by Hadley Wickham and also it uses hblar,stringr ,reshape,magrittr,readxl and writexl,svdialog

# Input Data Sources 
There are 9 Data sources namely - primary sales , Secondary Sales , Secondary Stock , Dispactch Rate (DR) ,I trust , penetration , TTS and BMI, Sellout. Out of 9 data source - 5 data source are only used in MT models which are primary sales , DR,TTS,BMI.Sellout. For GT models 8 data sources are used excluding Sell out.

# Script Synopsis ( Generalised , GT Model and MT Model)
The entire project has been divided into 3 scripts namely - 1. Generalised template 2.GT Model 3.MT Model 
1. Generalised template - The generalised template is used to do the prework on the input files which we receive from the system. It has preprocessing functions which is used to create the tidy data sets from the messy dataset .
2.GT Model - The GT Model is used for running the 31 GT models at various granularity which includes - National , Channel , Packgroup ,Region.
3.MT Model - The MT Model is used for running the 33 MT models at various granularity - National , Account , Basepack.

# How to use the Model 
The Generalised template is called in both the GT and MT model which reduces duplication of the code. 

# Project Benchmark ( Run time execution) 
1. GT Model : 
2. MT Model :




