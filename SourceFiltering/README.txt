This script filters out certain sources that are related to what the researcher is looking for based on the prevalence of keywords in the file. 


Here are the file names and what they do:

FilteringScript.py is the main script that should be run in order to produce results. 
Keywords.txt is where keywords that the researcher wants to appear in the article are located. When showing results, the minimum number of keywords can be set so that results will only show up if they are above the minimum.
Whitelist.txt is where whitelisted words will be pulled from. These are words that may produce false positives and create unnecessary sources that do not have anything to do with the topic. 
Sources.txt is where the sources are pulled from, and can be changed in the main script to be better renamed. 
Sources.txt.result.txt is the result from the main script, and the name will be changed based on the sources file name. 


This script should be run in python 3. 

Note: For the excel version of the script, it will automatically write the results to excel files. However, libraries need to be installed before running.


In the cmd of windows, these commands have to be run in order:


python -m pip install -U pip
pip install xlsxwriter 
