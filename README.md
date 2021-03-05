# Awesome_Gear_Pro_Report

Gleason Gear Inspection System can report the Gear Profile Slop Dev. <img src="https://render.githubusercontent.com/render/math?math=fH\alpha m"> according to different height levels. This allows GM to assign different tolerance zone to the coast side of the grinding gears due to the grinding gear's nature bias.

For example, if a helical gear's profile was measured on 3 teeth (#1, 11, 21),  and 3 height levels (top, middle, bottom). Gleason can report can report:


<img src="https://render.githubusercontent.com/render/math?math=fH\alpha m(Top)=\frac{fH\alpha1(Top) %2B fH\alpha11(Top)%2BfH\alpha21(Top)}{3}">


<img src="https://render.githubusercontent.com/render/math?math=fH\alpha m(Mid)=\frac{fH\alpha1(Mid)%2BfH\alpha11(Mid)%2BfH\alpha21(Mid)}{3}">


<img src="https://render.githubusercontent.com/render/math?math=fH\alpha m(Bot)=\frac{fH\alpha1(Bot)%2BfH\alpha11(Bot)%2BfH\alpha21(Bot)}{3}">


While, Zeiss Gear Pro can only report the grand mean over all twist checks:

<img src="https://render.githubusercontent.com/render/math?math=fH\alpha m =\frac{fH\alpha1(Top)%2BfH\alpha11(Top)%2BfH\alpha21(Top)%2BfH\alpha1(Mid)%2BfH\alpha11(Mid)%2BfH\alpha21(Mid)%2BfH\alpha1(Bot)%2BfH\alpha11(Bot)%2BfH\alpha21(Bot)}{9}">





With this Awesome_Gear_Pro_Report.py, you can define a CSV file as template, and printout a fully customized gear result based on the template.

Here is how it works:

1) put the **Awesome_Gear_Pro_Report.exe** and the  **Awesome_Gear_Pro_Report_Config.txt** files into the same folder, under the Calypso Program, modify (or add if you do not have one yet) the **inspection_end.bat ** file so that Awesome_Gear_Pro_Report.exe will be automatically executed after inspection finished .

2) Inside of the  **Awesome_Gear_Pro_Report_Config.txt**, you put in 3 information:

- Where to find the Zeiss _chr.txt report

- What's the name of the Zeiss Gear Pro program with * as the wildcard for regular expression

- Where to find the customized template CSV file

3) Inside the customized template CSV file, you will define from which lines of the Zeiss_chr.txt report, a particular dimension should be calculated. 4 Functions are supported: "AVERAGE, EQUAL, MIN, and MAX".

You will put the dimension numbers and brief descriptions in column A and column B, you will put a USL and a LSL to control each of the dimensions in column C and column D, in column E, you put a Function name, and in column F, you specify from which lines (0 index) of the Zeiss_chr.txt report, a particular dimension should be calculated.

For example, if you put AVERAGE in column E, and put 0-3-6-9 in column F, this is going to telling **Awesome_Gear_Pro_Report.exe** that you want to calculate the average of the actual reading from Zeiss_chr.txt's line 0, 3, 6, 9.

4) The **Awesome_Gear_Pro_Report.exe** will always find the last report that match your regular expression in the config file, generate a result Excel file, highlight the dimensions that are out of tolerance based on your template CSV file, and call the system printer to printout the customized report.  
