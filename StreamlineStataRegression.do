*************************************************************************
					*Streamlining regression models in STATA*
					
*Looping example of regression with multiple exposures/outcomes, several specific adjustment models, several modes (categorical and continuous exposures)
*Excel output example of regression model as part of loop incl. LR test
*Excel functions for clean and adaptable output
				
************************************************************************
*Author: Anika Knuppel; Date: 28/12/2021; contact: https://github.com/aknuppel or @anika_knuppel




*Load test dataset
use https://www.stata-press.com/data/r17/nhanes2
cd "/"

*Outcomes: highbp diabetes
*Exposures: tcresult tgresult iron albumin vitaminc zinc copper (continuous), quintilestrends and quintiles
*Three adjustment models that differ* by outcome:
	*for highbp
		*Mod1:sex race age; Mod2:Mod1+hsizgp + region; Mod3: Mod2+bmicat+ diabetes*
	*for diabetes
		*Mod1:sex race age; Mod2:Mod1+hsizgp + region; Mod3: Mod2+bmicat+ highbp*

*Dataprep
label def quintL 1"Q1" 2"Q2" 3"Q3" 4"Q4" 5"Q5" 9"missing"
foreach biomarker in tcresult tgresult iron albumin vitaminc zinc copper{
xtile `biomarker'q = `biomarker', nq(5)
replace `biomarker'q =9 if `biomarker'q ==.
*assigning original variable label to new variable
local lbl : var  label `biomarker' 
label var `biomarker'q "`lbl'"
label values `biomarker'q  quintL 	
tabstat `biomarker', by(`biomarker'q) s(n mean min max)
}

recode bmi (0/18.49999=0 "<18.5kg/m2") (18.5/24.99999=1 "18.5-24.9kg/m2") (25/29.99999=2 "25-29.9kg/m2")(30/34.99999=3 "30-34.9kg/m2") (35/65=4 ">=35kg/m2"), gen(bmicat) label(bmicatL)

egen nomissing=rowmiss(sex race age hsizgp region bmicat diabetes highbp)
tab nomissing

drop if nomissing==1

*************************************************************************
					**# *Analysis* #
				
*************************************************************************

*Two modes for cont and cat exposures
local cont ""
local cat "i."
*exposures
local cont_exp "tcresult tgresult iron albumin vitaminc zinc copper tcresultq tgresultq ironq albuminq vitamincq zincq copperq"
local cat_exp "tcresultq tgresultq ironq albuminq vitamincq zincq copperq"
*Outcomes (could list several here if adjustment the same):highbp, diabetes
*Three adj models
local highbp_adj1 "i.sex i.race age" 
local diabetes_adj1 "i.sex i.race age"
*11 lines 
local highbp_adj2 "i.sex i.race age i.hsizgp i.region" 
local diabetes_adj2 "i.sex i.race age i.hsizgp i.region" 
*20 lines
local highbp_adj3 "i.sex i.race age i.hsizgp i.region ib1.bmicat diabetes" 
local diabetes_adj3 "i.sex i.race age i.hsizgp i.region ib1.bmicat highbp" 
*27 lines
foreach outcome of varlist (highbp diabetes){ // Loop through outcome variables (example for foreach with varlist)
foreach mode in cont cat { // Loop through cont and cat modes (example for foreach with any item)
foreach exp of local `mode'_exp { // Loop through locals that include exposure variables (example for foreach with local, loop through each item of a local)
local h=3
forval model=1/3{ // Loop through numbers to specify adjustment model (example for forval loop through consecutive numbers)
di _n  "Association between `exp' and `outcome', Model `model'"
logistic `outcome' ``mode''`exp' ``outcome'_adj`model'' if `exp'!=9
  matrix table = r(table)
  mat list table 
/*choose the below from here e.g. if you'd like SE
  table[ first number:row=estimate, second number: column variable levels (in rows in the regression output)]
  Three options for the second number: 1]=only first column; 1...]= all columns from 1 onwards; 1..4] first 4 columns only.
  I prefer showing all for quality control - visible if there is a missing or unexpected category.*/
  
  mat OR = table[1, 1...]'
  mat ll = table[5, 1...]'
  mat ul = table[6, 1...]'
  mat p = table[4, 1...]'
  
  *choose model info from ereturnlist
  mat ntotal = e(N) 
  mat psr2 = e(r2_p) 

est store a`model'

putexcel set `outcome', sheet(`mode'`exp')  modify
putexcel A1=("`outcome'") A2=("model") B2=("covariate") C2=("OR") D2=("LCI") E2=("UCI") F2=("p-value") G2=("n-total") H2=("r2 p") I2=("modelp")
putexcel D`h'=mat(ll) E`h'=mat(ul) F`h'=mat(p) G`h'=mat(ntotal) H`h'=mat(psr2) 
putexcel A`h'=matrix(OR), rownames // note that when using "rownames" the name goes into the columns specified (here covers 2 columns of names)
sleep 50 // sleep 0.5 second added as sometimes the files can struggle with saving - "file ...xlsx could not be saved" error message- extend time if needed for your model.
putexcel  A`h'=("Model `model'"), overwrite
local h=`h'+25 // this no is based on the approximate no of rows of the model output
matrix drop _all
}
lrtest a1 a2
putexcel I28=`r(p)', nformat(scientific_d2)
lrtest a2 a3
putexcel I53=`r(p)', nformat(scientific_d2)
*based on `h' at each model
}
}
}

*************************************************************************
					**# Excel combination #
				
*************************************************************************

/*To support being able to get an overview of the output or use it in a Table it helps to prepare an excel file that combines excel sheets based on the above set up

The key functions you need are =INDIRECT and =CONCAT

INDIRECT can be used to link document by defining where the information is (e.g. B6: Dataset; C6: Sheetname, D6: Wherethe estimate of interest is posted
can be used with =INDIRECT("'["&B6&"]"&C6&"'!"&D6) and thereby equals =INDIRECT("'[sample data.xlsx]Sheet1'!A1"), more info on https://exceljet.net/formula/dynamic-workbook-reference
Since the above defines output file names, sheet names and estimate position it can be easily prepared after using putexcel. Note that indirect only populates if the linked files are open at the same time. 

CONCAT allows to combine cells and add any strings "" - this is helpful when wanting to present OR and CIs as OR(lCI, uCI). When referring to numbers with decimal points an additional function has to be used to specify the no of decimal points presented such as =TEXT(numbersource, "0.00") (example 2 decimals).

Note that all functions in excel can be changed using Ctrl+H for example if you would like to change OR(lCI, uCI) to OR(lCI-uCI) you can just search and replace ", " to "-"

SEE https://github.com/aknuppel/StreamlineSTATA/blob/main/CombineOutputIndirect.xlsx for full example.

*/








