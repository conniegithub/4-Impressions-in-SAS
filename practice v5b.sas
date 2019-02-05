options mprint mlogic symbolgen;
*create macro programs for ease of use;
*================================================================================;
*IMPORTING;
%macro import(address, sheet, dataset);
	proc import datafile = &address
				out = &dataset
				dbms = xlsx
				replace;
				sheet = "&sheet";
				getnames = yes;
	run;
%mend import;
*--------------------------------------------------------------------------------;
*PRINTING;
%macro print(dataset);
	proc print data = &dataset label;
	run;
%mend print;
*CONTENTS;
%macro content(dataset);
	proc contents data = &dataset;
	run;
%mend content;
*FREQUENCY;
%macro freq(dataset);
	proc freq data = &dataset;
	run;
%mend freq;
*--------------------------------------------------------------------------------;
*CONVERTING CHAR TO NUM AND RENAME, REMOVE MISSING DATA IN VISIT DATA;
%macro convert(old, new);
	data temp;
		set &old;
		visitor_ID = input(visit_ID, 8.);
		drop visit_ID;
	;

	data &new;
		set temp;
		where visitor_ID ~= . and visit_dt ~= .;
		if response = . then response = 0;
	;
%mend convert;

*MERGING ALL DATA FROM 2 SHEETS BY ID;
%macro combine(new, data1, data2);
	proc sql;
		create table &new as
		select &data1..visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
		from &data1 left join &data2 on (&data1..visitor_ID = &data2..visitor_ID);
	quit;
	;
%mend combine;

*CREATE FORMATS FOR RANGES;
proc format;
	value age
	20 - 39 = "20s - 30s"
	40 - 59 = "40s - 50s"
	60 - 79 = "60s - 70s"
	80 - 90 = "80s or above"
	;

	value wkday
	1 = "Weekend"
	2 - 6 = "Weekday"
	7 = "Weekend"
	;
run;

/*
proc format;
	value age
	20 - 29 = "20s"
	30 - 39 = "30s"
	40 - 49 = "40s"
	50 - 59 = "50s"
	60 - 69 = "60s"
	70 - 79 = "70s"
	80 - 90 = "80s or above"
	;

	value wkday
	1 = "Sunday"
	2 = "Monday"
	3 = "Tuesday"
	4 = "Wednesday"
	5 = "Thursday"
	6 = "Friday"
	7 = "Saturday"
	;
run;
*/
*================================================================================;
%let xlsx = 'C:\Users\..\impression data st.xlsx';
%let sheet1 = GMO_data;
%let sheet2 = camp_data;

*import sheet 1 with demographics data;
title "Demographics Data";
%import(&xlsx, &sheet1, demog)
%print(demog)
%content(demog)
%freq(demog)

*import sheet 2 with visit data;
*WARNING: contains missing value and value that are not in correct form;
title "Visit Data";
%import(&xlsx, &sheet2, resp)
%print(resp)
%content(resp)
%freq(resp)

*---------------------------------------------------------------------------------;
*sheet 2 with visitor ID converted to correct format and remove missing data;
%convert(resp, nResp)
%print(nResp)
%content(nResp)
%freq(nResp)

title "Combined Data";
%combine(Sheet, demog, nResp)
%print(Sheet)
%content(Sheet)
%freq(Sheet)

*change categorical variables from char to 0/1;
title "Combined Data with numeric gender and Inet owner";
data nSheet;
	set Sheet;
	if gender = 'F' then Ngender = 0;
	else if gender = 'M' then Ngender = 1;
	if Inet_owner = 'y' then NInet = 1;
	else if Inet_owner = 'n' then NInet = 0;
	drop gender Inet_owner;
;
%print(nSheet)
%content(nSheet)
%freq(nSheet)

*change age to age range, change dates to weekday/weekend (for contrasts later);
title "Combined Data with age range and weekday";
data newSheet;
	set Sheet;
	age_range = put(age, age.);
	day_type = put(weekday(visit_dt), wkday.);
	drop age visit_dt;
;
%print(newSheet)
%content(newSheet)
%freq(newSheet)

/*
proc freq data = Sheet;
	table Yr_income;
run;
*/

title;

*statistical analysis
/*
questions:
1. which groups responded in overall dataset?
2. in those who responded, which groups had the most response?
3. in those who responded, on which dates did most response occur?
*/
*================================================================================;
/*
proc corr data = Sheet plots = matrix fisher(biasadj = NO) plots(maxpoints = none);
run;

*income, age, owner;
proc logistic data = Sheet outest=betas covout;
	class gender Inet_owner;
	model response(event = '1') = visit_dt age Yr_income gender Inet_owner
								/ selection = stepwise
								  slentry=0.3
                    			  slstay=0.35
                    			  details
                     			  lackfit;
    output out=pred p=phat lower=lcl upper=ucl predprob = (individual crossvalidate);
run;
*/

*check for extreme values;
proc univariate data = Sheet;
	var age Yr_income;
run;

*checking for collinearity;
proc reg data = nSheet plots(maxpoints = none);
	model response = age Yr_income Ngender NInet / tol vif collin;
run;



*age_range, Yr_income, Inet_owner, day_type;
proc logistic data = newSheet outest=betas covout;
	class day_type age_range gender Inet_owner;
	model response(event = '1') = day_type age_range Yr_income gender Inet_owner
								/ selection = stepwise
								  slentry=0.3
                    			  slstay=0.35
                    			  details
                     			  lackfit;
    output out=pred p=phat lower=lcl upper=ucl predprob = (individual crossvalidate);
run;

proc logistic data = newSheet;
	class day_type age_range Inet_owner / param = ref;
	model response(event = '1') = day_type age_range Yr_income Inet_owner;
run;

proc logistic data = newSheet plots(maxpoints = none);
	class day_type age_range Inet_owner / param = ref;
	model response(event = '1') = day_type age_range Yr_income Inet_owner / influence iplots;
run;

proc logistic data = newSheet plots(maxpoints = none) plots(only label) = (phat leverage dpc);
	class day_type age_range Inet_owner / param = ref;
	model response(event = '1') = day_type age_range Yr_income Inet_owner;
run;

proc logistic data = newSheet;
	class day_type age_range Inet_owner / param = ref;
	model response(event = '1') = day_type age_range Yr_income Inet_owner;
	contrast '20 - 30s vs 40 - 50s' age_range 1 -1 0 /e estimate = parm;
	contrast '20 - 30s vs 60 - 70s' age_range 1 0 -1 /e estimate = parm;
	contrast '20 - 30s vs 80s or above' age_range 1 0 0 /e estimate = parm;
	contrast '40 - 50s vs 60 - 70s' age_range 0 1 -1 /e estimate = parm;
	contrast '40 - 50s vs 80s or above' age_range 0 1 0 /e estimate = parm;
	contrast '60 - 70s vs 80s or above' age_range 0 0 1 /e estimate = parm;
	contrast 'Weekday vs Weekend' day_type 1 /e estimate = parm;
	contrast 'Inet Owner vs Not' Inet_owner 1 /e estimate = parm;
run;

proc means data = newSheet;
run;

*predicted probability of response assuming mean income;
proc logistic data = newSheet;
	class day_type age_range Inet_owner / param = ref;
	model response(event = '1') = day_type age_range Yr_income Inet_owner;
	contrast '20 - 30s with Inet on weekday' intercept 1 age_range 1 0 0 Yr_income 217851.18 Inet_owner 1 day_type 1 / estimate = prob;
	contrast '40 - 50s with Inet on weekday' intercept 1 age_range 0 1 0 Yr_income 217851.18 Inet_owner 1 day_type 1 / estimate = prob;
	contrast '60 - 70s with Inet on weekday' intercept 1 age_range 0 0 1 Yr_income 217851.18 Inet_owner 1 day_type 1 / estimate = prob;
	contrast '80s or above with Inet on weekday' intercept 1 age_range 0 0 0 Yr_income 217851.18 Inet_owner 1 day_type 1 / estimate = prob;

	contrast '20 - 30s with Inet on weekend' intercept 1 age_range 1 0 0 Yr_income 217851.18 Inet_owner 1 day_type 0 / estimate = prob;
	contrast '40 - 50s with Inet on weekend' intercept 1 age_range 0 1 0 Yr_income 217851.18 Inet_owner 1 day_type 0 / estimate = prob;
	contrast '60 - 70s with Inet on weekend' intercept 1 age_range 0 0 1 Yr_income 217851.18 Inet_owner 1 day_type 0 / estimate = prob;
	contrast '80s or above with Inet on weekend' intercept 1 age_range 0 0 0 Yr_income 217851.18 Inet_owner 1 day_type 0 / estimate = prob;

	contrast '20 - 30s without Inet on weekday' intercept 1 age_range 1 0 0 Yr_income 217851.18 Inet_owner 0 day_type 1 / estimate = prob;
	contrast '40 - 50s without Inet on weekday' intercept 1 age_range 0 1 0 Yr_income 217851.18 Inet_owner 0 day_type 1 / estimate = prob;
	contrast '60 - 70s without Inet on weekday' intercept 1 age_range 0 0 1 Yr_income 217851.18 Inet_owner 0 day_type 1 / estimate = prob;
	contrast '80s or above without Inet on weekday' intercept 1 age_range 0 0 0 Yr_income 217851.18 Inet_owner 0 day_type 1 / estimate = prob;

	contrast '20 - 30s without Inet on weekend' intercept 1 age_range 1 0 0 Yr_income 217851.18 Inet_owner 0 day_type 0 / estimate = prob;
	contrast '40 - 50s without Inet on weekend' intercept 1 age_range 0 1 0 Yr_income 217851.18 Inet_owner 0 day_type 0 / estimate = prob;
	contrast '60 - 70s without Inet on weekend' intercept 1 age_range 0 0 1 Yr_income 217851.18 Inet_owner 0 day_type 0 / estimate = prob;
	contrast '80s or above without Inet on weekend' intercept 1 age_range 0 0 0 Yr_income 217851.18 Inet_owner 0 day_type 0 / estimate = prob;
run;
