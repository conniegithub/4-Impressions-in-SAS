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

*PRINTING;
%macro print(dataset);
	proc print data = &dataset;
	run;
%mend print;

*CONTENTS AND FREQUENCE SUMMARY;
%macro confreq(dataset);
	proc contents data = &dataset;
	run;

	proc freq data = &dataset;
	run;
%mend confreq;

*SUBSETTING, CONVERTING CHAR TO NUM AND RENAME, REMOVE MISSING DATA;
%macro subconv(old, new, response);
	data temp;
		set &old;
		if response = &response;
		visitor_ID = input(visit_ID, 8.);
		drop visit_ID;
	;

	data &new;
		set temp;
		where visitor_ID ~= . and visit_dt ~= .;
		if response = . then response = 0;
	;
%mend subconv;

*MATCHING RESPONSES FROM 2 SHEETS BY ID;
*double periods are needed since one is needed to add a phrase following a macro;
%macro match(new, data1, data2, loc);
	proc sql;
		create table &new as
		select &data1..visitor_ID from &data1
		where &data1..visitor_ID &loc (select unique &data2..visitor_ID from &data2);
	quit;
%mend match;

*MERGING ALL DATA FROM 2 SHEETS BY ID;
%macro combine(new, data1, data2, data3);
	proc sql;
		create table &new as
		select &data1..visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
		from &data1 left join &data2 on (&data1..visitor_ID = &data2..visitor_ID)
		left join &data3 on (&data1..visitor_ID = &data3..visitor_ID);
	quit;
%mend combine;
*================================================================================;
%let xlsx = 'C:\Users\yongy\Documents\My SAS Files\9.4\impression data st.xlsx';
%let sheet1 = GMO_data;
%let sheet2 = camp_data;

*import sheet 1 with demographics data;
title "Demographics Data";
%import(&xlsx, &sheet1, demog)
%print(demog)
%confreq(demog)

*import sheet 2 with visit data;
*WARNING: contains missing value and value that are not in correct form;
title "Visit Data";
%import(&xlsx, &sheet2, res)
%print(res)
%confreq(res)

*---------------------------------------------------------------------------------;
*sheet 2 with visitor ID converted to correct format and remove missing data;
%subconv(res, nRes, 1 or response = .)
%print(nRes)
%confreq(nRes)

title "Combined Data";
proc sql;
	create table sheet as
	select demog.visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
	from demog left join nRes on (demog.visitor_ID = nRes.visitor_ID);
quit;
%print(Sheet)
%confreq(Sheet)
*---------------------------------------------------------------------------------;
*sheet 2 with at least 1 response;
title "With Response";
%subconv(res, yesRes, 1)
%print(yesRes)
%confreq(yesRes)

*sheet 2 with at least 1 non-response;
*this cannot be used as no response at all since they might have responded in a previous visit;
title "With No Response";
%subconv(res, noRes, .)
%print(noRes)
%confreq(noRes)
/*
*find where missing values appear;
data miss;
	set Res;
	where visit_ID = '' or visit_dt = .;
;
title;
%print(miss)
*/

*IDs with at least 1 response;
%match(ResID, demog, yesRes, in)
%print(ResID)
*IDs with no response at all;
%match(NoResID, demog, yesRes, not in)
%print(NoResID)

*complete sheet for at least 1 response;
title "With Response (combined data)";
%combine(ResSheet, ResID, demog, yesRes)
%print(ResSheet)
%confreq(ResSheet)

*complete sheet for no response at all;
title "With No Response (combined data)";
%combine(NoResSheet, NoResID, demog, noRes)
%print(NoResSheet)
%confreq(NoResSheet)

title;

*statistical analysis
/*
questions:
1. which groups responded in overall dataset?
2. in those who responded, which groups had the most response?
3. in those who responded, on which dates did most response occur?
*/
*================================================================================;
proc corr data = Sheet plots = matrix fisher(biasadj = NO);
run;


