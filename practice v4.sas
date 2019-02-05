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

*MERGING ALL DATA FROM 2 SHEETS BY ID, ADD LABELS, CHANGE FORMATS;
%macro combine(new, data1, data2);
	proc sql;
		create table temp as
		select &data1..visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
		from &data1 left join &data2 on (&data1..visitor_ID = &data2..visitor_ID);
	quit;

	data &new;
		set temp;
		label visitor_ID = "Visitor ID"
				visit_dt = "Visit Date"
				Yr_income = "Yearly Income"
				response = "Response"
				age = "Age";

		if gender = 'M' then
			do;
				genderN = 1;
				label genderN = "Gender";
   			end;
		else if gender = 'F' then
			do;
				genderN = 0;
				label genderN = "Gender";
   			end;

		if Inet_owner = 'y' then
			do;
				Inet_ownerN = 1;
				label Inet_ownerN = "Inet Owner";
   			end;
		else if Inet_owner = 'n' then
			do;
				Inet_ownerN = 0;
				label Inet_ownerN = "Inet Owner";
   			end;

		drop gender Inet_owner;
	;
%mend combine;
*================================================================================;
%let xlsx = 'C:\Users\yongy\Documents\My SAS Files\9.4\impression data st.xlsx';
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

proc logistic data = Sheet plots = all;
	class visitor_ID genderN Inet_ownerN;
	model response = visitor_ID visit_dt age Yr_income genderN Inet_ownerN;
run;

