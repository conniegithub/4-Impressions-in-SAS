%let txt = 'C:\Users\yongy\Documents\My SAS Files\9.4\impression data st.txt';
%let xlsx = 'C:\Users\yongy\Documents\My SAS Files\9.4\impression data st.xlsx';


*import tab-delimited TXT file using DATA step;
data pracD;
	infile &txt delimiter = '09'x missover firstobs = 2;
	length gender $ 1 Inet_owner $ 1;
	input visitor_ID gender $ age Inet_owner $ Yr_income;
run;

*import EXCEL file using IMPORT procedure;
proc import datafile = &xlsx
			out=pracI
			dbms = xlsx
			replace;
			sheet = "GMO_data";
			getnames = yes;
run;


*test print data and attributes containing demographics data;
proc print data = pracD;
run;

proc contents data = pracD;
run;

proc print data = pracI;
run;

proc contents data = pracI;
run;

*compare;
proc compare base = pracD compare = pracI;
run;


*import sheet 2 containing response data and visit dates;
proc import datafile = &xlsx
			out=pracI2
			dbms = xlsx
			replace;
			sheet = "camp_data";
			getnames = yes;
run;

proc print data = pracI2;
run;

proc contents data = pracI2;
run;

*subset for with response, convert char to num and rename, contains corresponding visit dates;
data pracRes;
	set pracI2;
	if response = 1;
	visitor_ID = input(visit_ID, 8.);
	drop visit_ID;
run;

proc print data = pracRes;
run;

proc contents data = pracRes;
run;

data pracNoRes;
	set pracI2;
	if response = .;
	visitor_ID = input(visit_ID, 8.);
	drop visit_ID;
run;

proc print data = pracNoRes;
run;

proc contents data = pracNoRes;
run;

*check for returning visits;
proc freq data = pracRes;
	tables visit_ID visit_dt;
run;

*merging sheet 2 responses to sheet 1 using pracRes;
proc sql;
	create table ResSheet as
	select pracRes.visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
	from pracRes left join pracI on (pracRes.visitor_ID = pracI.visitor_ID);
quit;

title "With Response";
proc print data = ResSheet;
run;

proc contents data = ResSheet;
run;

/*
*subset for no response;
data pracNoRes;
	set pracI2;
	if response = .;
	visitor_ID = input(visit_ID, 8.);
	drop visit_ID;
run;

proc print data = pracNoRes;
run;

proc contents data = pracNoRes;
run;
*/

*selecting matching responses and no responses from sheet 1;
proc sql;
	create table ResID as
	select pracI.visitor_ID from pracI
	where pracI.visitor_ID in (select pracRes.visitor_ID from pracRes);

	create table NoResID as
	select pracI.visitor_ID from pracI
	where pracI.visitor_ID not in (select pracRes.visitor_ID from pracRes);
quit;

proc print data = ResID;
run;

proc print data = NoResID;
run;

/*
*merging sheet 2 responses to sheet 1 using ResID;
proc sql;
	create table ResSheet as
	select ResID.visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
	from ResID left join pracRes on (ResID.visitor_ID = pracRes.visitor_ID)
	left join pracI on (ResID.visitor_ID = pracI.visitor_ID);
quit;

proc compare base = ResSheet compare = NoResSheet;
run;
*/

*merging sheet 2 no responses to sheet 1 using ResID;
proc sql;
	create table NoResSheet as
	select NoResID.visitor_ID, visit_dt, gender, age, Inet_owner, Yr_income, response
	from NoResID left join pracNoRes on (NoResID.visitor_ID = pracNoRes.visitor_ID)
	left join pracI on (NoResID.visitor_ID = pracI.visitor_ID);
quit;

title "Without Response";
proc print data = NoResSheet;
run;

proc contents data = NoResSheet;
run;

