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
title "Combined Data with numeric gender (F=0) and Inet owner (n=0), and day type";
data nSheet;
	set Sheet;
	if gender = 'F' then Ngender = 0;
	else if gender = 'M' then Ngender = 1;
	if Inet_owner = 'y' then NInet = 1;
	else if Inet_owner = 'n' then NInet = 0;
	day_type = weekday(visit_dt);

	drop gender Inet_owner visit_dt;
;
%print(nSheet)
%content(nSheet)
%freq(nSheet)

title;

*#2;
*---------------------------------------------------------------------------;
title "visit frequency for each buyer with respect to response rate";
proc sql;
create table test1 as
	select  visitor_ID, count(visitor_ID) as freq, sum(response) as indiv_buys
	from Sheet
	group by visitor_ID;
quit;
%print(test1)

title "visit frequency rate with respect to total response rate";
proc sql;
create table test2 as
	select freq, count(visitor_ID) as indv, sum(indiv_buys) as tot_buys, 
	(calculated tot_buys)/(calculated indv) as conversion_rate
	from test1
	group by freq;
quit;
%print(test2)

/*
                                                               conversion_
                            Obs    freq    indv    tot_buys        rate

                             1       2        1        0         0.00000
                             2       3     6348       42         0.00662
                             3       4     1918       10         0.00521
                             4       5      104        3         0.02885
                             5       6       42        0         0.00000
                             6       7       69        0         0.00000
                             7       8       13        0         0.00000
                             8      13        7        1         0.14286
                             9      14        8        0         0.00000

*/
*---------------------------------------------------------------------------;

*creating bins;
%macro quint(dsn,var,quintvar);

/* calculate the cutpoints (variable values) for the corresponding quintiles */
proc univariate noprint data=&dsn;
  var &var;
  output out=quintile pctlpts= 10 20 30 40 50 60 70 80 90 pctlpre=pct;
run;

/* write the quintiles to macro variables */
data _null_;
  set quintile;
  call symput('q1',pct10) ;
  call symput('q2',pct20) ;
  call symput('q3',pct30) ;
  call symput('q4',pct40) ;
  call symput('q5',pct50) ;
  call symput('q6',pct60) ;
  call symput('q7',pct70) ;
  call symput('q8',pct80) ;
  call symput('q9',pct90) ;
run;

/* create the new variable in the main dataset */
data &dsn;
  set &dsn;
  /*indicator marking if variable value falls in the specified range*/
  q_1_&var = 0;
  q_2_&var = 0;
  q_3_&var = 0;
  q_4_&var = 0;
  q_5_&var = 0;
  q_6_&var = 0;
  q_7_&var = 0;
  q_8_&var = 0;
  q_9_&var = 0;
  q_10_&var = 0;

       if &var le &q1 then do; rank_&quintvar=1; q_1_&var = 1; end;
  else if &var le &q2 then do; rank_&quintvar=2; q_2_&var = 1; end;
  else if &var le &q3 then do; rank_&quintvar=3; q_3_&var = 1; end; 
  else if &var le &q4 then do; rank_&quintvar=4; q_4_&var = 1; end;
  else if &var le &q5 then do; rank_&quintvar=5; q_5_&var = 1; end;
  else if &var le &q6 then do; rank_&quintvar=6; q_6_&var = 1; end;
  else if &var le &q7 then do; rank_&quintvar=7; q_7_&var = 1; end;
  else if &var le &q8 then do; rank_&quintvar=8; q_8_&var = 1; end;
  else if &var le &q9 then do; rank_&quintvar=9; q_9_&var = 1; end;
                      else do; rank_&quintvar=10; q_10_&var = 1; end;

  /*only marked variable values in corresponding bin*/
  bin_1_&var = &var * q_1_&var;
  bin_2_&var = &var * q_2_&var;
  bin_3_&var = &var * q_3_&var;
  bin_4_&var = &var * q_4_&var;
  bin_5_&var = &var * q_5_&var;
  bin_6_&var = &var * q_6_&var;
  bin_7_&var = &var * q_7_&var;
  bin_8_&var = &var * q_8_&var;
  bin_9_&var = &var * q_9_&var;
  bin_10_&var = &var * q_10_&var;
 
run;

%mend quint;

%macro create_bin(data=,vars=);
  %local i;

  /*selecting the ith variable from a macro list of variables*/
  %let i = 1;
  %do %while ( "%scan(&vars,&i)" ne "");
     %quint(&data, %scan(&vars,&i), %scan(&vars,&i));
     %let i=%eval(&i+1);
  %end;
%mend create_bin;


%let variables = Yr_income age;

%create_bin(data = nSheet, vars = &variables)
%print(nSheet)
%content(nSheet)

data nSheet1;
	set nSheet(keep = NInet Ngender day_type rank_Yr_income rank_age response visitor_ID);
run;
%print(nSheet1)
%content(nSheet1)


%global q_vars rank_vars bin_vars;
proc contents data = nSheet out=contents;
run;

proc sql;
select name
       into : q_vars separated by ', '
       from contents
       where name like 'q_%';
select name
       into : rank_vars separated by ', '
       from contents
       where name like 'rank_%';
select name
       into : bin_vars separated by ', '
       from contents
       where name like 'bin_%';
quit;

proc sql;
create table hist as
	select &rank_vars from nSheet;
quit;
%print(hist)
%freq(hist)

/*
proc sql;
create table hist as
	select rank_age, count(rank_age) as num_age from nSheet
	group by rank_age;
quit;
%print(hist)

proc gplot data = hist;
	plot num_age*rank_age = 'Q' / vaxis = 0 to 3000 by 500;
run;

proc sql;
create table hist as
	select rank_Yr_income, count(rank_Yr_income) as num_Yr_income from nSheet
	group by rank_Yr_income;
quit;
%print(hist)

proc gplot data = hist;
	plot num_Yr_income*rank_Yr_income = 'Q' / vaxis = 0 to 3000 by 500;
run;
*/



*data analysis
*================================================================================;

*check for extreme values;
proc univariate data = nSheet noprint;
	var age Yr_income;
	output out = temp
	std = astd ystd
	mean = amean ymean;
run;

%macro outlier(vstd, vmean, stdata, data, var);
	data eSheet;
		set &stdata;
		&var.ulimit = &vmean + 3*&vstd;
		&var.llimit = &vmean - 3*&vstd;
	run;
	%print(eSheet)

	proc sql;
	create table oSheet as
		select &var, &var.ulimit, &var.llimit from &data, eSheet 
		where &var > &var.ulimit or &var < &var.llimit;
	quit;
	%print(oSheet)
%mend outlier;

%outlier(astd, amean, temp, nSheet, age)
%outlier(ystd, ymean, temp, nSheet, Yr_income)

/*
outlier if 3 standard deviations away from the mean:
if age > 115.954603905 or age < -6.91181748
if Yr_income > 602161.60325 or Yr_income < -166459.24056999
*/

proc means data = nSheet;
run;

proc sgplot data = nSheet;
	histogram age;
	density age;
run;

proc sgplot data = nSheet;
	hbox age;
run;

proc sgplot data = nSheet;
	histogram Yr_income;
	density Yr_income;
run;

proc sgplot data = nSheet;
	hbox Yr_income;
run;

*checking for collinearity;
proc reg data = nSheet plots(maxpoints = none);
	model response = age Yr_income day_type Ngender NInet / tol vif collin;
run;

*logistic regression;
proc logistic data = nSheet1 outest=betas;
	class NInet Ngender day_type rank_Yr_income rank_age;
	model response(event = '1') = NInet Ngender day_type rank_Yr_income rank_age
								/ selection = stepwise
								  slentry = 0.01
                    			  slstay = 0.01
                    			  details
                     			  lackfit;
    output out=pred p=phat lower=lcl upper=ucl predprob = (individual crossvalidate);
run;

proc logistic data = nSheet1;
	class NInet Ngender day_type rank_Yr_income rank_age;
	model response(event = '1') = NInet Ngender day_type rank_Yr_income rank_age / firth;
run;
