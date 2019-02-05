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
	20 - 29 = "20s"
	30 - 39 = "30s"
	40 - 49 = "40s"
	50 - 59 = "50s"
	60 - 69 = "60s"
	70 - 79 = "70s"
	80 - 90 = "80s or above"
	;
run;
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

*change age to age range (for contrasts later);
data new;
	set Sheet;
	age_range = put(age, age.);
	drop age;
;
%print(new)

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
proc corr data = Sheet plots = matrix fisher(biasadj = NO);
run;

proc logistic data = Sheet outest=betas covout;
	class gender Inet_owner;
	model response(event = '1') = visitor_ID|visit_dt|age|Yr_income|gender|Inet_owner
								/ selection = stepwise
								  slentry=0.3
                    			  slstay=0.35
                    			  details
                     			  lackfit;
    output out=pred p=phat lower=lcl upper=ucl predprob = (individual crossvalidate);
run;

proc print data = betas;
	title2 'Parameter Estimates and Covariance Matrix';
run;

proc print data = pred;
	title2 'Predicted Probabilities and 95% Confidence Limits';
run;

/*
                                   Summary of Stepwise Selection

                                                        Variable
                                    Step  Pr > ChiSq    Label

                                       1      <.0001    visitor_ID
                                       2      <.0001    Yr_income
                                       3      0.0002    age
                                       4      0.0010    Inet_owner

*/

/*
checking assumptions: how to check linearity of independent variables and log odds?
should we include interaction terms?
should visitor ID be included?
*/
proc logistic data = Sheet;
	class gender Inet_owner / param = ref;
	model response(event = '1') = visitor_ID visit_dt age Yr_income gender Inet_owner;
run;

proc logistic data = Sheet plots(maxpoints = none);
	class Inet_owner / param = ref;
	model response(event = '1') = age Yr_income Inet_owner / influence iplots;
run;

/*
                           Analysis of Maximum Likelihood Estimates

                                                Standard          Wald
              Parameter       DF    Estimate       Error    Chi-Square    Pr > ChiSq

              Intercept        1     -3.8041      0.3918       94.2490        <.0001
              age              1     -0.0260     0.00716       13.2173        0.0003
              Yr_income        1    -4.82E-6    1.216E-6       15.7269        <.0001
              Inet_owner n     1     -1.3408      0.4324        9.6153        0.0019


                                     Odds Ratio Estimates

                                               Point          95% Wald
                       Effect               Estimate      Confidence Limits

                       age                     0.974       0.961       0.988
                       Yr_income               1.000       1.000       1.000
                       Inet_owner n vs y       0.262       0.112       0.611
*/

proc logistic data = new;
	class age_range Inet_owner / param = ref;
	model response(event = '1') = age_range Yr_income Inet_owner;
	contrast '20s vs 30s' age_range 1 -1 0 0 0 0 / estimate = parm;
	contrast '20s vs 40s' age_range 1 0 -1 0 0 0 / estimate = parm;
	contrast '20s vs 50s' age_range 1 0 0 -1 0 0 / estimate = parm;
	contrast '20s vs 60s' age_range 1 0 0 0 -1 0 / estimate = parm;
	contrast '20s vs 70s' age_range 1 0 0 0 0 -1 / estimate = parm;
	contrast '30s vs 40s' age_range 0 1 -1 0 0 0 / estimate = parm;
	contrast '30s vs 50s' age_range 0 1 0 -1 0 0 / estimate = parm;
	contrast '30s vs 60s' age_range 0 1 0 0 -1 0 / estimate = parm;
	contrast '30s vs 70s' age_range 0 1 0 0 0 -1 / estimate = parm;
	contrast '40s vs 50s' age_range 0 0 1 -1 0 0 / estimate = parm;
	contrast '40s vs 60s' age_range 0 0 1 0 -1 0 / estimate = parm;
	contrast '40s vs 70s' age_range 0 0 1 0 0 -1 / estimate = parm;
	contrast '50s vs 60s' age_range 0 0 0 1 -1 0 / estimate = parm;
	contrast '50s vs 70s' age_range 0 0 0 1 0 -1 / estimate = parm;
	contrast '60s vs 70s' age_range 0 0 0 0 1 -1 / estimate = parm;
run;

/*
                            Analysis of Maximum Likelihood Estimates

                                                      Standard          Wald
         Parameter                  DF    Estimate       Error    Chi-Square    Pr > ChiSq

         Intercept                   1     -7.2210      1.0177       50.3464        <.0001
         age_range  20s              1      1.9742      1.0693        3.4084        0.0649
         age_range  30s              1      2.8970      1.0293        7.9216        0.0049
         age_range  40s              1      2.7658      1.0311        7.1942        0.0073
         age_range  50s              1      2.5722      1.0411        6.1035        0.0135
         age_range  60s              1      0.7282      1.2250        0.3534        0.5522
         age_range  70s              1      0.0683      1.4144        0.0023        0.9615
         Yr_income                   1    -4.74E-6    1.214E-6       15.2706        <.0001
         Inet_owner n                1     -1.3285      0.4326        9.4314        0.0021

                                      Odds Ratio Estimates

                                                     Point          95% Wald
                Effect                            Estimate      Confidence Limits

                age_range  20s vs 80s or above       7.201       0.885      58.559
                age_range  30s vs 80s or above      18.120       2.410     136.241
                age_range  40s vs 80s or above      15.891       2.106     119.914
                age_range  50s vs 80s or above      13.094       1.702     100.761
                age_range  60s vs 80s or above       2.071       0.188      22.855
                age_range  70s vs 80s or above       1.071       0.067      17.124
                Yr_income                            1.000       1.000       1.000
                Inet_owner n vs y                    0.265       0.113       0.618

                          Contrast Estimation and Testing Results by Row

                                      Standard                                    Wald
Contrast    Type       Row  Estimate     Error   Alpha   Confidence Limits  Chi-Square  Pr > ChiSq

20s vs 30s  PARM         1   -0.9228    0.4500    0.05   -1.8047   -0.0409      4.2063      0.0403
20s vs 40s  PARM         1   -0.7916    0.4542    0.05   -1.6818    0.0986      3.0375      0.0814
20s vs 50s  PARM         1   -0.5980    0.4765    0.05   -1.5318    0.3359      1.5751      0.2095
20s vs 60s  PARM         1    1.2459    0.8023    0.05   -0.3265    2.8184      2.4119      0.1204
20s vs 70s  PARM         1    1.9059    1.0694    0.05   -0.1901    4.0018      3.1763      0.0747
30s vs 40s  PARM         1    0.1313    0.3496    0.05   -0.5540    0.8166      0.1409      0.7073
30s vs 50s  PARM         1    0.3249    0.3782    0.05   -0.4164    1.0661      0.7379      0.3903
30s vs 60s  PARM         1    2.1688    0.7481    0.05    0.7026    3.6350      8.4049      0.0037
30s vs 70s  PARM         1    2.8287    1.0294    0.05    0.8112    4.8462      7.5518      0.0060
40s vs 50s  PARM         1    0.1936    0.3833    0.05   -0.5576    0.9448      0.2552      0.6135
40s vs 60s  PARM         1    2.0375    0.7506    0.05    0.5663    3.5088      7.3678      0.0066
40s vs 70s  PARM         1    2.6975    1.0311    0.05    0.6764    4.7185      6.8433      0.0089
50s vs 60s  PARM         1    1.8439    0.7643    0.05    0.3460    3.3419      5.8208      0.0158
50s vs 70s  PARM         1    2.5039    1.0412    0.05    0.4631    4.5446      5.7830      0.0162
60s vs 70s  PARM         1    0.6599    1.2250    0.05   -1.7411    3.0610      0.2902      0.5901

*/

proc means data = new;
run;

*predicted probability of response assuming mean income and owner;
proc logistic data = new;
	class age_range Inet_owner / param = ref;
	model response(event = '1') = age_range Yr_income Inet_owner;
	contrast '20s' intercept 1 age_range 1 0 0 0 0 0 Yr_income 217851.18 Inet_owner 1 / estimate = prob;
	contrast '30s' intercept 1 age_range 0 1 0 0 0 0 Yr_income 217851.18 Inet_owner 1 / estimate = prob;
	contrast '40s' intercept 1 age_range 0 0 1 0 0 0 Yr_income 217851.18 Inet_owner 1 / estimate = prob;
	contrast '50s' intercept 1 age_range 0 0 0 1 0 0 Yr_income 217851.18 Inet_owner 1 / estimate = prob;
	contrast '60s' intercept 1 age_range 0 0 0 0 1 0 Yr_income 217851.18 Inet_owner 1 / estimate = prob;
	contrast '70s' intercept 1 age_range 0 0 0 0 0 1 Yr_income 217851.18 Inet_owner 1 / estimate = prob;
run;

/*
owner = 1
 20s       PROB         1  0.000496  0.000270    0.05  0.000170   0.00144    194.5502      <.0001
 30s       PROB         1   0.00125  0.000575    0.05  0.000505   0.00308    209.8150      <.0001
 40s       PROB         1   0.00109  0.000517    0.05  0.000433   0.00276    207.1594      <.0001
 50s       PROB         1  0.000901  0.000441    0.05  0.000345   0.00235    204.7912      <.0001
 60s       PROB         1  0.000143  0.000116    0.05  0.000029  0.000698    119.5040      <.0001
 70s       PROB         1  0.000074  0.000079    0.05  8.941E-6  0.000608     78.0764      <.0001
owner = 0
 20s       PROB         1   0.00187  0.000729    0.05  0.000870   0.00401    258.4170      <.0001
 30s       PROB         1   0.00469   0.00123    0.05   0.00281   0.00783    416.2715      <.0001
 40s       PROB         1   0.00412   0.00111    0.05   0.00242   0.00699    409.6425      <.0001
 50s       PROB         1   0.00339   0.00103    0.05   0.00188   0.00613    351.5172      <.0001
 60s       PROB         1  0.000539  0.000384    0.05  0.000133   0.00218    111.2588      <.0001
 70s       PROB         1  0.000278  0.000280    0.05  0.000039   0.00199     66.3806      <.0001
income = 100000
 20s       PROB         1  0.000867  0.000469    0.05  0.000300   0.00250    169.3020      <.0001
 30s       PROB         1   0.00218  0.000993    0.05  0.000891   0.00532    179.9508      <.0001
 40s       PROB         1   0.00191  0.000889    0.05  0.000767   0.00475    180.2523      <.0001
 50s       PROB         1   0.00158  0.000767    0.05  0.000607   0.00409    175.1952      <.0001
 60s       PROB         1  0.000250  0.000202    0.05  0.000051   0.00121    105.4142      <.0001
 70s       PROB         1  0.000129  0.000139    0.05  0.000016   0.00106     69.4463      <.0001
income = 400000
 20s       PROB         1  0.000209  0.000130    0.05  0.000062  0.000704    186.8074      <.0001
 30s       PROB         1  0.000526  0.000288    0.05  0.000180   0.00154    189.7396      <.0001
 40s       PROB         1  0.000461  0.000259    0.05  0.000153   0.00139    186.8584      <.0001
 50s       PROB         1  0.000380  0.000216    0.05  0.000125   0.00116    191.2607      <.0001
 60s       PROB         1  0.000060  0.000052    0.05  0.000011  0.000325    127.3612      <.0001
 70s       PROB         1  0.000031  0.000035    0.05  3.485E-6  0.000277     86.3757      <.0001
*/
