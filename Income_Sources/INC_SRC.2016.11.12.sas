/* Section 1: Set Conditions */;
%LET TAX=/Directory/; LIBNAME TAX "&TAX.";

/* Section 4 - Using SCF 2013 Data. */;
/* PART 1 */;
%LET SCF13=/Directory2/; LIBNAME SCF13 "&SCF13.";

%MACRO ACONV(F);
((&F=2)*52.18 + (&F=3)*26.09 + (&F=4)*12 + (&F=5)*4 + (&F=6) + (&F=8) + 
(&F=11)*2 + (&F=12)*6 + (&F=31)*24 + (&F=14) + (&F=22))
%MEND ACONV;

PROC SQL;
	CREATE TABLE TAX.SCF13 AS
	SELECT
	YY1, Y1, SUBSTR(PUT(Y1,5.),5,1) AS WAVE,
	X7001 AS PEUNUM, X101 AS HHNUM,
	X42001 AS WGT,
	X8000,
	X8021 AS SEX,
	X14 AS AGE, X19 AS SP_AGE,
	CASE WHEN X4104 IN(-1,1900:2013) THEN 1 ELSE 0 END AS DISAB,
	CASE WHEN X4704 IN(-1,1900:2013) THEN 1 ELSE 0 END AS SP_DISAB,
	MIN(SUM( (X108 = 4), (X114 = 4), (X120 = 4)+ (X126 = 4), (X132 = 4), (X202 = 4), (X208 = 4), (X214 = 4), (X220 = 4), (X226 = 4), 0), 2) AS CHILDREN,
	X5901 AS EDUC_GRADE, X5905 AS EDUC_DEGREE,
	CASE X5906 WHEN 1 THEN 1 WHEN 5 THEN 0 ELSE . END AS SERVEMIL,
	CASE X7004 WHEN 1 THEN 1 WHEN 5 THEN 0 ELSE . END AS HISPANIC,
	X6809 AS RACEID1, X6810 AS RACEID2,
	X7372 AS MARITAL,
	X8020 AS RELP_R, X108 AS RELP3, X114 AS RELP4, X120 AS RELP5, X126 AS RELP6, X132 AS RELP7, X202 AS RELP8, X208 AS RELP9, X214 AS RELP10, X220 AS RELP11, X226 AS RELP12,
	CASE X7020 WHEN 1 THEN 0 WHEN 2 THEN 1 ELSE . END AS SPOUSE_PRESENT,
	X5702 AS WAGES,
	X5704 AS SELF_INC,
	X5706 AS NONTAX_INVEST,
	X5708 AS OTHER_INTEREST,
	X5710 AS DIVIDEND,
	X5712 AS STOCKS,
	CASE	WHEN X7193 IN(0,-1) THEN 0
			WHEN X7193 = 25 THEN X3928/2
			WHEN X7193 IN(6,8) THEN X3928
			WHEN X7193 = 11 THEN X3928*2
			WHEN X7193 = 5 THEN X3928*4
			WHEN X7193 = 8 THEN X3928*6
			WHEN X7193 = 4 THEN X3928*12
			WHEN X7193 = 3 THEN X3928*26.25
			WHEN X7193 = 2 THEN X3928*52.50
			WHEN X7193 = 1 THEN X3928*365
			ELSE . END AS ANNUAL_ACTIVITY,
	X5714 AS RENT_TRUST,
	X5716 AS UNEMP_COMP,
	X5718 AS CHILD_SUP_ALIM,
	X5722 AS PENSIONS,
	X5720 AS TANF_SSI_FOODSTMP,
	X5724 AS OTHER_INC,

	X6558 + X6566 + X6574 + MAX(0,(X6464*%ACONV(X6465)))
		+ MAX(0,(X6469*%ACONV(X6470))) + MAX(0,(X6474*%ACONV(X6475)))
		+ MAX(0,(X6479*%ACONV(X6480))) + MAX(0,(X6965*%ACONV(X6966)))
		+ MAX(0,(X6971*%ACONV(X6972))) + MAX(0,(X6977*%ACONV(X6978)))
		+ MAX(0,(X6983*%ACONV(X6984))) AS PENACCTWD,
	X5729 + CALCULATED PENACCTWD - X5706 - X5720 AS TINCOME,
/*	X5744, X5746,*/

	CASE WHEN CALCULATED TINCOME = -1 THEN 0 ELSE CALCULATED TINCOME END AS AGI,

	X5702+X5704+X5706+X5708+X5710+X5712+X5714+X5716+X5718+X5720+X5722+X5724 AS CALCINC,
	CALCULATED CALCINC + CALCULATED PENACCTWD AS CALCINC2,
	X5712, X5729 AS TOTINC
	FROM SCF13.SCF2013
	WHERE CALCULATED WAVE = "1";
QUIT;

DATA TAX.SCF13_NW;
	SET SCF13.SCFP2013;
	WAVE = INPUT(SUBSTR(PUT(Y1,5.),5,1),BEST8.);
	IF WAVE = 1;
	IF YY1 = 3764 THEN DO;
		ASSET = ASSET + 2453000;
		NETWORTH = 0;
	END;
	DROP WAVE;
	KEEP YY1 NETWORTH ASSET DEBT;
RUN;

DATA TAX.SCF13B;
	MERGE TAX.SCF13 TAX.SCF13_NW;
	BY YY1;
	IF CALCINC2 IN(-999999:-2,.) THEN INCCAT = -1;
		ELSE IF CALCINC2 <= 00000000 THEN INCCAT =      -1;
		ELSE IF CALCINC2 <= 00009999 THEN INCCAT =       0;
		ELSE IF CALCINC2 <= 00019999 THEN INCCAT =   10000;
		ELSE IF CALCINC2 <= 00029999 THEN INCCAT =   20000;
		ELSE IF CALCINC2 <= 00039999 THEN INCCAT =   30000;
		ELSE IF CALCINC2 <= 00049999 THEN INCCAT =   40000;
		ELSE IF CALCINC2 <= 00074999 THEN INCCAT =   50000;
		ELSE IF CALCINC2 <= 00099999 THEN INCCAT =   75000;
		ELSE IF CALCINC2 <= 00249999 THEN INCCAT =  100000;
		ELSE IF CALCINC2 <= 00499999 THEN INCCAT =  250000;
		ELSE IF CALCINC2 <= 00999999 THEN INCCAT =  500000;
		ELSE IF CALCINC2 <= 01999999 THEN INCCAT = 1000000;
		ELSE IF CALCINC2 <= 04999999 THEN INCCAT = 2000000;
		ELSE IF CALCINC2 <= 09999999 THEN INCCAT = 5000000;
		ELSE INCCAT = 10000000;
	IF NETWORTH IN(-300000000:-2) THEN NWCAT = -1;
		ELSE IF NETWORTH IN(.,-1,0) THEN NWCAT = 0;
/*		ELSE IF NETWORTH < 100 THEN NWCAT = 2;*/
		ELSE IF NETWORTH < 1000 THEN NWCAT = 3;
		ELSE IF NETWORTH < 10000 THEN NWCAT = 4;
		ELSE IF NETWORTH < 100000 THEN NWCAT = 5;
		ELSE IF NETWORTH < 1000000 THEN NWCAT = 6;
		ELSE IF NETWORTH < 10000000 THEN NWCAT = 7;
		ELSE IF NETWORTH < 100000000 THEN NWCAT = 8;
		ELSE IF NETWORTH < 10000000000 THEN NWCAT = 9;
RUN;

DATA TAX.SCF13_TAXDATA;
	MERGE TAX.SCF13B;
	BY YY1;
		ABS_WAGES = ABS(WAGES);
		ABS_SELF_INC = ABS(SELF_INC);
		ABS_NONTAX_INVEST = ABS(NONTAX_INVEST);
		ABS_OTHER_INTEREST = ABS(OTHER_INTEREST);
		ABS_DIVIDEND = ABS(DIVIDEND);
		ABS_STOCKS = ABS(STOCKS);
		ABS_RENT_TRUST = ABS(RENT_TRUST);
		ABS_UNEMP_COMP = ABS(UNEMP_COMP);
		ABS_CHILD_SUP_ALIM = ABS(CHILD_SUP_ALIM);
		ABS_PENSIONS = ABS(PENSIONS);
		ABS_TANF_SSI_FOODSTMP = ABS(TANF_SSI_FOODSTMP);
		ABS_OTHER_INC = ABS(OTHER_INC);
		ABS_PENACCTWD = ABS(PENACCTWD);
		INC_ACTIVITY = SUM(OF ABS_:);
	IF INC_ACTIVITY IN(-999999:-2,.,1,0,-1) THEN INCACT_CAT =   -1;
		ELSE IF INC_ACTIVITY <= 00009999 THEN INCACT_CAT =       0;
		ELSE IF INC_ACTIVITY <= 00019999 THEN INCACT_CAT =   10000;
		ELSE IF INC_ACTIVITY <= 00029999 THEN INCACT_CAT =   20000;
		ELSE IF INC_ACTIVITY <= 00039999 THEN INCACT_CAT =   30000;
		ELSE IF INC_ACTIVITY <= 00049999 THEN INCACT_CAT =   40000;
		ELSE IF INC_ACTIVITY <= 00074999 THEN INCACT_CAT =   50000;
		ELSE IF INC_ACTIVITY <= 00099999 THEN INCACT_CAT =   75000;
		ELSE IF INC_ACTIVITY <= 00249999 THEN INCACT_CAT =  100000;
		ELSE IF INC_ACTIVITY <= 00499999 THEN INCACT_CAT =  250000;
		ELSE IF INC_ACTIVITY <= 00999999 THEN INCACT_CAT =  500000;
		ELSE IF INC_ACTIVITY <= 01999999 THEN INCACT_CAT = 1000000;
		ELSE IF INC_ACTIVITY <= 04999999 THEN INCACT_CAT = 2000000;
		ELSE IF INC_ACTIVITY <= 09999999 THEN INCACT_CAT = 5000000;
		ELSE INCACT_CAT = 10000000;
	IF ASSET IN(-300000000:-2) THEN ASSCAT = -1;
		ELSE IF ASSET IN(.,-1,0) THEN ASSCAT = 0;
/*		ELSE IF ASSET < 100 THEN ASSCAT = 2;*/
		ELSE IF ASSET < 1000 THEN ASSCAT = 3;
		ELSE IF ASSET < 10000 THEN ASSCAT = 4;
		ELSE IF ASSET < 100000 THEN ASSCAT = 5;
		ELSE IF ASSET < 1000000 THEN ASSCAT = 6;
		ELSE IF ASSET < 10000000 THEN ASSCAT = 7;
		ELSE IF ASSET < 100000000 THEN ASSCAT = 8;
		ELSE IF ASSET < 10000000000 THEN ASSCAT = 9;
RUN;

%MACRO INCTYPE(VARNAME);
PROC SQL;
	CREATE TABLE TAX.SCF13_&VARNAME.1 AS
	SELECT
	&VARNAME.,
	FREQ, POP,
	WAGES, GAINS, DIVIDEND, RENT_TRUST, RETIREMENT, ASSISTANCE, OTHER,
	WAGES+GAINS+DIVIDEND+RENT_TRUST+RETIREMENT+ASSISTANCE+OTHER FORMAT=COMMA25.0 AS TOTINC
	FROM (SELECT
			&VARNAME.,
			COUNT(WGT) FORMAT=COMMA25.0 AS FREQ,
			SUM(WGT) FORMAT=COMMA25.0 AS POP,
			SUM(WAGES*WGT) FORMAT=COMMA25.0 AS WAGES,
			SUM(RENT_TRUST*WGT) FORMAT=COMMA25.0 AS RENT_TRUST,
			SUM(STOCKS*WGT) FORMAT=COMMA25.0 AS GAINS,
			SUM(DIVIDEND*WGT) FORMAT=COMMA25.0 AS DIVIDEND,
			SUM(PENSIONS*WGT) FORMAT=COMMA25.0 AS RETIREMENT,
			SUM(TANF_SSI_FOODSTMP*WGT) FORMAT=COMMA25.0 AS ASSISTANCE,
			SUM(OTHER_INC*WGT) + SUM(SELF_INC*WGT) + SUM(UNEMP_COMP*WGT) + SUM(CHILD_SUP_ALIM*WGT) + SUM(NONTAX_INVEST*WGT) + SUM(OTHER_INTEREST*WGT)
				FORMAT=COMMA25.0 AS OTHER
		FROM TAX.SCF13_TAXDATA
		GROUP BY &VARNAME.);
QUIT;
PROC SQL;
	CREATE TABLE TAX.SCF13_&VARNAME.2 AS
	SELECT
	&VARNAME.,
	FREQ, POP,
	ABS_WAGES, ABS_GAINS, ABS_DIVIDEND, ABS_RENT_TRUST, ABS_RETIREMENT, ABS_ASSISTANCE, ABS_OTHER,
	ABS_WAGES+ABS_GAINS+ABS_DIVIDEND+ABS_RENT_TRUST+ABS_RETIREMENT+ABS_ASSISTANCE+ABS_OTHER FORMAT=COMMA25.0 AS ABSINC
	FROM (SELECT
			&VARNAME.,
			COUNT(WGT) FORMAT=COMMA25.0 AS FREQ,
			SUM(WGT) FORMAT=COMMA25.0 AS POP,
			SUM(ABS_WAGES*WGT) FORMAT=COMMA25.0 AS ABS_WAGES,
			SUM(ABS_RENT_TRUST*WGT) FORMAT=COMMA25.0 AS ABS_RENT_TRUST,
			SUM(ABS_STOCKS*WGT) FORMAT=COMMA25.0 AS ABS_GAINS,
			SUM(ABS_DIVIDEND*WGT) FORMAT=COMMA25.0 AS ABS_DIVIDEND,
			SUM(ABS_PENSIONS*WGT) FORMAT=COMMA25.0 AS ABS_RETIREMENT,
			SUM(ABS_TANF_SSI_FOODSTMP*WGT) FORMAT=COMMA25.0 AS ABS_ASSISTANCE,
			SUM(ABS_OTHER_INC*WGT) + SUM(ABS_SELF_INC*WGT) + SUM(ABS_UNEMP_COMP*WGT) + SUM(ABS_CHILD_SUP_ALIM*WGT) + SUM(ABS_NONTAX_INVEST*WGT) + SUM(ABS_OTHER_INTEREST*WGT)
				FORMAT=COMMA25.0 AS ABS_OTHER
		FROM TAX.SCF13_TAXDATA
		GROUP BY &VARNAME.);
QUIT;

PROC TRANSPOSE DATA=TAX.SCF13_&VARNAME.1 OUT=TAX.SCF13_&VARNAME.1B(RENAME=(_NAME_=INCTYPE COL1=VALUE &VARNAME.=CAT));
	BY &VARNAME. FREQ POP TOTINC;
RUN;

PROC TRANSPOSE DATA=TAX.SCF13_&VARNAME.2 OUT=TAX.SCF13_&VARNAME.2B(RENAME=(_NAME_=INCTYPE COL1=ABSVALUE &VARNAME.=CAT));
	BY &VARNAME. FREQ POP ABSINC;
RUN;

DATA TAX.SCF13_&VARNAME.3;
	MERGE	TAX.SCF13_&VARNAME.1B
			TAX.SCF13_&VARNAME.2B(DROP=INCTYPE);
RUN;

DATA TAX.SCF13_FINAL_&VARNAME.;
	RETAIN CATTYPE CAT FREQ POP INCTYPE VALUE TOTINC ABSVALUE ABSINC;
	FORMAT CATTYPE $10. PCT PCT2 9.6;
	SET TAX.SCF13_&VARNAME.3;
	CATTYPE = "&VARNAME.";
	PCT = VALUE / TOTINC;
	PCT2 = ABSVALUE / ABSINC;
RUN;
%MEND;
%INCTYPE(INCCAT);
%INCTYPE(NWCAT);
%INCTYPE(ASSCAT);

DATA TAX.TABLEAU;
	SET TAX.SCF13_FINAL:;
RUN;

PROC DATASETS LIB=TAX NOLIST;
	DELETE SCF13_:  SCF13B SCF13C;
RUN; QUIT;

DATA _NULL_;
	SET TAX.TABLEAU;
	FORMAT FREQ POP VALUE ABSVALUE TOTINC ABSINC 20.;
	FILE "&TAX.SCF13_INCOME_DISTRIBUTION.CSV" DSD DLM=",";
	IF _n_ = 1 THEN PUT "CATTYPE,CAT,FREQ,POP,INCTYPE,VALUE,ABSVALUE,TOTINC,ABSINC,PCT,PCT2";
	PUT CATTYPE CAT FREQ POP INCTYPE VALUE ABSVALUE TOTINC ABSINC PCT PCT2;
RUN;

/* IRS SOI */;
LIBNAME TABLE1_4 PCFILES PATH="&TAX.TAB_1.4.XLS" ACCESS=READONLY;
DATA TAX.IRS_SOI_TAB14;
	SET TABLE1_4.'TAB14$'n;
RUN;
DATA TAX.TAB14_ITEMS;
	SET TABLE1_4.'ITEMS$'n;
RUN;
LIBNAME TABLE1_4 CLEAR;

DATA TAX.IRS_SOI_TAB14_2;
	SET TAX.IRS_SOI_TAB14;
	ARRAY ITEM(*) I_003-I_136;
	DO i = 1 TO DIM(ITEM) BY 2;
		ITEM(i) = .;
	END;
	DROP i ;
RUN;

PROC TRANSPOSE DATA=TAX.IRS_SOI_TAB14_2 OUT=TAX.IRS_SOI_TAB14_3(WHERE=(COL1 NE .) DROP=_LABEL_);
	BY LOW LABEL;
RUN;

DATA TAX.IRS_SOI_TAB14_4;
	SET TAX.IRS_SOI_TAB14_3(RENAME=(COL1 = VALUE LABEL=LOW_LABEL));
	ITEM = INPUT(SUBSTR(_NAME_,3,3),3.);
	DROP _NAME_;
RUN;

DATA _NULL_;
	SET TAX.TAB14_ITEMS END=EOF;
	IF _n_ = 1 THEN CALL EXECUTE("
		DATA TAX.IRS_SOI_TAB14_5(WHERE=(ITEM1040 NE """" OR ITEM = 1));
			SET TAX.IRS_SOI_TAB14_4;
			FORMAT ITEM_LABEL $100. ITEM1040 $3.;
	");
	CALL EXECUTE("
			IF ITEM = "||COMPRESS(ITEMNO)||" THEN DO;
				ITEM_LABEL = """||TRIM(LABEL)||""";
				ITEM1040 = """||COMPRESS(ITEM1040)||""";
				IRS_CONTRIB = "||ORIG||";
				INC_CONTRIB = "||NEW||";
			END;
	");
	IF EOF THEN CALL EXECUTE("RUN;");
RUN;

DATA TAX.IRS_SOI_TAB14_6;
	SET TAX.IRS_SOI_TAB14_5;
	FORMAT INCTYPE $20. LOW2_LABEL $20.;
	IF ITEM IN(6) THEN INCTYPE = "WAGES";
		ELSE IF ITEM IN(1) THEN INCTYPE = "TAX_RETURNS";
		ELSE IF ITEM IN(52,56,60) THEN INCTYPE = "POS_RENT_TRUST";
		ELSE IF ITEM IN(54,58,62) THEN INCTYPE = "NEG_RENT_TRUST";
		ELSE IF ITEM IN(24,26) THEN INCTYPE = "POS_GAINS";
		ELSE IF ITEM IN(28) THEN INCTYPE = "NEG_GAINS";
		ELSE IF ITEM IN(36,70) THEN INCTYPE = "RETIREMENT";
		ELSE IF ITEM IN(12) THEN INCTYPE = "DIVIDEND";
		ELSE IF ITEM IN(0) THEN INCTYPE = "ASSISTANCE";
		ELSE IF ITEM IN(8,10,16,18,20,22,64,66,68,74,76,78,80,82,84,86) THEN INCTYPE = "OTHER_INC";

	IF INCTYPE = "OTHER_INC" & INC_CONTRIB = 1 THEN INCTYPE = "POS_OTHER";
		ELSE IF INCTYPE = "OTHER_INC" & INC_CONTRIB = -1 THEN INCTYPE = "NEG_OTHER";
	IF INCTYPE NE "";
	NEWVAL = VALUE*INC_CONTRIB;

	*NEW INCOME RANGES;
	IF LOW = -1 THEN DO;
		LOW2 = -1; LOW2_LABEL = "0 or Less";
	END;
	IF LOW IN(1,5000) THEN DO;
		LOW2 = 1; LOW2_LABEL = "1 to <10K";
	END;
	IF LOW IN(10000,15000) THEN DO;
		LOW2 = 10000; LOW2_LABEL = "10K to <20K";
	END;
	IF LOW IN(20000,25000) THEN DO;
		LOW2 = 20000; LOW2_LABEL = "20K to <30K";
	END;
	IF LOW IN(30000) THEN DO;
		LOW2 = 30000; LOW2_LABEL = "30K to <40K";
	END;
	IF LOW IN(40000) THEN DO;
		LOW2 = 40000; LOW2_LABEL = "40K to <50K";
	END;
	IF LOW IN(50000) THEN DO;
		LOW2 = 50000; LOW2_LABEL = "50K to <75K";
	END;
	IF LOW IN(75000) THEN DO;
		LOW2 = 75000; LOW2_LABEL = "75K to <100K";
	END;
	IF LOW IN(100000,200000) THEN DO;
		LOW2 = 100000; LOW2_LABEL = "100K to <250K";
	END;
	IF LOW IN(250000) THEN DO;
		LOW2 = 250000; LOW2_LABEL = "250K to <500K";
	END;
	IF LOW IN(500000) THEN DO;
		LOW2 = 500000; LOW2_LABEL = "500K to <1M";
	END;
	IF LOW IN(1000000,1500000) THEN DO;
		LOW2 = 1000000; LOW2_LABEL = "1M to <2M";
	END;
	IF LOW IN(2000000) THEN DO;
		LOW2 = 2000000; LOW2_LABEL = "2M to <5M";
	END;
	IF LOW IN(5000000) THEN DO;
		LOW2 = 5000000; LOW2_LABEL = "5M to <10M";
	END;
	IF LOW IN(10000000) THEN DO;
		LOW2 = 10000000; LOW2_LABEL = "10M and Up";
	END;
RUN;

PROC SQL;
	CREATE TABLE TAX.IRS_TABLEAU AS
	SELECT
		A.LOW, A.LABEL, A.INCTYPE, B.RETURNS, A.VALUE,
		B.TOTINC, B.POSINC
	FROM (
		SELECT
			LOW2 AS LOW,
			MIN(LOW2_LABEL) AS LABEL,
			INCTYPE,
			SUM(NEWVAL) AS VALUE
		FROM TAX.IRS_SOI_TAB14_6(WHERE=(ITEM NE 1))
		GROUP BY LOW2, INCTYPE) AS A
	INNER JOIN (
		SELECT
			A.LOW, B.RETURNS, A.TOTINC, A.POSINC
		FROM
			(SELECT
				LOW2 AS LOW,
				SUM(NEWVAL) AS TOTINC,
				SUM(MAX(NEWVAL,0)) AS POSINC
			FROM TAX.IRS_SOI_TAB14_6(WHERE=(ITEM NE 1))
			GROUP BY LOW2) AS A
		INNER JOIN
			(SELECT
				LOW2 AS LOW,
				SUM(DISTINCT VALUE) AS RETURNS
			FROM TAX.IRS_SOI_TAB14_6(WHERE=(ITEM = 1))
			GROUP BY LOW2) AS B
		ON A.LOW = B.LOW) AS B
	ON A.LOW = B.LOW;
QUIT;

PROC DATASETS LIB=TAX NOLIST;
	DELETE IRS_SOI_:;
RUN; QUIT;

DATA _NULL_;
	SET TAX.IRS_TABLEAU;
	FORMAT VALUE TOTINC POSINC 20.;
	FILE "&TAX.IRS2012_INCOME_DISTRIBUTION.CSV" DSD DLM=",";
	IF _n_ = 1 THEN PUT "LOW,LABEL,INCTYPE,RETURNS,VALUE,TOTINC,POSINC";
	PUT LOW LABEL INCTYPE RETURNS VALUE TOTINC POSINC;
RUN;