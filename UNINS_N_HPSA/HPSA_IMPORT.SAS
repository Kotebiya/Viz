%LET SOURCE=/Directory/;
LIBNAME HPSA "&SOURCE.";

%MACRO HPSA_EXCEL(TYPE);
PROC IMPORT DATAFILE="&SOURCE.BCD_HPSA_FCT_DET_&TYPE..xlsx" OUT=&TYPE. DBMS=XLSX REPLACE;
	GETNAMES=NO;
	SHEET=BCD_HPSA_FCT_DET;
	DATAROW=2;
RUN;
%MEND;

%HPSA_EXCEL(PC);
%HPSA_EXCEL(DH);
%HPSA_EXCEL(MH);

PROC IMPORT DATAFILE="&SOURCE.MUA_DET.xlsx" OUT=MUA DBMS=XLSX REPLACE;
	GETNAMES=NO;
	SHEET=MUA_DET;
	DATAROW=2;
RUN;

ODS TRACE ON;
PROC CONTENTS DATA=PC VARNUM;
RUN;
ODS TRACE OFF;

PROC CONTENTS DATA=PC VARNUM;
ODS SELECT Position;
RUN;
PROC CONTENTS DATA=DH VARNUM;
ODS SELECT Position;
RUN;
PROC CONTENTS DATA=MH VARNUM;
ODS SELECT Position;
RUN;

%MACRO HPSA_CLEAN(TYPE);
DATA HPSA.&TYPE.;
	SET &TYPE.(RENAME=(
		B = HPSA_SOURCE_ID
		C = HPSA_RECORD_TYP
		D = HPSA_SOURCE_NM
		E = HPSA_STATUS_CD
		F = HPSA_STATUS_DESC
		G = HPSA_TYP_CD
		H = HPSA_TYP_DESC
		I = HPSA_ADDRESS
		J = HPSA_CITY
		K = HPSA_STATE_ABBR
		L = HPSA_ZIP_CD
		M = HPSA_DEGREE_OF_SHORTAGE
		N = HPSA_DESIGNATION_DT
		O = HPSA_DESIGNATION_LAST_UPDATE_
		P = HPSA_DESIGNATION_POP
		Q = HPSA_ESTIMATED_UNDERSERVED_P
		R = HPSA_ESTIMATED_SERVED_POP
		S = HPSA_FORMAL_RATIO
		T = HPSA_FTE
		U = HPSA_METRO_IND_CD
		V = HPSA_METRO_IND_DESC
		W = HPSA_POP_PROVIDER_RATIO_GOAL
		X = HPSA_POVERTY
		Y = HPSA_SCORE
		Z = HPSA_SERVICE_AREA
		AA = HPSA_SHORTAGE
		AB = DISCIPLINE_CLASS_NUM
		AC = DISCIPLINE_CLASS_DESC
		AD = COMP_SOURCE_ID
		AE = COMP_SOURCE_NM
		AF = COMP_STATUS_CD
		AG = COMP_STATUS_DESC
		AH = COMP_TYP_CD
		AI = COMP_TYP_DESC
		AJ = COMP_STATE_ABBR
		AK = COMP_ZIP_CD
		AL = COMP_DESIGNATION_DT
		AM = COMP_DESIGNATION_DT_TXT
		AN = COMP_DESIGNATION_LAST_UPDATE_
		AO = GEO_ID
		AP = COUNTY_FIPS_CD
		AQ = COUNTY_NM
		AR = STATE_COUNTY_FIPS_CD
		AS = STATE_FIPS_CD
		AT = STATE_ABBR
		AU = STATE_NM
		AV = PRIMARY_STATE_NM
		AW = PRIMARY_STATE_FIPS_CD
		AX = PRIMARY_REGION_NM
		AY = US_MEXICO_BORDER_COUNTY_IND
		AZ = US_MEXICO_BORDER_100KM_IND
		BA = DW_RECORD_CREATE_DT
		BB = DW_RECORD_CREATE_DT_TXT
		BC = HPSA_NM
		BD = COMP_NM
		BE = BREAK_DESIGNATION_IND
		BF = X
		BG = Y
		BH = HPSA_POPULATION_TYP_CD
		BI = HPSA_POPULATION_TYP_DESC
		BJ = HPSA_RES_CIV_POP
		BK = CMN_CITY_NM
		BL = CMN_ZIP_CD
		BM = CMN_COUNTY_NM_STATE_ABBR
		BN = CMN_STATE_COUNTY_FIPS_CD
		BO = CMN_STATE_ABBR
		BP = CMN_STATE_NM
		BQ = CMN_STATE_FIPS_CD
		BR = CMN_REGION_NM
		BS = HPSA_WITHDRAWAL_DT
		BT = HPSA_WITHDRAWAL_DT_TXT
		BU = PROVIDER_TYP_CD_DESC
		BV = SURROUNDING_HPSA_ID
		BW = SURROUNDING_HPSA_NM
		BX = SURROUNDING_HPSA_SCORE
		BY = SURROUNDING_HPSA_SOURCE_ID));
	LABEL	HPSA_SOURCE_ID = "HPSA Source Identification Number; HPSA ID; HPSA Source ID"
			HPSA_RECORD_TYP = "HPSA Record Type"
			HPSA_SOURCE_NM = "HPSA Source Name; HPSA Name"
			HPSA_STATUS_CD = "HPSA Status Code"
			HPSA_STATUS_DESC = "HPSA Status Description; HPSA Status"
			HPSA_TYP_CD = "HPSA Type Code; HPSA Designation Type Code"
			HPSA_TYP_DESC = "HPSA Type Description; Designation Type; HPSA Designation Type"
			HPSA_ADDRESS = "HPSA Address"
			HPSA_CITY = "HPSA City; City; Health Professional Shortage Area City"
			HPSA_STATE_ABBR = "HPSA State Abbreviation; State Abbreviation"
			HPSA_ZIP_CD = "HPSA Postal Code; HPSA Facility Postal Code; HPSA ZIP Code"
			HPSA_DEGREE_OF_SHORTAGE = "HPSA Degree of Shortage; Degree of Shortage"
			HPSA_DESIGNATION_DT = "HPSA Designation Date; Designation Date"
			HPSA_DESIGNATION_LAST_UPDATE_ = "HPSA Designation Last Update Date; Date HPSA Last Updated"
			HPSA_DESIGNATION_POP = "HPSA Designation Population; Designation Population"
			HPSA_ESTIMATED_UNDERSERVED_P = "HPSA Estimated Underserved Population; Estimated Underserved"
			HPSA_ESTIMATED_SERVED_POP = "HPSA Estimated Served Population"
			HPSA_FORMAL_RATIO = "HPSA Formal Ratio; Ratio"
			HPSA_FTE = "HPSA Total Full-Time Equivalent Clinicians; FTE; HPSA FTE; Total FTE"
			HPSA_METRO_IND_CD = "HPSA Metropolitan Indicator Code; Metropolitan Indicator Code"
			HPSA_METRO_IND_DESC = "HPSA Metropolitan Indicator Description; Metropolitan Indicator"
			HPSA_POP_PROVIDER_RATIO_GOAL = "HPSA Provider Ratio Goal; Provider Ratio Goal"
			HPSA_POVERTY = "HPSA Percent of Population Below Poverty Level; % of Population Below"
			HPSA_SCORE = "HPSA Score"
			HPSA_SERVICE_AREA = "HPSA Service Area; Service Area"
			HPSA_SHORTAGE = "HPSA Shortage; # of FTEs Short"
			DISCIPLINE_CLASS_NUM = "Discipline Class Number; HPSA Discipline Class Number"
			DISCIPLINE_CLASS_DESC = "Discipline Class Description; Discipline; Discipline Class; Discipline Group;"
			COMP_SOURCE_ID = "HPSA Component Source Identification Number; Component ID; HPSA"
			COMP_SOURCE_NM = "HPSA Component Source Name; HPSA Component Name"
			COMP_STATUS_CD = "HPSA Component Status Code; Component Status Code"
			COMP_STATUS_DESC = "HPSA Component Status Description; Component Status"
			COMP_TYP_CD = "HPSA Component Type Code; Component Type Code"
			COMP_TYP_DESC = "HPSA Component Type Description; Component Type"
			COMP_STATE_ABBR = "HPSA Component State Abbreviation"
			COMP_ZIP_CD = "HPSA Component Postal Code; HPSA Component ZIP Code"
			COMP_DESIGNATION_DT = "HPSA Component Designation Date; Component Designation Date"
			COMP_DESIGNATION_DT_TXT = "HPSA Component Designation Date String; HPSA Component Designation"
			COMP_DESIGNATION_LAST_UPDATE_ = "HPSA Component Designation Last Update Date; Date Component Last"
			GEO_ID = "HPSA Geography Identification Number; Geography ID; GEOID; HPSA"
			COUNTY_FIPS_CD = "County FIPS; County FIPS Code; County or County Equivalent Federal"
			COUNTY_NM = "County Equivalent Name; County; County Equivalent; County Name;"
			STATE_COUNTY_FIPS_CD = "State County FIPS Code; State and County Federal Information Processing"
			STATE_FIPS_CD = "State FIPS Code; State Federal Information Processing Standard Code;"
			STATE_ABBR = "State Abbreviation; State Code; State_Cd"
			STATE_NM = "State Name; Component State; Site State; State; State Label"
			PRIMARY_STATE_NM = "Primary State Name; HPSA Primary State Name; MUA/P Primary State;"
			PRIMARY_STATE_FIPS_CD = "Primary State FIPS Code; HPSA Primary State FIPS Code; MUA/P Primary"
			PRIMARY_REGION_NM = "Primary HHS Region Name; HHS Region; HPSA Primary Region Name;"
			US_MEXICO_BORDER_COUNTY_IND = "U.S. - Mexico Border County Indicator; Primary Care Service Area U.S. -"
			US_MEXICO_BORDER_100KM_IND = "U.S. - Mexico Border 100 Kilometer Indicator; La Paz Agreement; Primary"
			DW_RECORD_CREATE_DT = "Data Warehouse Record Create Date; SNAP; Snapshot; Snapshot Date"
			DW_RECORD_CREATE_DT_TXT = "Data Warehouse Record Create Date Text; SNAP; Snapshot Date;"
			HPSA_NM = "HPSA Name"
			COMP_NM = "HPSA Component Name; Component Name"
			BREAK_DESIGNATION_IND = "Break in Designation; Break Designation"
			X = "Geocoding Artifact Address Primary X Coordinate; Address Primary X"
			Y = "Geocoding Artifact Address Primary Y Coordinate; Address Primary Y"
			HPSA_POPULATION_TYP_CD = "HPSA Population Type Code; Population Type Code"
			HPSA_POPULATION_TYP_DESC = "HPSA Population Type Description; HPSA Population Type"
			HPSA_RES_CIV_POP = "HPSA Resident Civilian Population; Resident Civilian Population"
			CMN_CITY_NM = "Common City Name with State Abbreviation; Common City Name"
			CMN_ZIP_CD = "Common ZIP Code; Common Postal Code; ZIP Code"
			CMN_COUNTY_NM_STATE_ABBR = "Common County Name with State Abbreviation; Common County Name;"
			CMN_STATE_COUNTY_FIPS_CD = "Common State County FIPS Code; State County FIPS Code"
			CMN_STATE_ABBR = "Common State Abbreviation; Common State Name; State Abbreviation"
			CMN_STATE_NM = "Common State Name; State; State Name"
			CMN_STATE_FIPS_CD = "Common State FIPS Code; State FIPS Code"
			CMN_REGION_NM = "Common Region Name; Common HHS Region Name; HHS Region; HHS"
			HPSA_WITHDRAWAL_DT = "HPSA Withdrawn Date; Withdrawn Date"
			HPSA_WITHDRAWAL_DT_TXT = "HPSA Withdrawn Date; Withdrawn Date"
			PROVIDER_TYP_CD_DESC = "Provider Type"
			SURROUNDING_HPSA_ID = "Surrounding HPSA Source Surrogate Key"
			SURROUNDING_HPSA_NM = "Surrounding HPSA Name"
			SURROUNDING_HPSA_SCORE = "Surrounding HPSA Score"
			SURROUNDING_HPSA_SOURCE_ID = "Surrounding HPSA Source Identification Number";
RUN;
%MEND;

%HPSA_CLEAN(PC);
%HPSA_CLEAN(DH);
%HPSA_CLEAN(MH);

%LET KEEPVAR=HPSA_SOURCE_ID HPSA_TYP_CD COMP_TYP_DESC
			HPSA_STATUS_CD HPSA_STATUS_DESC
			/* Geographical ID's */
			GEO_ID STATE_FIPS_CD STATE_ABBR STATE_NM COUNTY_FIPS_CD COUNTY_NM
			DISCIPLINE_CLASS_NUM DISCIPLINE_CLASS_DESC
			HPSA_DESIGNATION_POP HPSA_ESTIMATED_UNDERSERVED_P HPSA_ESTIMATED_SERVED_POP HPSA_FORMAL_RATIO HPSA_FTE
			HPSA_POP_PROVIDER_RATIO_GOAL;

DATA HPSA.JOIN;
	RETAIN	&KEEPVAR.;
	SET	HPSA.PC
		HPSA.DH
		HPSA.MH;
	KEEP	&KEEPVAR.;
	WHERE	STATE_FIPS_CD = "06" & COUNTY_FIPS_CD IN("037","059")
		 	& HPSA_STATUS_CD = "D" & HPSA_TYP_CD = "Hpsa Pop" & COMP_TYP_DESC = "Census Tract"
		 	& DISCIPLINE_CLASS_NUM = 1
		 	;
RUN;

DATA HPSA.HPSA;
	RETAIN GEO_ID HPSA_SOURCE_ID HPSA_DESIGNATION_POP HPSA_FTE HPSA_FORMAL_RATIO;
	SET HPSA.JOIN;
	KEEP GEO_ID HPSA_SOURCE_ID HPSA_DESIGNATION_POP HPSA_FTE HPSA_FORMAL_RATIO RATIO;
	RENAME	GEO_ID = GEOID
			HPSA_SOURCE_ID = HPSA_ID
			HPSA_DESIGNATION_POP = POP
			HPSA_FTE = FTE;
	ATTRIB RATIO FORMAT=11.2;
	RATIO = HPSA_DESIGNATION_POP / HPSA_FTE;
RUN;
