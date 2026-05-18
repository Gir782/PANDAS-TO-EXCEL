import pandas as pd
from datetime import *

FILE_PATH=input("ENTER YOUR FILE PATH: ").strip().replace('"', '')
FILE_EXPORT=rf"C:\Users\GIRISH\Desktop"
source=pd.read_excel(rf"{FILE_PATH}",sheet_name="FGET",header=None,engine="pyxlsb")
MODE=source.iloc[0,1]
PORT=source.iloc[1,1]
BETypeCode=source.iloc[2,1]
IMPORTER=source.iloc[3,1]
ADCODE=source.iloc[4,1]
Importer_RefNo=source.iloc[5,1]
CountryOfOriginCode=source.iloc[6,1]
PortOfShipmentCode=source.iloc[7,1]
CountryOfShipmentCode=source.iloc[8,1]
BE_HEADING=source.loc[9,1]
TRANSACTION_DEFERRED=source.iloc[10,1]
PRIOR_NORMAL=source.iloc[11,1]
IsTranshipment=source.iloc[12,1]


GENERAL=pd.DataFrame(columns=['TransportModeCode',
'CustomsHouseCode',
'BETypeCode',
'Importer',
'Branch Name',
'AD_Code',
'Importer_RefNo',
'CountryOfOriginCode',
'PortOfShipmentCode',
'CountryOfShipmentCode',
'BE-Heading',
'DutyPaymentStatus_T_D',
'AdvancePriorNormal',
'IsUnderSec46',
'IsUnderSec48',
'IsFirstCheck',
'IsGreenChannel',
'IsKachchaBE',
'IsHSS',
'IsBondsCertificates',
'IsTranshipment',
'ITC_Lic_details'
])



GENERAL.loc[0,"TransportModeCode"]=MODE
GENERAL.loc[0,"CustomsHouseCode"]=PORT
GENERAL.loc[0,"BETypeCode"]=BETypeCode
GENERAL.loc[0,"Importer"]=IMPORTER
GENERAL.loc[0,"AD_Code"]=ADCODE
GENERAL.loc[0,'Importer_RefNo']=Importer_RefNo
GENERAL.loc[0,'CountryOfOriginCode']=CountryOfOriginCode
GENERAL.loc[0,'PortOfShipmentCode']=PortOfShipmentCode
GENERAL.loc[0,'CountryOfShipmentCode']=CountryOfShipmentCode
GENERAL.loc[0,'BE-Heading']=BE_HEADING
GENERAL.loc[0,'DutyPaymentStatus_T_D']=TRANSACTION_DEFERRED
GENERAL.loc[0,'AdvancePriorNormal']=PRIOR_NORMAL
GENERAL.loc[0,'IsTranshipment']=IsTranshipment

GENERAL_HIDE_COLUMNS=["Branch Name","IsUnderSec46","IsUnderSec48","IsFirstCheck","IsGreenChannel","IsKachchaBE","IsHSS","IsBondsCertificates","ITC_Lic_details"]




inbond_exbond=pd.DataFrame(columns=['WH_Code',
'WH_Name',
'WH_Add1',
'WH_Add2',
'WH_City',
'WH_PIN',
'WH_Country',
'InBond_BENo',
'InBond_BEDate',
'Bond_No',
'Bond_Date',
'Bond_ExpiryDate',
'IsWareHouseSale',
'IsSEC65ManufacturingWH',
])

# SHIPMENT WORK
#path "C:\Users\girish\Desktop\MASTER_FILE\TESTING FOR EXCEL.xlsb"
Source2=pd.read_excel(rf"{FILE_PATH}",sheet_name="FGET",header=None,engine="pyxlsb")


INWARD_DATE=Source2.iloc[0,4]
FlightNo_VoyageNo=Source2.iloc[1,4]
IGM_No=Source2.iloc[2,4]
IGM_Date=Source2.iloc[3,4]
MAWB_MBL_No=Source2.iloc[4,4]

MAWB_BL_Date=pd.to_datetime(Source2.iloc[5,4],dayfirst=True ,errors='coerce').date()


HAWB_HBL_No=Source2.iloc[6,4]
HAWB_HBL_Date=pd.to_datetime(Source2.iloc[7,4],dayfirst=True ,errors='coerce').date()
PKG=Source2.iloc[8,4]
PkgUnitCode=Source2.iloc[9,4]
Gross_weight=Source2.iloc[10,4]
GrWtUnitCode=Source2.iloc[11,4]

COMBINE_marks_n_nos=f"{Source2.iloc[4,4]}\n................\n{Source2.iloc[6,4]}"

SHIPMENT=pd.DataFrame(columns=["VesselName",
"Flight_Inward_Date",
"FlightNo_VoyageNo",
"LineNo",
"IGM_No",
"IGM_Date",
"MAWB_MBL_No",
"AWB_BL_Date",
"HAWB_HBL_No",
"HAWB_HBL_Date",
"No_of_Pkg",
"PkgUnitCode",
"GrWt",
"GrWtUnitCode",
"NtWt",
"NtWtUnitCode",
"Marks_&_Nos",
"Port_of_Reporting",
"Gateway_IGM_No",
"Gateway_IGM_Date",
"Gateway_Inward_Date",
"ChgWt",
"ChgWtUnitCode"
])




SHIPMENT.loc[0,"Flight_Inward_Date"]=INWARD_DATE
SHIPMENT.loc[0,"FlightNo_VoyageNo"]=FlightNo_VoyageNo
SHIPMENT.loc[0,"IGM_No"]=IGM_No

SHIPMENT.loc[0,"IGM_Date"]=IGM_Date
SHIPMENT.loc[0,"MAWB_MBL_No"]=MAWB_MBL_No
SHIPMENT.loc[0,"AWB_BL_Date"]=MAWB_BL_Date
SHIPMENT.loc[0,"HAWB_HBL_No"]=HAWB_HBL_No
SHIPMENT.loc[0,"HAWB_HBL_Date"]=HAWB_HBL_Date
SHIPMENT.loc[0,"No_of_Pkg"]=PKG
SHIPMENT.loc[0,"PkgUnitCode"]=PkgUnitCode
SHIPMENT.loc[0,"GrWt"]=Gross_weight
SHIPMENT.loc[0,"GrWtUnitCode"]=GrWtUnitCode
SHIPMENT.loc[0,"Marks_&_Nos"]=COMBINE_marks_n_nos
SHIPMENT_TAB_HIDE=["VesselName","LineNo","NtWt","NtWtUnitCode",'Port_of_Reporting','Gateway_IGM_No','Gateway_IGM_Date','Gateway_Inward_Date']


CONTAINERS=pd.DataFrame(columns=["IGM Sr.No",
"Container No",
"Seal No",
"FCL_LCL",
"ContainerTypeCode",
"ContainerSize",
"Truck Number",
"EmptyContainerLocation",
"PackagesStuffed",
"GrWt"
])
#paste here "C:\Users\girish\Desktop\MASTER_FILE\TESTING FOR EXCEL.xlsb"
source3=pd.read_excel(rf"{FILE_PATH}",sheet_name="FGET",header=None,engine="pyxlsb")


# here
file=rf"{FILE_PATH}"
source3=pd.read_excel(file,sheet_name="MAIN_DATA",header=8,usecols=["SR_NO.","INVOICE_NO.","Date","T. AMOUNT"],engine="pyxlsb")
source3.drop_duplicates(subset=["INVOICE_NO.","Date"],keep="first",inplace=True)
inv=source3.reset_index(drop=True)
finv=inv.loc[:,"INVOICE_NO."]
fdate=inv.loc[:,"Date"]
inv_sr=pd.Series([i for i in range(1,len(finv)-1)])
#
source4=pd.read_excel(file,sheet_name="MAIN_DATA",header=8,usecols=["INVOICE_NO.","Date"],engine="pyxlsb")
source4_1=source4.drop_duplicates(subset="INVOICE_NO.",keep="first")
source4_1.reset_index(drop=True,inplace=True)
#ADDED CHT
source4_1.loc[:, "Date"] = pd.to_datetime(source4_1["Date"], dayfirst=True, errors="coerce").dt.date

#ADDED CHT
source4_2=source4_1["Date"]

FOB_CIF=Source2.iloc[13,4]
TERMS=pd.Series([FOB_CIF for i in range(len(inv_sr))])

# for inv currency 
inv_currency=pd.read_excel(file,sheet_name="MAIN_DATA",engine="pyxlsb",header=None,nrows=10)
currency=inv_currency.iloc[2,1]

# FOR EACH INVOICE AMOUNT
source_amount=pd.read_excel(file,sheet_name="MAIN_DATA",header=8,usecols=["INVOICE_NO.","Date","T. AMOUNT","UQC"],engine="pyxlsb")

amount=source_amount.dropna(subset=["T. AMOUNT"])
inv_value=amount.groupby(by="INVOICE_NO.",sort=False)["T. AMOUNT"].agg("sum").reset_index()
inv_amount=inv_value["T. AMOUNT"]




# source=pd.read_excel(rf"{FILE_PATH}",sheet_name="FGET",header=None,engine="pyxlsb")
S_FREIGHT=""
if source.iloc[13,4]=="FOB":
    S_FREIGHT="Y"
else:
    S_FREIGHT="N"

if S_FREIGHT=="Y":
    FREIGHT_AMT=input("ENTER THE FREIGHT AMOUNT:-")
    FREIGHT_CURY=input("ENTER THE FREIGHT CURRENCY:-")
    insurance_cury=input("is insurance currency different from inv currency\nY FOR YES N FOR NO:- ")
    if insurance_cury=="N":
        insur_percentage=pd.read_excel(file,sheet_name="MAIN_DATA",header=27,engine="pyxlsb",nrows=1)
        insur_percentage=insur_percentage.iloc[0,0]
        insur_curreny1=pd.read_excel(file,sheet_name="MAIN_DATA",header=1,engine="pyxlsb",nrows=1)
        insur_curreny1=insur_curreny1.iloc[0,1]
    elif insurance_cury=="Y":
        insurance_amount=input("enter the amount")
        insur_curreny2=input("enter the currency")

    else:
        pass
else:
    pass

CSV_SOURCE=pd.read_excel(file,sheet_name="CSV")
Suplier_name=CSV_SOURCE["Consigner Name"]
Suplier_add=CSV_SOURCE["Consigner Address"]
suplier_country=CSV_SOURCE["Consigner Country"]

relation=input("Is consignee and suppler Party related\nY FOR YES N FOR NO:- ").upper()


Is_Related=""
related=""
Base=""
Condition=""
SVB_Ref_No=""
Custom_House_Code=""
SVB_Loading_Basis=""
SVB_Status_Assessable=""
SVB_Status_Duty=""
SVB_Date=""


if relation=="Y":
    Is_Related="Y"
    related="RELATED"
    Base1=pd.read_excel(file,sheet_name="MAIN_DATA",header=25,engine="pyxlsb",nrows=1)
    Base1=Base1.iloc[0,0]
    Base=Base1
    Condition="N/A"
    SVB_Ref_No,SVB_Date=Base.split("dt.")
    SVB_Ref_No=SVB_Ref_No.strip()
    SVB_Date=SVB_Date.strip()
    SVB_Date = pd.to_datetime(SVB_Date, dayfirst=True, errors="coerce").date()
    Custom_House_Code="INDEL4"
    SVB_Loading_Basis="A"
    SVB_Status_Assessable="F"
    SVB_Status_Duty="F"
else:
    pass



# INVOICE TAB
INVOICES=pd.DataFrame(columns=["InvSrNo",
"Invoice_No",
"Invoice_Date",
"TOI",
"TOI_Place",
"Inv_Currency",
"Product_Value",
"Is_Single_Frt_Ins_Other_Chrg",
"Frt_%",
"Frt_Amount",
"Frt_Currency",
"Ins_%",
"Ins_Amount",
"Ins_Currency",
"Misc_Charge_%",
"Misc_Charge_Amount",
"Misc_Charge_Currency",
"Agency_%",
"Agency_Amount",
"Agency_Currency",
"Discount_%",
"Discount_Amount",
"Discount_Currency",
"Loading_%",
"Loading_Amount",
"Loading_Currency",
"HSS_%",
"HSS_Amount",
"RD_%",
"RD_Basis",
"Supplier_Name",
"Supplier_Address",
"Supplier_City",
"Supplier_Country_Code",
"Is_Related",
"Relation",
"Base",
"Condition",
"SVB_Ref_No",
"SVB_Date",
"Custom_House_Code",
"SVB_Loading_Basis",
"SVB_Rate_Assessable",
"SVB_Status_Assessable",
"SVB_Rate_Duty",
"SVB_Status_Duty",
"PO_No",
"PO_Date",
"Contract.No",
"Contract_Date",
"Terms_of_Payment",
"Other_Terms_of_Payment_Remark",
"LC_No",
"LC_Date",
"Nature_of_Trans",
"Valuation_Method",
"Sale_Condition",
"Other_relevant_info",
"AEO_Code",
"AEO_Country",
"AEO_Rule",
"A - Brokerage and Commissions_%",
"A - Brokerage and Commissions_Amount",
"B - Cost of Container_%",
"B - Cost of Container_Amount",
"C - Cost of Packing_%",
"C - Cost of Packing_Amount",
"D - Handling Charges_%",
"D - Handling Charges_Amount",
"E - Cost of Goods and Services_%",
"E - Cost of Goods and Services_Amount",
"F - Documentation_%",
"F - Documentation_Amount",
"G - Country of origin Certificate_%",
"G - Country of origin Certificate_Amount",
"H - Royalties and Licence Fees_%",
"H - Royalties and Licence Fees_Amount",
"I - Value of Proceeds Which Accrue_%",
"I - Value of Proceeds Which Accrue_Amount",
"J - Cost of Warranty Services_%",
"J - Cost of Warranty Services_Amount",
"K - Other Cost or Payment_%",
"K - Other Cost or Payment_Amount",
"L - Other Charges and Payment_%",
"L - Other Charges and Payment_Amount",
"M - Loading Charges_%",
"M - Loading Charges_Amount",
"N - Unloading Charges_%",
"N - Unloading Charges_Amount",
"Supplier_Branch"
])
## 



INVOICES["InvSrNo"]=inv_sr
INVOICES["Invoice_No"]=finv
INVOICES["Invoice_Date"]=source4_2
INVOICES["TOI"]=TERMS
INVOICES["Inv_Currency"]=currency
INVOICES["Product_Value"]=inv_amount
INVOICES["Is_Single_Frt_Ins_Other_Chrg"]=S_FREIGHT



if S_FREIGHT=="Y":
    INVOICES["Frt_Amount"]=FREIGHT_AMT
    INVOICES["Frt_Currency"]=FREIGHT_CURY
    if insurance_cury=="N":
        INVOICES["Ins_%"]=insur_percentage
        INVOICES["Ins_Currency"]=insur_curreny1
    elif insurance_cury=="Y":
        INVOICES["Ins_Amount"]=insurance_amount
        INVOICES["Ins_Currency"]=insur_curreny2
    else:
        pass
else:
    pass











INVOICES["Supplier_Name"]=Suplier_name
INVOICES["Supplier_Address"]=Suplier_add
INVOICES["Supplier_City"]="N/A"
INVOICES["Supplier_Country_Code"]=suplier_country
INVOICES["Is_Related"]=Is_Related
INVOICES["Relation"]=related
INVOICES["Base"]=Base
INVOICES["Condition"]=Condition
INVOICES["SVB_Ref_No"]=SVB_Ref_No
INVOICES["SVB_Date"]=SVB_Date
INVOICES["Custom_House_Code"]=Custom_House_Code
INVOICES["SVB_Loading_Basis"]=SVB_Loading_Basis
INVOICES["SVB_Status_Assessable"]=SVB_Status_Assessable
INVOICES["SVB_Status_Duty"]=SVB_Status_Duty
INVOICES["Terms_of_Payment"]="OTHERS"
INVOICES["Nature_of_Trans"]="Sale on Consignment basis"
INVOICES["Valuation_Method"]="RULE 4 (TRANSACTION VALUE)"
INVOICES["Supplier_Branch"]=""


#                                ITEM 
CSV_SOURCE=pd.read_excel(file,sheet_name="CSV",engine="pyxlsb")
item_inv_sr=CSV_SOURCE.groupby(by="Inv No",sort=False).ngroup()+1
inv_prod_sr=CSV_SOURCE.groupby(by="Inv No",sort=False).cumcount()+1



ITEMS=pd.DataFrame(columns=["InvSrNo",
"ItemSrNo",
"Inbond_InvSrNo",
"Inbond_ItemSrNo",
"Product_Description",
"QTY",
"Unit",
"Unit_Price",
"CTH",
"RITC",
"CETH",
"PolicyPara",
"PolicyYear",
"General_Description",
"Brand",
"Model",
"End_Use",
"Country_of_Origin",
"Accessories_Status",
"Accessories_Details",
"WH_SalePrice_INR",
"Exim_Code",
"Exim_Notn",
"Exim_NotnSrNo",
"Standard_Preferential",
"Basic_Notn",
"Basic_NotnSrNo",
"SWS_Notn",
"SWS_NotnSrNo",
"IGST_LevyNotn",
"IGST_LevyNotnSrNo",
"IGST_LevyNotnFlag",
"IGST_ExemptionNotnType",
"IGST_ExemptionNotn",
"IGST_ExemptionNotnSrNo",
"IGST_ExemptionNotnFlag",
"IGST_CompCessNotn",
"IGST_CompCessNotnSrNo",
"IGST_CompCessNotnFlag",
"IGST_CompCessExemptionNotnType",
"IGST_CompCessExemptionNotn",
"IGST_CompCessExemptionNotnSrNo",
"IGST_CompCessExemptionNotnFlag",
"Road_Infra_Cess_Notn",
"Road_Infra_Cess_NotnSrNo",
"NCD_Notn",
"NCD_NotnSrNo",
"Aggregate_Duty_Notn",
"Aggregate_Duty_NotnSrNo",
"Safeguard_Duty_Notn",
"Safeguard_Duty_NotnSrNo",
"SAPTA_Notn",
"SAPTA_NotnSrNo",
"Tariff_Value_Notn",
"Tariff_Value_NotnSrNo",
"Tarrif_Value_Qty",
"Tarrif_Value_Currency",
"Tarrif_Value_Amount",
"ADD_Notn",
"ADD_NotnSrNo",
"CTHSrNo",
"SuppSrNo",
"ADD_Qty",
"ADD_Basis",
"ADD_%Rate",
"ADD_Currency",
"ADD_AmountPerUnit",
"ADD_AmountUnit",
"Other_Duty_Notn",
"Other_Duty_NotnSrNo",
"Other_Duty_Flag",
"Other_Duty_%Rate",
"Other_Duty_AmountPerUnit",
"Other_Duty_AmountUnit",
"Duty_Type",
"Addl_Duty_Flag",
"MFG_Name",
"MFG_Address",
"MFG_Country",
"MFG_State",
"MFG_PIN",
"Source_Country",
"Transit_Country",
"SVBRefNo",
"SVBRefDate",
"SVBCustomHouse",
"SVB_Loading_Basis",
"SVB_Rate_Assessable",
"SVB_Status_Assessable",
"SVB_Rate_Duty",
"SVB_Status_Duty",
"DUTY_ExemptionType",
"CHealthCess_Notn",
"CHealthCess_NotnSrNo",
"isFTAbenefitClaimed",
"COO_No",
"COO_Date_of_Issue",
"COO_Issuing_Country",
"COO_Origin_Criteria",
"COO_Origin_Criteris_Remarks",
"COO_Accumulation_Cumulation",
"Foc_Item",
"AIDC_LevyNotn",
"AIDC_LevyNotnSrNo",
"AIDC_ExemptionNotn",
"AIDC_ExemptionNotnSrNo",
"AIDCNotn_Excise",
"ADICNotnSr_Excise",
"MaterialCode",
"Previous_BENo",
"Previous_BEDate",
"Previous_BEIGMNo",
"Previous_BEIGMDate",
"Previous_BECurrency",
"Previous_BEUnitPrice",
"Previous_BECustomHouse",
"COO_Retroactive_Issuance",
"COO_Direct_Consignment",
"COO_TariffShift",
"GST_Comp_Cess_SalePrice_INR",
"CVD_Notn",
"CVD_NotnSrNo",
"CVD_Rate",
"CVD_CalculatedOn",
"CVD_ItemSrNo",
"CVD_SupplierSrNo",
"COO_ItemSrNoCert"
])






ITEMS["InvSrNo"]=item_inv_sr
ITEMS["ItemSrNo"]=inv_prod_sr
ITEMS["Product_Description"]=CSV_SOURCE["Product Desc"]
ITEMS["QTY"]=CSV_SOURCE["Quantity"]
ITEMS["Unit"]=CSV_SOURCE["Unit"]
ITEMS["Unit_Price"]=CSV_SOURCE["Rate"]
ITEMS["CTH"]=CSV_SOURCE["CTH"]
ITEMS["RITC"]=CSV_SOURCE["CTH"]
ITEMS["CETH"]=CSV_SOURCE["CETH"]
ITEMS["General_Description"]=CSV_SOURCE["Generic Description"]
ITEMS["Brand"]=CSV_SOURCE["BRAND"].fillna("N/A")
ITEMS["Model"]=CSV_SOURCE["Model"].fillna("N/A")
ITEMS["End_Use"]=CSV_SOURCE["END USE"]
ITEMS["Country_of_Origin"]=CSV_SOURCE["Country of origin"]
ITEMS["Basic_Notn"]=CSV_SOURCE["Basic notification"]
ITEMS["Basic_NotnSrNo"] = CSV_SOURCE.iloc[:, CSV_SOURCE.columns.get_loc("Sr no")]
ITEMS["IGST_LevyNotn"]=CSV_SOURCE["IGST Notn No"]
ITEMS["IGST_LevyNotnSrNo"]=CSV_SOURCE["IGST SrNo"]


ITEMS["IGST_ExemptionNotnType"]="G"
ITEMS["IGST_CompCessExemptionNotnType"]="G"


ITEMS["MFG_Name"]=CSV_SOURCE["Manufacturer"]
ITEMS["MFG_Address"]=CSV_SOURCE["Address"]
ITEMS["Source_Country"]=CSV_SOURCE["COO_Issuing Country "]
ITEMS["Transit_Country"]=CSV_SOURCE["COO_Issuing Country "]


# FOR FTA BENEFITS
FTA_CEPA=["046/2011","151/2009","152/2009","069/2011"]

isFTAbenefitClaimed = CSV_SOURCE["Basic notification"].apply(
    lambda x: "Y" if x in FTA_CEPA else "")




ITEMS["isFTAbenefitClaimed"]=isFTAbenefitClaimed
ITEMS["COO_No"]=CSV_SOURCE["COO_No "]


# Convert COO_Date_of_Issue to datetime (day first) and remove time
ITEMS["COO_Date_of_Issue"] = pd.to_datetime(
    CSV_SOURCE["COO_Date of Issue "], 
    dayfirst=True, 
    errors="coerce"
).dt.normalize().dt.date


ITEMS["COO_Issuing_Country"]=CSV_SOURCE["COO_Issuing Country "]
ITEMS["COO_Origin_Criteria"]=CSV_SOURCE["COO_Origin Criteria "]
ITEMS["COO_Origin_Criteris_Remarks"]=CSV_SOURCE["COO_Origin Criteria remarks "]
ITEMS["COO_Accumulation_Cumulation"]=CSV_SOURCE["COO_Accumulation/Cumulation Status "]
#ITEM
ITEMS["Foc_Item"]=CSV_SOURCE["FOC Item"]
#HERE
ITEMS["AIDC_LevyNotn"]=CSV_SOURCE["AIDC Levy Notification No "]
ITEMS["AIDC_LevyNotnSrNo"]=CSV_SOURCE["AIDC Levy Notification Sr.No. "]
ITEMS["COO_Retroactive_Issuance"]=CSV_SOURCE["COO_Retroactive Issuance Status"]
ITEMS["COO_Direct_Consignment"]=CSV_SOURCE["COO_Direct Consignment Status"]
ITEMS["COO_TariffShift"]=CSV_SOURCE["COO_Tariff Shift"]
ITEMS["COO_ItemSrNoCert"]=CSV_SOURCE["COO_ItemSrNoCert"]

    
# IN STATEMENTS
STATEMENT=pd.DataFrame(columns=['Inv_SrNo',
'Item_SrNo',
'StatementType',
'StatementCode',
'StatementRemarks'
])

statement_inv_no=CSV_SOURCE.groupby(by="Inv No").ngroup()+1
statement_inv_no=statement_inv_no.drop_duplicates(keep="first").to_list()
rows=[]
statement_code=["CUV01", "CUV02", "CUV03", "CUV04"]

if relation=="Y":
    statement_code[-1]="CUV05"

for statements_invoice_sr in statement_inv_no:
    for code in statement_code:
        rows.append([statements_invoice_sr,0,"DEC",code])

DATAFRAME=pd.DataFrame(rows,columns=["inv_sr_nos","Item_SrNo","StatementType","Statementcode"])

STATEMENT["Inv_SrNo"]=DATAFRAME["inv_sr_nos"]
STATEMENT["Item_SrNo"]=DATAFRAME["Item_SrNo"]
STATEMENT["StatementType"]=DATAFRAME["StatementType"]
STATEMENT["StatementCode"]=DATAFRAME["Statementcode"]




SEC65_EXBOND_INFO=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"GSTInvoiceNo",
"GSTInvoiceDate",
"FinishedProductCTH",
"FinishedProductDesc",
"FinishedProductQty",
"FinishedProductQtyUnit"
])



REIMPORT=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"SB_No",
"SB_Date",
"Port_of_Export",
"SB_Inv_SrNo",
"SB_Item_SrNo",
"Notn_No",
"Notn_SrNo",
"Exp._Freight",
"Exp.Insurance",
"Cus.Duty",
"Excise_Duty",
"File_No"
])

LICENSE=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"License_RefNo",
"License_No",
"License_Date",
"License_RegNo",
"License_RegDate",
"Reg_Port",
"License_ItemSrNo",
"CIF_Value",
"DebitDeutyValue",
"DebitQuantity",
"DebitQuantityUnitCode"
])

# CPhere







SW_ADDL_INFO=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"Info_Type",
"info_Qualifier",
"Info_Code_Description",
"Information",
"Measure",
"Measure_Unit"
])



# THROUGTH OLD FORMULA
# SW_ADDL_INFO["Inv_SrNo"]=item_inv_sr
# SW_ADDL_INFO["Item_SrNo"]=inv_prod_sr
# SW_ADDL_INFO["Info_Type"]="CHR"
# SW_ADDL_INFO["info_Qualifier"]="SQC"
# SW_ADDL_INFO["Measure"]=CSV_SOURCE["SQC_Qty"]
# SW_ADDL_INFO["Measure_Unit"]=source_amount["UQC"]
# TILL HERE THROUGTH OLD FORMULA

# HERE TAKING DATA DIRECTLY FROM EXCEL

master_excel = pd.read_excel(FILE_PATH, engine="pyxlsb", sheet_name="SINGLE_WINDOW")

SW_ADDL_INFO[["Inv_SrNo","Item_SrNo","Info_Type","info_Qualifier","Info_Code_Description","Information","Measure","Measure_Unit"]]=master_excel[["Inv_SrNo","Item_SrNo","Info_Type","info_Qualifier","Info_Code_Description","Information","Measure","Measure_Unit"]]





# TILL HERE 

SW_CONSTITUENT=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"Cons_Name",
"Cons_Code",
"Cons_Percentage",
"Cons_Yield%",
"Cons_Active"
])
SW_PRODUCTION=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"Prod_Batch_ID",
"Prod_Batch_Quantity",
"Prod_Batch_Unit",
"Prod_Manufacturer_Date",
"Prod_Expiry_Date",
"Prod_Best_before_Date"
])

SW_CONTROL=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"Control_Type",
"Control_Location",
"Control_Start_Date",
"Control_End_Date",
"Control_Result"
])

SEZ_INFO=pd.DataFrame(columns=["InvSrNo",
"ItemSrNo",
"SEZ_Z_InvSrNo",
"SEZ_Z_ItemSrNo",
"BE_No",
"BE_Date",
"BE_Location",
"Code",
"QTY",
"Unit"
])
HSS=pd.DataFrame(columns=["Level",
"HSS_Name",
"HSS_BranchName",
"HSS_BranchSr",
"HSS_IECode",
"HSS_ADCode",
"HSS_Address",
"HSS_City",
"HSS_Country",
"HSS_PostalCode"
])

BONDS_CERTIFICATES=pd.DataFrame(columns=["Bond_or_Certificate",
"Bond_Cert_Type",
"Bond_Cert_No",
"Bond_Cert_Date",
"Commissionerate",
"Division",
"Range",
"Registration_Port_Code"
])

Bond_or_Certificate=""
Bond_Cert_Type=""
Bond_Cert_No=""
Bond_Cert_Date=""







if CSV_SOURCE["Basic notification"].isin(FTA_CEPA).any():
    Bond_or_Certificate="C"
    Bond_Cert_Type="MS"
    FTA = CSV_SOURCE[["COO_No ", "COO_Date of Issue "]].drop_duplicates(keep="first",subset=["COO_No "])
    FTA = FTA.dropna(subset=["COO_No ","COO_Date of Issue "])
    FTA["COO_Date of Issue "] = pd.to_datetime(
        FTA["COO_Date of Issue "], dayfirst=True, errors="coerce").dt.date
    Bond_Cert_No = FTA["COO_No "].tolist()


    Bond_Cert_Date=FTA["COO_Date of Issue "].tolist()
    Bond_or_Certificate = [Bond_or_Certificate] * len(Bond_Cert_No)
    Bond_Cert_Type = [Bond_Cert_Type] * len(Bond_Cert_No)
else:
    pass



BONDS_CERTIFICATES["Bond_or_Certificate"]=Bond_or_Certificate
BONDS_CERTIFICATES["Bond_Cert_Type"]=Bond_Cert_Type
BONDS_CERTIFICATES["Bond_Cert_No"]=Bond_Cert_No
BONDS_CERTIFICATES["Bond_Cert_Date"]=Bond_Cert_Date




SUPPORTING_DOCS=pd.DataFrame(columns=["Inv_SrNo",
"Item_SrNo",
"Doc_ICEGATE_ID",
"Doc_IRN",
"Doc_Upload_DateTime",
"Doc_Type",
"File_Type",
"Icegate_File_Name",
"Document_Name",
"Reference_No.",
"Doc_Issued_At",
"Doc_Issued_Date",
"Doc_Expiry_Date",
"Doc_Issuing_Party_Name",
"Doc_Issuing_Party_Code",
"Doc_Issuing_Party_Add1",
"Doc_Issuing_Party_Add2",
"Doc_Issuing_Party_City",
"Doc_Issuing_Party_Pin_Code",
"Doc_Beneficiary_Party_Name",
"Doc_Beneficiary_Party_Code",
"Doc_Beneficiary_Party_Add1",
"Doc_Beneficiary_Party_Add2",
"Doc_Beneficiary_Party_City",
"Doc_Beneficiary_Party_Pin_Code"
])
EXCHANGE_RATE=pd.DataFrame(columns=["CURRENCY_CODE",
"EXCHANGE_RATE",
"BANK_NAME",
"BANK_CERTIFICATE",
"BANK_CERTIFICATE_DATE"
])



# stopped for prventing making excel


with pd.ExcelWriter(rf"{FILE_EXPORT}\XL FILE GENERATED.xlsx",engine="xlsxwriter") as writer:
    GENERAL.to_excel(writer,sheet_name="GENERAL",index=False)
    workbook=writer.book
    worksheet=writer.sheets["GENERAL"]
    for col in GENERAL_HIDE_COLUMNS:
        if col in GENERAL.columns:
            col_idx=GENERAL.columns.get_loc(col)
            worksheet.set_column(col_idx,col_idx,None,None,{'hidden':True})      


    inbond_exbond.to_excel(writer,sheet_name="inbond_exbond".upper(),index=False)
    SHIPMENT.to_excel(writer,sheet_name="SHIPMENT",index=False)
    workbook=writer.book
    worksheet=writer.sheets["SHIPMENT"]
    for ship in SHIPMENT_TAB_HIDE:
        if ship in SHIPMENT.columns:
            ship_index=SHIPMENT.columns.get_loc(ship)
            
            worksheet.set_column(ship_index,ship_index,None,None,{"hidden":True})
            

    CONTAINERS.to_excel(writer,sheet_name="CONTAINERS",index=False)
    INVOICES.to_excel(writer,sheet_name="INVOICES",index=False)
    ITEMS.to_excel(writer,sheet_name="ITEMS",index=False)
    STATEMENT.to_excel(writer,sheet_name="STATEMENT",index=False)
    SEC65_EXBOND_INFO.to_excel(writer,sheet_name="SEC65_EXBOND_INFO",index=False)
    REIMPORT.to_excel(writer,sheet_name="RE-IMPORT",index=False)    
    LICENSE.to_excel(writer,sheet_name="LICENSE",index=False)
    SW_ADDL_INFO.to_excel(writer,sheet_name="SW_ADDL_INFO",index=False)
    SW_CONSTITUENT.to_excel(writer,sheet_name="SW_CONSTITUENT",index=False)
    SW_PRODUCTION.to_excel(writer,sheet_name="SW_PRODUCTION",index=False)
    SW_CONTROL.to_excel(writer,sheet_name="SW_CONTROL",index=False)
    SEZ_INFO.to_excel(writer,sheet_name="SEZ_INFO",index=False)
    HSS.to_excel(writer,sheet_name="HSS",index=False)
    BONDS_CERTIFICATES.to_excel(writer,sheet_name="BONDS_CERTIFICATES",index=False)
    SUPPORTING_DOCS.to_excel(writer,sheet_name="SUPPORTING_DOCS",index=False)
    EXCHANGE_RATE.to_excel(writer,sheet_name="EXCHANGE_RATE",index=False)
    

GREEN="\033[92m"
RESET="\033[0m"
print(GREEN+"IN EXCEL PLEASE PASTE DECLARATION DATA FROM SINGLE WINDOW DATA EXCEL SHEET"+RESET)