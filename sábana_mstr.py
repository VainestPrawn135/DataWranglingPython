from mstrio import connection
from mstrio.project_objects import OlapCube, Report

base_url_dev = "http://siaazmicdev:8080/MicroStrategyLibrary/api"
base_url_prod = "http://siaazmicapprd:8080/MicroStrategyLibrary/api"

conn = connection.Connection(
    base_url=base_url_dev,
    username="dmorenom",
    password="Welcome2022.08",
    project_name="INNOVATION"
)
REPORT_ID_PROD = 'F884AE844F503775C215608F2F1AF624'
REPORT_ID_DEV = '3CC1A69F46591C438C6BCA822A7FFFA8'

my_report = Report(connection=conn, id=REPORT_ID_DEV, parallel=False)
my_report_df = my_report.to_dataframe()

new_df = my_report_df[["0CALMONTH_KEY",
	   "0CALMONTH2_KEY",
	   "0CALQUART1_KEY",
	   "0CALQUARTER_KEY",
	   "0CALYEAR_KEY",
	   "YSD_CH194_KEY",
	   "YSD_CH178_KEY",
	   "YSD_CH012__YSD_CH013_KEY",
	   "YSD_CH200__YSD_CH195_KEY",
	   "YSD_CH200__YSD_CH196_KEY",
	   "YSD_CH200__YSD_CH197_KEY",
	   "YSD_CH001_KEY",
       "YSD_KF019_0",
	   "TARGET_AMOUNT",
	   "TARGET_PERCENTAGE",
	   "CC_NETSALES_USD",
	   "CC_VARCOST_USD",
	   "CC_FIXCOST_USD",
	   "CC_MARGCONTRIB_USD",
	   "CC_GROSSPROFIT_USD",
	   "CC_NETVOLUME_KG"]]

print(new_df)

new_df.to_excel('Validaciones.xlsx')

conn.close()