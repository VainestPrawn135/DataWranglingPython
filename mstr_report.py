from email.mime import base
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

REPORT_ID = '89B856D14067B9A8EF62B3A5ED656935'
#REPORT_ID = '7CC43D764BB1E854D88F59B38AEB6BA6'

my_report = Report(connection=conn, id=REPORT_ID, parallel=False)
my_report_df = my_report.to_dataframe()

new_df = my_report_df[['Year', 'Month@ID', 'Net Sales', 'Net Sales YTD', 'Innovation Sales', '% Innovation Sales / Net Sales', 'Target $', 'Target %', 'Gross Profit', 'Gross Profit (Flag Innovation)', 'Innovation vs. Core (Gross Profit)', 'Marginal Contribution', 'Marginal Contribution (Flag Innovation)', 'Innovation vs. Core (Marginal Contribution)', 'Net Volume KG']]

print(new_df)

new_df.to_excel('mstr.xlsx')

conn.close()