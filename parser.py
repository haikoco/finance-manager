import requests
from bs4 import BeautifulSoup
import pandas as pd

writer = pd.ExcelWriter('hh.xlsx', engine='xlsxwriter')

yourData.to_excel(writer, 'Sheet1')

writer.save()