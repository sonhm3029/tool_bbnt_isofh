# Working with logtime file
import pandas as pd
# Working with word
from docxtpl import DocxTemplate, RichText
from docx.shared import Cm, Inches, Mm, Emu
from datetime import date, datetime
import calendar
# Working with env variables
import os
from dotenv import load_dotenv
load_dotenv()


df = pd.read_csv('./list_ticket.csv')


# Show df
print(df[['Issue Key', 'Issue summary']])


# Show duplicate value
print(f"Duplicate value: {df.duplicated('Issue Key')}")

# Remove duplicate value
df.drop_duplicates(inplace=True,subset=['Issue Key'] )

# Take information
new_df = df[['Issue Key', 'Issue summary']]
# for item in new_df.loc:
#     print(item[0])
df_np_arr = new_df.to_numpy()
# # print to new csv
# new_df.to_csv('cv_t7_filter.csv', index=False)
# Generate docx templates:

template = DocxTemplate('./template.docx')

# Template variables
month = date.today().month
year = date.today().year
first, last = calendar.monthrange(year, month)

from_date = datetime(year, month, 1).strftime('%d-%m-%Y')
to_date = datetime(year, month, last).strftime('%d-%m-%Y')

# Template table ticket:
table_tickets = [{'ticket_title':f'{item[0]} {item[1]}', 'link_ticket':f'https://jira.isofh.com.vn/browse/{item[0]}', 'issue_key':item[0]} for item in df_np_arr]
# add hyperlinks
for item in table_tickets:
    rt = RichText()
    rt.add(item['link_ticket'], url_id=template.build_url_id(item['link_ticket']),color="blue",underline=True)
    item['link_ticket']=rt
    

context = {
    'from_date':from_date,
    'to_date':to_date,
    # ÔNg/bà:
    'your_name': os.environ['your_name'],
    # CMND/CCCD số	
    'cccd': os.environ['cccd'],
    # Ngày sinh
    'born_date': os.environ['born_date'],
    # Ngày cấp
    'ngay_cap': os.environ['ngay_cap'],
    # Địa chỉ thường trú
    'dia_chi_tc': os.environ['dia_chi_tc'],
    # Địa chỉ liên hệ
    'dia_chi_lh': os.environ['dia_chi_lh'],
    # Điện thoại
    'dien_thoai': os.environ['dien_thoai'],
    # Mã số thuế
    'ma_so_thue': os.environ['ma_so_thue'],
    # Tài khoản	:
    'tai_khoan': os.environ['tai_khoan'],
    # Tại Ngân hàng	
    'ngan_hang': os.environ['ngan_hang'],
    # Mã nhân viên:
    'your_staff_code':os.environ['your_staff_code'],
    # list ticket
    'table_ticket':table_tickets,
}

file_name = 'Hoàng Minh Sơn- BB nghiệm thu CTV-T7.2022'

template.render(context)
template.save(f"{file_name}.docx")

