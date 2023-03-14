import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import datetime
import calendar
import tempfile
import convertapi

# hide hamburger menu
st.markdown('<style>#MainMenu {visibility: hidden;}</style>', unsafe_allow_html=True)

# set page title tag
st.set_page_config(page_title='BCTDTS')

# write title and subtitle
st.title('💰 BCTDTS')
st.write('Quick timesheet maker for Baker boys. Created by Saleh.')

# create form and input widgets
form = st.form('input_form')
employee_name = form.text_input('Full Name')
employee_id = form.number_input('Employee ID', step=1, min_value=0)
employee_rate = form.number_input('Day Rate (SAR)', min_value=0.0)
date_start = form.date_input('Starting Date')
date_end = form.date_input('Ending Date')
rig_name = form.selectbox('Rig Name',
    ['', 'BCTD-4', 'BCTD-5'])
wstl_name = form.selectbox('Wellsite Team Leader',
    ['', 'Ahmed Mansour', 'Ken Lynn', 'Pete Riley', 'Steve Baranyi'])

# create form submit button
submitted = form.form_submit_button('Submit')

# on click
if submitted:

    # check/throw error: starting date after ending date 
    if date_start > date_end:
        st.error('Starting date must be earlier than ending date.')

    # check/throw error: starting date and ending date: same month/year 
    elif date_start.month != date_end.month or date_start.year != date_end.year:
         st.error('At the moment, timesheets can be only created for an individual month. \
         Hence, starting date and ending date must be in the same month/year. \
         The possibility of making timesheets that span over multiple months might be added \
         later depnding on my mood and freetime.')

    else:
        with st.spinner('Working on your timesheet...'):
            # quick access easter-egg
            quick_accessor = employee_name
            if quick_accessor.startswith('#'):
                try:
                    # retrieve data form environment variables (streamlit secrets)
                    employee_id = st.secrets[quick_accessor[1:]]['id']
                    employee_name = st.secrets[quick_accessor[1:]]['name']
                    employee_rate = st.secrets[quick_accessor[1:]]['rate']
                    rig_name = st.secrets[quick_accessor[1:]]['rig']
                except:
                     pass # bad programming: guilty!

            # load workbook and set active worksheet
            wb = load_workbook(filename=r'template.xlsx', read_only=False)
            ws = wb['timesheet']

            # remove footer
            ws.oddFooter.left.text = ''
            ws.oddFooter.center.text = ''
            ws.oddFooter.right.text = ''

            # print month range (1 : end_of_month) in column A
            month_start = 1
            month_end = calendar.monthrange(date_start.year, date_start.month)[1] + 1
            for day in range(month_start, month_end):
                cell_a = 'A' + str(day + 1)
                ws[cell_a] = day

            # print hitch range (hitch_start : hitch_end) in column B and column C
            hitch_start = date_start.day
            hitch_end = date_end.day + 1
            for day in range(hitch_start, hitch_end):
                cell_b = 'B' + str(day + 1)
                cell_d = 'D' + str(day + 1)
                ws[cell_b] = 'ARAMCO'
                ws[cell_d] = rig_name.upper()

            # print all other info
            ws['Q2']= employee_name.upper()
            ws['Q3']= employee_id
            ws['Q4']= 'KSA'
            ws['Q5']= str(calendar.month_abbr[date_start.month].upper()) + ' ' + str(date_start.year) # month year

            hitch_days = len(range(hitch_start, hitch_end))
            ws['O8']= hitch_days
            ws['Q8']= employee_rate
            ws['T8']= hitch_days * employee_rate
            ws['T19']= hitch_days * employee_rate

            ws['O22']= employee_name.upper()
            ws['O24']= wstl_name.upper()
            ws['O26']= 'JAHAD ALDAWOOD'

            # set sheet name to 'MMMYYYY'
            sheet_name = '{}'.format(str(calendar.month_abbr[date_start.month].upper()) + str(date_start.year))
            ws.title = sheet_name

            # set file name to 'TS-ID-MMMYYYY'
            file_name = 'TS-{}-{}'.format(
                employee_id, str(calendar.month_abbr[date_start.month].upper()) + str(date_start.year))

            # save modified workbook to stream
            with tempfile.NamedTemporaryFile() as tmp:
                wb.save(tmp.name)
                output = BytesIO(tmp.read())

            # upload modified workbook and convert to pdf
            convertapi.api_secret = st.secrets['api_secret']
            upload_io = convertapi.UploadIO(output.getvalue(), 'ts-{}.xlsx'.format(employee_id))
            result = convertapi.convert('pdf', {'File': upload_io})
            saved_file = result.file.save(tempfile.gettempdir())

            # display success message and download button(s)
            st.success('Your timesheet has been successfully generated.')

            ## download excel file
            # st.download_button(
            #     label = 'Download EXCEL File',
            #     data = output.getvalue(),
            #     file_name = file_name,
            #     mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            # )

            # download pdf file
            with open(saved_file, 'rb') as file:
                btn = st.download_button(
                        label = 'Download Timesheet',
                        data = file,
                        file_name = file_name,
                        mime = 'application/pdf'
                    )