import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import datetime
import calendar
import tempfile
import convertapi
from tempfile import NamedTemporaryFile
import os
from staff import staff

def get_next_month(input_date):
    new_year = input_date.year
    
    new_month = input_date.month + 1
    if new_month > 12:
        new_year += 1
        new_month -= 12
   
    new_day = calendar.monthrange(new_year, new_month)[1]
        
    return datetime.date(new_year, new_month, new_day)

# set page title tag
st.set_page_config(page_title='BCTDTS')

# hide hamburger menu
st.markdown('<style>#MainMenu {visibility: hidden;}</style>', unsafe_allow_html=True)

# write title and subtitle
st.title('💰 BCTD Timesheet Maker')
st.write('The easiest way to make timesheets for Baker boys. Created by: Saleh.')

# decrypt imported staff dict from staff.py
staff_dict = {}
for k, v in staff.items():
    name = v[0]
    rig = 'BCTD-' + v[1].split(':')[0]
    # crew = 'CREW-' + v[1].split(':')[1]
    staff_dict[k] = [name, rig] # , crew]

# make presentable list for multiselect options
staff_list = []
for i in list(sorted(staff_dict.keys())):
    staff_list.append(str(i) + ': ' + staff_dict[i][0])

# display multiselect input widget
quick_multiselect = st.multiselect('quick_multiselect', staff_list, label_visibility='collapsed')

# initiate main lists to collect options from multiselect
ids = []
names = []
rigs = []
rates = []

# display other fields once selection is made 
if len(quick_multiselect) > 0:
	
	# loop over multiselect and append related values to main lists
	# display rate input widgets for each
    for i in quick_multiselect:

        idx = int(i.split(":")[0])
        name = staff_dict[idx][0]
        rig = staff_dict[idx][1]
        rate = st.number_input(
            'Day Rate (SAR) - {}: {}'.format(idx, name),
            min_value=0.0, key='rate_{}'.format(idx))

        ids.append(idx)
        names.append(name)
        rigs.append(rig)
        rates.append(rate)
	
	# display dates input widgets
	# date_end min_value is set to prohibit date_end < date_start
	# date_end max_value is set to prohibit date_end > end_of_next_month(date_start)
    date_start = st.date_input('Starting Date')
    date_end = st.date_input('Ending Date',
                    value=date_start, min_value=date_start, max_value=get_next_month(date_start))
	
	# create condition for later uses; check if desired output is double sheets				
    double_sheets = True if date_end.month == get_next_month(date_start).month else False
	
	# create and display rig options
    rig = rigs[0]
	
	# create rig:wstl dict for cascading options
    wstl_dict = {
        'BCTD-4': ['Ken Lynn', 'Pete Riley'],
        'BCTD-5': ['Steve Baranyi', 'Ahmed Mansour']
		}
	
	# create wstl label when outputting single sheet
    wstl_label_a = 'Wellsite Team Leader'
	
	# create wstl labels when outputting double sheet
    if double_sheets:
        wstl_label_a = 'Wellsite Team Leader - {} {}'.format(
                    str(calendar.month_abbr[date_start.month].upper()), str(date_start.year))

        wstl_label_b = 'Wellsite Team Leader - {} {}'.format(
                    str(calendar.month_abbr[date_end.month].upper()), str(date_end.year))
	
	# display wstl input widget for both single and double sheets
    wstl_month_a = st.selectbox(wstl_label_a, wstl_dict[rig])
	
	# display another wstl input widget for double sheets
    if double_sheets:
        wstl_month_b = st.selectbox(wstl_label_b, wstl_dict[rig])
		
	# display submit and generate file button
    submitted = st.button('Submit')

    if submitted:
        with st.spinner('Working on your timesheet...'):
			
			# Flow:
			# 1	load workbook and template worksheet
			# 2	fill-in common cells between all worksheets regardless
			# 	of date-specific or employee-specific info
			# 3	if double sheets; dublicte worksheet
			# 4	continue filling-in date-specific cells for each worksheet
			# 5	if mutliselects; copy as many related worksheets
			# 6	continue filling-in employee-specific cells for each worksheet
			
			# load workbook
            wb = load_workbook(filename=r'template.xlsx', read_only=False)
			
			# load base worksheet
            ws1 = wb['timesheet']

            # remove footer
            ws1.oddFooter.left.text = ''
            ws1.oddFooter.center.text = ''
            ws1.oddFooter.right.text = ''
			
			# fill-in common cells
            ws1['Q4']= 'KSA'
            ws1['O26']= 'JAHAD ALDAWOOD'
			
			# if double sheets; dublicte worksheet before filling-in date-specific cells 
            if double_sheets:
                ws2 = wb.copy_worksheet(ws1)
			
			# continue filling-in date-specific cells for worksheet_a
            # print month range (1 : end_of_month) in column A
            month1_start = 1
            month1_end = calendar.monthrange(date_start.year, date_start.month)[1] + 1
            for day in range(month1_start, month1_end):
                cell_a = 'A' + str(day + 1)
                ws1[cell_a] = day
			
			# fill-in month cell and set sheet title
            ws1['Q5']= str(calendar.month_abbr[date_start.month].upper()) + ' ' + str(date_start.year)
            ws1.title = '{}'.format(str(calendar.month_abbr[date_start.month].upper()) + str(date_start.year))
			
			# continue filling-in date-specific cells for worksheet_b
            # print month range (1 : end_of_month) in column A
            if double_sheets:
                month2_start = 1
                month2_end = calendar.monthrange(date_end.year, date_end.month)[1] + 1
                for day in range(month2_start, month2_end):
                    cell_a = 'A' + str(day + 1)
                    ws2[cell_a] = day
					
				# fill-in month cell and set sheet title
                ws2['Q5']= str(calendar.month_abbr[date_end.month].upper()) + ' ' + str(date_end.year)
                ws2.title = '{}'.format(str(calendar.month_abbr[date_end.month].upper()) + str(date_end.year))
			
			# copy either worksheet_a or/and worksheet_b and
			# fill-in employee-specific cells for each employee 
            for i in ids:
				# get employee info from main lists
                name = names[ids.index(i)].upper()
                rig = rigs[0].upper() #rigs[ids.index(i)]
                rate = rates[ids.index(i)]
				
				# copy worksheet_a
                ws1_aux = wb.copy_worksheet(ws1)
				
				# continue filling-in employee-specific cells
                # print hitch range (hitch_start : hitch_end) in column B and column C
                hitch1_start = date_start.day
                hitch1_end = month1_end if double_sheets else date_end.day + 1
                for day in range(hitch1_start, hitch1_end):
                    cell_b = 'B' + str(day + 1)
                    cell_d = 'D' + str(day + 1)
                    ws1_aux[cell_b] = 'ARAMCO'
                    ws1_aux[cell_d] = rig

                ws1_aux['Q2']= name
                ws1_aux['Q3']= i

                hitch = len(range(hitch1_start, hitch1_end))
                ws1_aux['O8']= hitch
                ws1_aux['Q8']= rate
                ws1_aux['T8']= hitch * rate
                ws1_aux['T19']= hitch * rate

                ws1_aux['O22']= name
                ws1_aux['O24']= wstl_month_a.upper()
				
				# set sheet title using month from source and employee id
                ws1_aux.title = str(i) + ' - ' + ws1.title
				
				# copy worksheet_b if double sheets
                if double_sheets:
                    ws2_aux = wb.copy_worksheet(ws2)
					
					# continue filling-in employee-specific cells
                    # print hitch range (hitch_start : hitch_end) in column B and column C
                    hitch2_start = 1
                    hitch2_end = date_end.day + 1
                    for day in range(hitch2_start, hitch2_end):
                        cell_b = 'B' + str(day + 1)
                        cell_d = 'D' + str(day + 1)
                        ws2_aux[cell_b] = 'ARAMCO'
                        ws2_aux[cell_d] = rig

                    ws2_aux['Q2']= name
                    ws2_aux['Q3']= i

                    hitch = len(range(hitch2_start, hitch2_end))
                    ws2_aux['O8']= hitch
                    ws2_aux['Q8']= rate
                    ws2_aux['T8']= hitch * rate
                    ws2_aux['T19']= hitch * rate

                    ws2_aux['O22']= name
                    ws2_aux['O24']= wstl_month_b.upper()
					
					# continue filling-in employee-specific cells
                    ws2_aux.title = str(i) + ' - ' + ws2.title
        
            # delete worksheets used as template for date-specific cells         
            del wb[ws1.title]
            if double_sheets:
                del wb[ws2.title]
		
            # set filename 
            file_name = ''.join(str(i) for i in ids)

            # save modified workbook to stream
            with tempfile.NamedTemporaryFile() as tmp:
                wb.save(tmp.name)
                output = BytesIO(tmp.read())

            # upload modified workbook and convert to pdf
            convertapi.api_secret = os.environ.get('api_secret')
            upload_io = convertapi.UploadIO(output.getvalue(), '{}.xlsx'.format(file_name))
            result = convertapi.convert('pdf', {'File': upload_io})
            saved_file = result.file.save(tempfile.gettempdir())

            # display success message and download button(s)
            st.success('Timesheet(s) has been successfully generated.')
        
            # # download excel file
            # st.download_button(
            #     label = 'Download EXCEL File',
            #     data = output.getvalue(),
            #     file_name = file_name,
            #     mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            # )

            # download pdf file
            with open(saved_file, 'rb') as file:
                st.download_button(
                    label = 'Download Timesheet(s)',
                    data = file,
                    file_name = 'TS-' + str(datetime.date.today()),
                    mime = 'application/pdf'
                )
			