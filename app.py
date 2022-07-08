import streamlit as st
import pyperclip as pc
from pynput.keyboard import Key, Controller
from docxtpl import DocxTemplate
from datetime import date
from docx2pdf import convert

# ***Date Format For Excel***
dt = date.today()
dt = dt.strftime("%d-%b-%Y")

# ***Keyboard Control to Clear the Page***
keyboard = Controller()

st.title('Direct Export App')

jobNum, hwb, mawb, bill = st.columns(4)

jobNum = jobNum.text_input('Job Number', max_chars=11)
hwb = hwb.text_input('HWB')
if hwb == '' or hwb.isnumeric():
    pass
else:
    st.error('Invalid HWB Number')
mawb = mawb.text_input('MAWB', max_chars=12)
# if mawb == '' or mawb == '-' or mawb.isnumeric():
#     pass
# else:
#     st.error('Invalid MAWB Number')
bill = bill.text_input('Billing', max_chars=3)
if bill.isnumeric():
    st.error('Invalid Billing')
else:
    pass

customer = st.text_input('Customer')

pu_apc, dest_apc, date_colu1, pu_time, airline = st.columns(5)

pu_date = date_colu1.date_input('P/U Date')
dest_apc = dest_apc.text_input('Arrival APC', max_chars=3)
pu_date = pu_date.strftime("%d-%b")
pu_time = pu_time.text_input('Pick Up Time', max_chars=5)
pu_apc = pu_apc.text_input('Pick Up APC', max_chars=3)
airline = airline.text_input('Airline', max_chars=2)

departing_apc, flt1, date_colu2, dept_time, flt1_arr_date, flt1_arr_time = st.columns(6)

departing_apc = departing_apc.text_input('Dep APC', max_chars=3)
flt1 = flt1.text_input('FLT-1#')
tend_date = date_colu2.date_input('FLT-1 Dep Date')
tend_date = tend_date.strftime("%d-%b")
dept_time = dept_time.text_input('FLT-1 Dep Time', max_chars=5)
flt1_arr_date = flt1_arr_date.date_input('FLT1 Arr Date')
flt1_arr_date = flt1_arr_date.strftime('%d-%b')
flt1_arr_time = flt1_arr_time.text_input('FLT1 Arr Time', max_chars=5)

xfer, flt2, date_colu3, xfer_dept_time, xfer_arr_date, xfer_time = st.columns(6)

xfer = xfer.text_input('Transfer APC', max_chars=3)
xfer_date = date_colu3.date_input('FLT2 Dep Date')
xfer_date = xfer_date.strftime("%d-%b")
flt2 = flt2.text_input('FLT-2#')
xfer_dept_time = xfer_dept_time.text_input('FLT-2 Dep Time', max_chars=5)
xfer_arr_date = xfer_arr_date.date_input('FLT2 Arr Date')
xfer_arr_date = xfer_arr_date.strftime("%d-%b")
xfer_time = xfer_time.text_input('FLT2 Arr Time', max_chars=5)

pcs, wgt, dimwgt, charges = st.columns(4)

pcs = pcs.text_input('Piece Count')
wgt = wgt.text_input('Weight (kgs)')
dimwgt = dimwgt.text_input('DimWeight(kgs)')
charges = charges.text_input('Freight Cost')

desc_handle_cont = st.container()

colu3, colu4 = desc_handle_cont.columns(2)

desc = colu3.text_input('Nature and Goods')
handling = colu4.text_input('Handling Instructions')
desc2 = colu3.text_input(' ')
handling2 = colu4.text_input('    ')
desc3 = colu3.text_input('     ')
handling3 = colu4.text_input('            ')

osi_acct_cont = st.container()

osi_colu, acct_colu = osi_acct_cont.columns(2)

osi = osi_colu.text_area('SSR')
acct = acct_colu.text_area('Accounting Information')

pre_paid, eap = st.columns(2)

pre_paid = pre_paid.checkbox('Pre-Paid', value=True)
eap = eap.checkbox('EAP', value=True)

pil, rds, col, crt, fro, ice = st.columns(6)

pil = pil.checkbox('PIL')
rds = rds.checkbox('RDS')
col = col.checkbox('COL')
crt = crt.checkbox('CRT')
fro = fro.checkbox('FRO')
ice = ice.checkbox('ICE')

ship_info, consgn_info = st.columns(2)

ship_info = ship_info.checkbox('Shipper Information')
consgn_info = consgn_info.checkbox('Conignee Information')

ship_consg_cont = st.container()

colu1, colu2 = ship_consg_cont.columns(2)

shipper_name = colu1.text_input('Shipper', max_chars=35)
consignee_name = colu2.text_input('Consignee', max_chars=35)
shipper_name2 = colu1.text_input(' ', max_chars=35)
consignee_name2 = colu2.text_input('    ', max_chars=35)
shipper_addr = colu1.text_input('Address', max_chars=35)
consignee_addr = colu2.text_input('Address    ', max_chars=35)
shipper_addr2 = colu1.text_input('   ', max_chars=35)
consignee_addr2 = colu2.text_input('      ', max_chars=35)
shipper_city = colu1.text_input('City', max_chars=17)
consignee_city = colu2.text_input('City    ', max_chars=17)
shipper_zip = colu1.text_input('Postal Code', max_chars=9)
consignee_zip = colu2.text_input('Postal Code     ', max_chars=9)
shipper_state = colu1.text_input('State', max_chars=9)
consignee_state = colu2.text_input('State   ', max_chars=9)
shipper_country = colu1.text_input('Country')
consignee_country = colu2.text_input('Country     ')
shipper_phone = colu1.text_input('Phone Number', max_chars=25)
consignee_phone = colu2.text_input('Phone Number     ', max_chars=25)
shipper_fax = colu1.text_input('Fax', max_chars=25)
consignee_fax = colu2.text_input('Fax     ', max_chars=25)


# ***FUNCTIONS FOR BUTTONS***
def clear():
    with keyboard.pressed(Key.ctrl):
        keyboard.press('r')
        keyboard.release('r')


def copy_to_log():
    log_info = '\t' + jobNum + '\t' + hwb + '\t' + mawb + '\t' + '\t' + pu_apc + '\t' + departing_apc + '\t' + xfer + '\t' \
               + '\t' + '\t' + dest_apc + '\t' + airline + '\t' + pu_date + '\t' + pu_time + '\t' + tend_date + '\t' + desc \
               + '\t' + wgt + '\t' + dimwgt + '\t' + '\t' + '\t' + '\t' + charges + '\t' + bill + '\t' + '\t' + customer
    pc.copy(log_info)


def agent_dispatch_cl():
    if xfer:
        doc = DocxTemplate('AGENT_DISPATCH_TEMPLATE.docx')
        context = {'date': dt, 'pickup_apc': pu_apc, 'job_num': jobNum, 'mawb': mawb,
                   'pickup_date': pu_date + '-' + '2022',
                   'route': pu_apc + '-' + xfer + '-' + dest_apc, 'dep_apc': departing_apc, 'xfer': xfer,
                   'airline': airline, 'flight1': flt1, 'dep_date_time': tend_date + '-2022' + '   ' +
                    dept_time, 'arr_apc': dest_apc, 'arr_date_time': flt1_arr_date + flt1_arr_time,
                   'flight2':
                       flt2,
                   'dep_date_time2': xfer_date + '-2022' + '   ' +
                    dept_time, 'arr_date_time2': xfer_arr_date + xfer_time}

        doc.render(context)
        doc.save('AgentDispatchForm_{}.docx'.format(jobNum))
    else:
        doc = DocxTemplate('AGENT_DISPATCH_TEMPLATE_2.docx')
        context = {'date': dt, 'pickup_apc': pu_apc, 'job_num': jobNum, 'mawb': mawb,
                   'pickup_date': pu_date + '-' + '2022',
                   'route': pu_apc + '-' + xfer + '-' + dest_apc, 'dep_apc': departing_apc, 'xfer': xfer,
                   'airline': airline, 'flight1': flt1,
                   'dep_date_time': tend_date + '-2022' + '   ' + dept_time}
        doc.render(context)
        doc.save('AgentDispatchForm_{}.docx'.format(jobNum))


# ***BUTTONS***
st.sidebar.button('Submit to WebDocs')
st.sidebar.button('Copy to Log', on_click=copy_to_log)
st.sidebar.button('Copy for Email')
st.sidebar.button('Agent Dispatch Check List', on_click=agent_dispatch_cl())
st.sidebar.markdown('##')
st.sidebar.markdown('##')
st.sidebar.markdown('##')
st.sidebar.button('Clear', on_click=clear)
