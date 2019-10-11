import datetime;
from datetime import datetime;
from datetime import timedelta

import pandas as pd

import google_sheet_common as gs


def gettimenow(timezone):
    try:
        # pno = phno[-10:][:3]
        # timezone = tzs[tzs['Area Code'] == str(pno)]['Timezone'].item()
        indtime = (datetime.utcnow() + timedelta(hours=5.5)).strftime('%H:%M')
        esttime = (datetime.utcnow() - timedelta(hours=4)).strftime('%H:%M')
        centraltime = (datetime.utcnow() - timedelta(hours=5)).strftime('%H:%M')
        pacifictime = (datetime.utcnow() - timedelta(hours=7)).strftime('%H:%M')
        mountaintime = (datetime.utcnow() - timedelta(hours=6)).strftime('%H:%M')
        # if timezone == 'Eastern':
        #     if datetime.strptime('08:30', '%H:%M').time() < \
        #             datetime.strptime(esttime, '%H:%M').time() < datetime.strptime('17:00', '%H:%M').time():
        #         return True
        # if timezone == 'Central':
        #     if datetime.strptime('09:00', '%H:%M').time() < \
        #             datetime.strptime(centraltime, '%H:%M').time() < datetime.strptime \
        #                 ('17:00', '%H:%M').time():
        #         return True
        # if timezone == 'Pacific':
        #     if datetime.strptime('09:00', '%H:%M').time() < \
        #             datetime.strptime(pacifictime, '%H:%M').time() < datetime.strptime \
        #                 ('17:00', '%H:%M').time():
        #         return True
        # if timezone == 'Mountain':
        #     if datetime.strptime('09:00', '%H:%M').time() < \
        #             datetime.strptime(mountaintime, '%H:%M').time() < datetime.strptime \
        #                 ('17:00', '%H:%M').time():
        #         return True
    except:
        print("Error while parsing phone number and timezone time")
        pass
    return True


def updatesheet(cdata):
    track = pd.DataFrame([cdata], columns=[1, 2, 3, 4, 5, 6])
    gs.update_sheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log', track)


def gettimezone(phno, tzs):
    try:
        timezone = tzs[tzs['Area Code'] == phno]['Timezone'].item()
        return timezone
    except:
        return 'Unknown'


# def lastcalldiff(phno):
#     oldcalls = gs.getsheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log')
#     oldcalls['MB Close'] = pd.to_datetime(oldcalls['MB Close'])
#     oldcalls = oldcalls[oldcalls['DateTime'] != '']
#     oldcalls.sort_values('MB Close', ascending=False, inplace=True)
#     oldcalls.drop_duplicates('Number', inplace=True)
#     oldcalls.reset_index(drop=True, inplace=True)
#     try:
#         lstcall = oldcalls[oldcalls['Number'] == phno]['DateTime'].item()
#         if (datetime.strptime(lstcall, '%Y-%m-%d %H:%M:%S')) > (datetime.now() - timedelta(hours=3)):
#             return False
#         else:
#             return True
#     except:
#         return True
