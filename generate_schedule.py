#!/usr/bin/env python
import datetime
import os
import openpyxl
import sys
import zoneinfo
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def is_merged_cell(sheet, cell):
    for merged_cell_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_cell_range:
            return True
    return False


def capitalize_place_name(name):
    return ' '.join([word.capitalize() for word in name.split()])


def get_waste_category(category):
    lc = category.lower()
    if 'zmieszane' in lc:
        return 'ZMIESZANE', '8'
    elif 'bio' in lc:
        return 'BIO', '6'
    elif 'szk' in lc:
        return 'SZKŁO', '10'
    elif 'pap' in lc:
        return 'PAPIER', '7'
    elif 'szt' in lc:
        return 'PLASTIK', '5'
    elif 'pla' in lc:
        return 'PLASTIK', '5'
    else:
        assert False


def get_datetime_month_name(month_name):
    pl = [
        'sty', 'lut', 'mar', 'kwi', 'ma', 'cze', 'lip', 'sie', 'wrz', 'paź',
        'lis', 'gru'
    ]
    en = [datetime.date(2008, i, 1).strftime('%B').lower()
          for i in range(1, 13)]
    for i, short in enumerate(pl):
        if month_name.startswith(short):
            return en[i].capitalize()
    return None


def get_months_column_ranges(sheet, range):
    months = {}
    current_month = None
    for cell in sheet[range][0]:
        if is_merged_cell(sheet, cell):
            column_letter = openpyxl.utils.cell.get_column_letter(cell.column)
            if cell.value is not None and not ' ' in cell.value.strip():
                current_month = get_datetime_month_name(cell.value)
                months[current_month] = [column_letter]
            if current_month and isinstance(cell, openpyxl.cell.cell.MergedCell):
                assert cell.value is None
                months[current_month].append(column_letter)
    return months


def get_schedule_summary_and_info(sheet, original_summary):
    villages = [
        'janowice wielkie', 'komarno', 'miedzianka', 'mniszków', 'radomierz',
        'trzcińsko'
    ]
    original_summary = original_summary.replace('\n', ' ').replace(
        '  ', ' ').strip('*').strip()
    info = []
    places = []
    if 'cała gmina' in original_summary.lower():
        summary = 'Gmina Janowice Wielkie - Cała Gmina'
        places.append('Gmina Janowice Wielkie')
    elif 'wielorodzinne' in original_summary.lower():
        summary = 'Gmina Janowice Wielkie - Wielorodzinne'
        places.append('Gmina Janowice Wielkie')
        info = [
            v.value.strip('* ') for v in sheet['A36:Z36'][0] if v.value is not None
        ]
    else:

        def get_villages(villages, s):
            s_lower = s.lower()
            ps = []  # position range of village and addresses
            for v in villages:
                p = s_lower.find(v)
                if p != -1:
                    ps.append([p, len(v), -1])

            for i, v in enumerate(ps):
                p = v[0] + v[1]
                p2_min = -1
                for v in villages:
                    p2 = s_lower.find(v, p)
                    if p2 != -1:
                        p2_min = min(p2, p2_min) if p2_min > -1 else p2
                        ps[i][2] = p2_min

            assert len(ps) <= len(villages), '{0} vs {1}'.format(ps, villages)

            places = {}
            for p in ps:
                b = p[0]
                e = p[2]
                v = s[b:e].strip().strip('-:, ')
                place = capitalize_place_name(v[0:p[1]].strip().strip('-:, '))
                places[place] = v[p[1]:].strip().strip('-:, ')

            return places

        places = get_villages(villages, original_summary)
        summary = ', '.join(places.keys())
    return summary, info, places


def generate_schedule(xlsx_path):
    assert os.path.exists(xlsx_path)
    book = openpyxl.load_workbook(xlsx_path)
    sheet = book.active
    schedule_info = [
        c.value.replace('…', '.').strip()
        for r in sheet['A1:Z9']
        for c in r
        if c.value is not None
    ]
    print(f"Schedule: {schedule_info[0]} - {schedule_info[1]}")
    schedule_year = [
        v.value for v in sheet['A10:AZ10'][0] if v.value is not None
    ][0]
    print(f"Schedule year: {schedule_year}")

    schedule = []
    months_ranges = get_months_column_ranges(sheet, 'A11:AZ11')
    waste_category = None
    for row_number in range(12, 35):
        summary_values = [
            v.value
            for v in sheet[f'A{row_number}:F{row_number}'][0]
            if v.value is not None
        ]
        assert len(summary_values) == 1
        summary = summary_values[0]

        if not any([v.value for v in sheet[f'G{row_number}:AJ{row_number}'][0]]):
            waste_category, _ = get_waste_category(summary)
            continue

        schedule_entry = {
            'summary': summary,
            'waste': waste_category,
            'info': None,
            'dates': {},
            'places': []
        }
        for month, month_cells in months_ranges.items():
            month_range = f'{month_cells[0]}{row_number}:{month_cells[-1]}{row_number}'
            month_days = [
                int(str(v.value).strip('*'))
                for v in sheet[month_range][0]
                if v.value is not None
            ]
            schedule_entry['dates'][month] = month_days

        summary, info, places = get_schedule_summary_and_info(sheet, summary)
        schedule_entry['summary'] = summary
        schedule_entry['info'] = info
        schedule_entry['places'] = places
        schedule.append(schedule_entry)

    return schedule_year, schedule_info, schedule


def get_google_calendar_credentials():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                             SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def list_calendars(service):
    calendars = []
    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            calendar = {
                'id': calendar_list_entry['id'],
                'summary': calendar_list_entry['summary']
            }
            if 'description' in calendar_list_entry:
                calendar['description'] = calendar_list_entry['description']
        calendars.append(calendar)

        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
    return calendars


def create_calendar(service, year, location, timezone):
    new_calendar = {
        'summary':
            'Janowice Wielkie - Odpady {0}'.format(year),
        'description':
            'Harmonogram odbioru odpadów komunalnych w gminie Janowice Wielkie w roku {0}'
            .format(year),
        'location':
            location,
        'timeZone':
            timezone
    }
    for existing_calendar in list_calendars(service):
        if new_calendar['summary'].lower().strip(
        ) == existing_calendar['summary'].lower().strip():
            print('Calendar \'{0}\' exists. Deleting.'.format(
                existing_calendar['summary']))
            # TODO: Clear events instead of deleting calendar, but this command seems broken
            # service.calendars().clear(calendarId=existing_calendar['id']).execute()
            service.calendars().delete(
                calendarId=existing_calendar['id']).execute()

    print("Creating calendar '{0}'".format(new_calendar['summary']))
    created = service.calendars().insert(body=new_calendar).execute()
    # Make the calendar public and read-only for public
    rule = {
        'scope': {
            'type': 'default',
            'value': '',
        },
        'role': 'reader'
    }
    created_rule = service.acl().insert(
        calendarId=created['id'], body=rule).execute()
    assert created_rule['kind'] == 'calendar#aclRule'
    # Modify default colours
    calendar_list_entry = service.calendarList().get(
        calendarId=created['id']).execute()
    assert created['id'] == calendar_list_entry['id']
    calendar_list_entry['colorRgbFormat'] = 'True'
    calendar_list_entry['backgroundColor'] = '#c2c2c2'
    calendar_list_entry['foregroundColor'] = '#000000'
    calendar_list_entry['colorId'] = '19'
    updated_calendar = service.calendarList().update(
        calendarId=calendar_list_entry['id'], body=calendar_list_entry).execute()
    assert calendar_list_entry['id'] == updated_calendar['id']
    calendar_list_entry = service.calendarList().get(
        calendarId=updated_calendar['id']).execute()
    # Retrieve CID
    from base64 import b64encode
    calendar_list_entry['cid'] = b64encode(
        calendar_list_entry['id'].encode('utf-8')).decode().rstrip('=')
    return calendar_list_entry


def make_event_description(entry, schedule_info):
    d = '<h2>Miejsca zbiórki odpadów:</h2>'
    d += '<ul>'
    if isinstance(entry['places'], dict):
        for k, v in entry['places'].items():
            d += '<li>{0}: {1}</li>'.format(k, v)
    if isinstance(entry['places'], list):
        for v in entry['places']:
            d += '<li>{0}</li>'.format(v)
    for v in entry['info']:
        d += '<li>{0}</li>'.format(v)
    d += '</ul>'

    d += '<h2>Informacje:</h2>'

    d += '<p><a href="https://www.janowicewielkie.eu/index.php/dla-mieszkanca/gospodarka-odpadami-komunalnymi">Gospodarka odpadami komunalnymi</a></p>'
    for v in schedule_info:
        d += '<p>{0}</p>'.format(v)

    return d


def make_event_datetime(year, month_name, day, timezone):
    dt = datetime.datetime.strptime(month_name, "%B")
    start = datetime.datetime(
        year, dt.month, day, 6, tzinfo=zoneinfo.ZoneInfo(timezone))
    end = datetime.datetime(
        year, dt.month, day, 20, tzinfo=zoneinfo.ZoneInfo(timezone))
    return start.isoformat(), end.isoformat()


def main(xlsx_path):
    schedule_location = 'Gmina Janowice Wielkie, Poland'
    schedule_timezone = 'Europe/Warsaw'
    xlsx_path = os.path.abspath(xlsx_path)
    schedule_year, schedule_info, schedule = generate_schedule(xlsx_path)

    created_events = []
    try:
        service = build(
            'calendar', 'v3', credentials=get_google_calendar_credentials())
        calendar = create_calendar(service, schedule_year, schedule_location,
                                   schedule_timezone)
    except HttpError as error:
        sys.exit(f'An error occurred: {error}')

    event_count = 0
    for i, entry in enumerate(schedule):
        summary = '{0}: {1}'.format(entry['waste'], entry['summary'])
        print(f'{i}\t{summary}')
        description = make_event_description(entry, schedule_info)
        _, event_color_id = get_waste_category(entry['waste'])
        for month_name, days in entry['dates'].items():
            for day in days:
                start_dt, end_dt = make_event_datetime(schedule_year, month_name, day,
                                                       schedule_timezone)
                event = {
                    'summary': summary,
                    'description': description,
                    'start': {
                        'dateTime': start_dt,
                        'timeZone': schedule_timezone
                    },
                    'end': {
                        'dateTime': end_dt,
                        'timeZone': schedule_timezone
                    },
                    'colorId': event_color_id,
                    'location': schedule_location,
                    'reminders': {
                        'useDefault': False
                    },
                }

                try:
                    created_event = service.events().insert(
                        calendarId=calendar['id'], body=event).execute()
                    event_count += 1
                    print('Event #{0}: {1}'.format(
                        event_count, created_event.get('htmlLink')))
                    created_events.append(created_event)
                except HttpError as error:
                    sys.exit(f'An error occurred: {error}')

    log_file = os.path.join(os.path.dirname(xlsx_path),
                            'schedule_generated_events_log.json')
    with open('schedule_generated_events_log.json', 'w', encoding='utf-8') as f:
        print(f'Saving events log in {log_file}')
        from json import dump
        dump(created_events, f, ensure_ascii=False)

    print('https://calendar.google.com/calendar/embed?src={0}&ctz=Europe%2FWarsaw'.format(calendar['id']))

if __name__ == "__main__":
    if len(sys.argv) != 2 or not os.path.exists(sys.argv[1]):
        sys.exit('Path to .xls file is missing')

    main(sys.argv[1])
