# encoding: utf-8

import gc
import jpegdec
import ssl
import urllib
import usocket
import ujson
import time
import machine
import ntptime

from urllib import urequest

from ujson import load

gc.collect()

graphics = None
WIDTH = None
HEIGHT = None

rtc = machine.RTC()

SCOPE = "https://graph.microsoft.com/.default"

API_URL = ""

try:
    from secrets import API_MANDANT, API_SECRET, API_CLIENT, API_ROOM, SIGN_TITLE, DAY_LIGHT_SAVING
except ImportError:
    print("Create secrets.py with your O365 credentials")

# Length of time between updates in minutes.
# Frequent updates will reduce battery life!
#UPDATE_INTERVAL = 240
UPDATE_INTERVAL = 5

token_data = {"access_token": "", "token_type": "", "expires_in": "", "ext_expires_in": ""}

def sort_helper(e):
  return e['start_epoch']

def symbol_sanizer(string_to_update):   
    string_to_update = string_to_update.replace("\\u00f6", "ö")
    string_to_update = string_to_update.replace("\\u00e4", "ä")
    string_to_update = string_to_update.replace("\\u00fc", "ü")
    
    string_to_update = string_to_update.replace("\\u006d", "Ö")
    string_to_update = string_to_update.replace("\\u00c4", "Ä")
    string_to_update = string_to_update.replace("\\u00dc", "Ü")
    
    string_to_update = string_to_update.replace("\\u00e9", "é")
    
    string_to_update = string_to_update.replace("\\u0021", "!")  
    string_to_update = string_to_update.replace("\\u0022", '"')
    string_to_update = string_to_update.replace("\\u0023", "#")
    string_to_update = string_to_update.replace("\\u0024", "$")
    string_to_update = string_to_update.replace("\\u0025", "%")
    
    string_to_update = string_to_update.replace("\\u0026", '&')
    string_to_update = string_to_update.replace("\\u0027", "'")
    string_to_update = string_to_update.replace("\\u0028", "(")
    string_to_update = string_to_update.replace("\\u0029", ")")
    
    string_to_update = string_to_update.replace("\\u002a", '*')
    string_to_update = string_to_update.replace("\\u002b", "+")
    string_to_update = string_to_update.replace("\\u002c", ",")
    string_to_update = string_to_update.replace("\\u002d", "-")
    string_to_update = string_to_update.replace("\\u002e", ".")
    string_to_update = string_to_update.replace("\\u002f", "/")
    
    
    return string_to_update




def token_data_populate(toparsing_string):
    global token_data
    
    tmp_array = toparsing_string.strip("'{}'").split(',')
    
    for token_item in tmp_array:
        k,v = token_item.split(":")
        token_data[k.strip('"')] = v
     
    return token_data  

def http_get_buffered(url, headers, buffer_size=128):
    """
    Perform an HTTP GET request using raw sockets with buffered reading.
    Supports both HTTP and HTTPS.
    """    
    header_url = ""
    meta_information = ""
    meta_information_array = ""
    end_transmission = None
    start_json = ""
    end_json = ""
    
    try:
        for item in headers:
            header_url = header_url + item + ': ' + headers[item] + '\r\n'

    except Exception as e:
            print("Error:", e)
            return None
   
    try:
        # Parse URL
        proto, _, host, path = url.split('/', 3)
        path = '/' + path
        port = 443 if proto == 'https:' else 80

        # Resolve host
        addr_info = usocket.getaddrinfo(host, port)[0][-1]

        # Create socket
        sock = usocket.socket()
        sock.connect(addr_info)

        # Wrap in SSL if HTTPS
        if proto == 'https:':
            sock = ssl.wrap_socket(sock, server_hostname=host)

        # Send HTTP request
        request = "GET {} HTTP/1.1\r\nHost: {}\r\n{}Connection: close\r\n\r\n".format(path, host, header_url)
        print("REQUEST: ", request)
        #print(dir(sock))
        sock.write(request.encode())

        # Read response in chunks
        response_data = b""
        while True:
            chunk = sock.read(buffer_size)
            if not chunk:
                break
            response_data += chunk

        sock.close()
        # Decode byte date to string
        ret_response_data = response_data.decode()
        
        # Identifying metainformation like http return 200
        end_transmission = ret_response_data.find("Connection: close")    
        meta_information = ret_response_data[:end_transmission]
        meta_information_array = meta_information.split("\r\n")
               
        ret_response_data = ret_response_data[end_transmission:]
        
        start_json = ret_response_data.find("{")
        end_json = ret_response_data.find("}")
        
        ret_response_data = ret_response_data[start_json:]
        
        ret_response_data = ret_response_data[:-5]
        ret_response_data = ret_response_data.strip("\r")
        ret_response_data = ret_response_data.strip("\n")
      
        return meta_information_array, ret_response_data

    except Exception as e:
        print("Error:", e)
        return None





# Access Token holen
def get_access_token():
    global token_data
    token_url = f"https://login.microsoftonline.com/{API_MANDANT}/oauth2/v2.0/token"
    headers_content = {"Content-Type": "application/x-www-form-urlencoded"}
    body = (
        f"client_id={API_CLIENT}"
        f"&scope={SCOPE}"
        f"&client_secret={API_SECRET}"
        f"&grant_type=client_credentials"
    )
    try:
        print("Fordere Access Token an...")
        ret = urequest.urlopen(token_url, data=body, method="POST")
        if ret != None:
            token_data = token_data_populate(ret.readline().decode())
            ret.close()         
            return token_data
        else:
            print("Token-Fehler:", ret)
            ret.close()
            return None
    except Exception as e:
        print("Token-Anfrage fehlgeschlagen:", e)
        return None

def epoch_from_iso8601short(epoch_obj):
    time_tuple_obj = (int(epoch_obj[0:4]), int(epoch_obj[5:7]), int(epoch_obj[8:10]), int(epoch_obj[11:13]), int(epoch_obj[14:16]), int(epoch_obj[17:20]), 0, 0, 0)
    ret_epoch = time.mktime(time_tuple_obj)   
    return ret_epoch

def string_ast_odata_helper(string_data):
    # i dont can't load ast to simple parse my string
    # level1 is in my object context as @odata.context , [list of events] , "@odata.nextLink"
    return_data_test = {"@odata.context": "" , "value": "" , "@odata.nextLink": ""}
    return_data = {"@odata.context": "" , "value": "" , "@odata.nextLink": ""}
    next_link_tmp = ""

    value_tmp = ""
    calenda_data = []
    
    
    base_struct = string_data.split(",")
            
    return_data["@odata.context"] = base_struct[0].strip('{"@odata.context":')
    _, next_link_tmp = string_data.split('"@odata.nextLink":')
    return_data["@odata.nextLink"] = next_link_tmp[:-3]
    
    _ , value_tmp = string_data.split('"value":')
    value_tmp, _ = value_tmp.split("@odata.nextLink")
    value_tmp = value_tmp[:-2]
    
    # value_tmp3 hold the @odata.etag dicts
    string_to_strip = value_tmp.strip("\r")
    
    count_off_entry = len(value_tmp.split('"@odata.etag":'))
    
    value_tmp = value_tmp.split(',{"@odata.etag":')
    
    # Cleanup from odata crap that send by the database
    
    for item_stripes in value_tmp:
        entry = {}
        _, start_zeit = item_stripes.split('"start":{"dateTime":"')
        _, end_zeit = item_stripes.split('"end":{"dateTime":"')
        _, subject = item_stripes.split('"subject":"')
        subject, _ = subject.split('","bodyPreview":"')      
        
        entry["subject"] = symbol_sanizer(str(subject))      
        entry["start_zeit"] = start_zeit[:19]
        entry["end_zeit"] = end_zeit[:19]
        entry["start_epoch"] = epoch_from_iso8601short(start_zeit[:19])
        entry["end_epoch"] = epoch_from_iso8601short(end_zeit[:19])

        if len(entry["start_zeit"]) == 19 and len(entry["end_zeit"]) == 19 and type(entry["subject"]) == str :
            calenda_data.append(entry)
        else:
            print("Entry Validation Faild")
            return False 
     
    return_data["value"] = calenda_data           
    gc.collect()
   
    return True , return_data



# Gruppenkalender abrufen
def get_group_events(access_token):
    url = f"https://graph.microsoft.com/v1.0/users/{API_ROOM}/calendar/events"

    tmp_access_token = access_token.strip('"')
    headers = {
        "Authorization": f"Bearer {tmp_access_token}",
        "Content-Type": "application/json",
        "charset": "UTF-8"
    }
    data_dump = ""
    try:
        print("Rufe Gruppenkalender ab...")
        meta, data = http_get_buffered(url, headers , buffer_size=512)       
        gc.collect()
        
        print(meta[0])
        if meta[0] == 'HTTP/1.1 200 OK':
            print("Response received ({} bytes)".format(len(data)))
                                    
            ret_value , events = string_ast_odata_helper(data)
            if not ret_value:
                print("Keine Termine gefunden.")
                return False, "None"
            else:
                events = events["value"]
                gc.collect()
                return True, events
            
        else:
            print("API-Fehler:", "Connection Faild")
            gc.collect()
            return False, "Connection Error"
    except Exception as e:
        print("API-Anfrage fehlgeschlagen:", e)
        gc.collect()
        return False, "Critical Connection Error"
    

def sort_and_filter_events(events, time_frame):
    events_tmp = []
    current_meeting = None
    next_meeting = None
    sleep = False

    today_filter = f'{time_frame[0]}-{time_frame[1]}-{time_frame[2]}T'
    
    hours = time_frame[4]
    if time_frame[4] < 10:
        hours = f"0{time_frame[4]}"
    
    minutes = time_frame[5]
    if time_frame[5] < 10:
        minutes = f"0{time_frame[5]}"
    
    time_frame_filter = f'{hours}:{minutes}:00'

    events.sort(key=sort_helper)
    
    #Filter Today
    for ev in events:
        if ev['start_zeit'][:-8] == today_filter:
            events_tmp.append(ev)
    
    if len(events_tmp) == 0:
        sleep = True
           
    for index, ev in enumerate(events_tmp):
        start_zeit_epoch = epoch_from_iso8601short(ev['start_zeit'])
        end_zeit_epoch = epoch_from_iso8601short(ev['end_zeit'])
        if start_zeit_epoch < time.time():
            # Start Passed
            current_meeting = ev
            try:
                next_meeting = events_tmp[index+1]
            except:
                next_meeting = None
            
            # Event End Passed
            if end_zeit_epoch < time.time():
                current_meeting = None
    
    return sleep, current_meeting, next_meeting

def draw_frame(ret_time, current_meeting_fill, next_meeting_fill):
    gc.collect()
    
    # Apply an offset for the Inky Frame 5.7".
    if HEIGHT == 448:
        y_offset = 20
    # Inky Frame 7.3"
    elif HEIGHT == 480:
        y_offset = 35
    # Inky Frame 4"
    else:
        y_offset = 0

    # Draws the menu
    graphics.set_pen(1)
    graphics.clear()
    graphics.set_pen(0)

    graphics.set_pen(graphics.create_pen(0, 0, 255))
    graphics.rectangle(0, 0, WIDTH, 100)
    graphics.set_pen(1)
    title = SIGN_TITLE
    title_len = graphics.measure_text(title, 4) // 2
    graphics.text(title, (WIDTH // 2 - title_len), 10, WIDTH, 4)
    
    graphics.set_pen(1)
    
    day_light_houre = ret_time[4] + DAY_LIGHT_SAVING
    
    if day_light_houre < 10:
        day_light_houre = f"0{day_light_houre}"
    
    minutes = ret_time[5]
    if ret_time[5] < 10:
        minutes = f"0{ret_time[5]}" 
    
    date_time = f"{ret_time[2]}.{ret_time[1]}.{ret_time[0]} - {day_light_houre}:{minutes}"
    graphics.text(date_time, int(WIDTH*0.4) , HEIGHT - (380 + y_offset), 600, 2)


    graphics.set_pen(3)
    graphics.rectangle(30, HEIGHT - (300 + y_offset), WIDTH - 250, 200)
    graphics.set_pen(1)
    graphics.text("Aktuelles Meeting:", 35, HEIGHT - (280 + y_offset), 600, 2)
    if current_meeting_fill == None:
        graphics.text("Kein Meeting", 35, HEIGHT - (240 + y_offset), 600, 6)
        time_range_string = ""
        graphics.text(time_range_string , 35, HEIGHT - (180 + y_offset), 600, 2)        
    else:
        graphics.text(current_meeting_fill["subject"], 35, HEIGHT - (240 + y_offset), 600, 6)
        
        daylight_saving_hour = int(current_meeting_fill["start_zeit"][11:][:-6]) + DAY_LIGHT_SAVING
        daylight_saving_minutes = current_meeting_fill["start_zeit"][14:][:-3]
        
        daylight_saving_hour_end = int(current_meeting_fill["end_zeit"][11:][:-6]) + DAY_LIGHT_SAVING
        daylight_saving_minutes_end = current_meeting_fill["end_zeit"][14:][:-3]
                
        time_range_string = f' {daylight_saving_hour}:{daylight_saving_minutes} - {daylight_saving_hour_end}:{daylight_saving_minutes_end}'
        graphics.text(time_range_string , 35, HEIGHT - (180 + y_offset), 600, 2)


    graphics.set_pen(0)
    graphics.rectangle(30, HEIGHT - (100 + y_offset), WIDTH - 300, 50)
    graphics.set_pen(1)
    if next_meeting_fill == None:
        folgende_string = ""
        graphics.text(folgende_string, 35, HEIGHT - (85 + y_offset), 600, 2)        
    else:
        daylight_saving_hour = int(next_meeting_fill["start_zeit"][11:][:-6]) + DAY_LIGHT_SAVING
        daylight_saving_minutes = next_meeting_fill["start_zeit"][14:][:-3]
        
        daylight_saving_hour_end = int(next_meeting_fill["end_zeit"][11:][:-6]) + DAY_LIGHT_SAVING
        daylight_saving_minutes_end = next_meeting_fill["end_zeit"][14:][:-3]
                
        folgende_string = f'Folgendes: {next_meeting_fill["subject"]} --- {daylight_saving_hour}:{daylight_saving_minutes} - {daylight_saving_hour_end}:{daylight_saving_minutes_end}'
        graphics.text(folgende_string, 35, HEIGHT - (85 + y_offset), 600, 2)
    
    graphics.update()
    
def draw_frame_error():
    gc.collect()
    
    # Apply an offset for the Inky Frame 5.7".
    if HEIGHT == 448:
        y_offset = 20
    # Inky Frame 7.3"
    elif HEIGHT == 480:
        y_offset = 35
    # Inky Frame 4"
    else:
        y_offset = 0

    # Draws the menu
    graphics.set_pen(1)
    graphics.clear()
    graphics.set_pen(0)

    graphics.set_pen(graphics.create_pen(0, 0, 255))
    graphics.rectangle(0, 0, WIDTH, 100)
    graphics.set_pen(1)
    title = SIGN_TITLE
    title_len = graphics.measure_text(title, 4) // 2
    graphics.text(title, (WIDTH // 2 - title_len), 10, WIDTH, 4)   
    graphics.set_pen(1)

    graphics.set_pen(6)
    graphics.rectangle(30, HEIGHT - (300 + y_offset), WIDTH - 60, 200)
    graphics.set_pen(1)

    graphics.text("Verbindungsfehler", 65, HEIGHT - (240 + y_offset), 600, 6)

    graphics.update()
      
    
def time_update():
    # grab the current time from the ntp server and update the Pico RTC
    ret =  False
    
    try:
        ntptime.settime()
        ret = True
    except OSError:
        print("Unable to contact NTP server")

    current_t = rtc.datetime()

    return ret , current_t
    
