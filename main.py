# encoding: utf-8

import gc
import time
from machine import reset

import inky_helper as ih

sleep_mode_toggle = False

# Uncomment the line for your Inky Frame display size
# from picographics import PicoGraphics, DISPLAY_INKY_FRAME_4 as DISPLAY  # 4.0"
# from picographics import PicoGraphics, DISPLAY_INKY_FRAME as DISPLAY      # 5.7"
from picographics import PicoGraphics, DISPLAY_INKY_FRAME_7 as DISPLAY  # 7.3"

# Create a secrets.py with your Wifi details to be able to get the time
#
# secrets.py should contain:
# WIFI_SSID = "Your WiFi SSID"
# WIFI_PASSWORD = "Your WiFi password"

# A short delay to give USB chance to initialise
time.sleep(0.5)

# Setup for the display.
graphics = PicoGraphics(DISPLAY)
WIDTH, HEIGHT = graphics.get_bounds()
graphics.set_font("bitmap8")

# Turn any LEDs off that may still be on from last run.
ih.clear_button_leds()
ih.led_warn.off()

if ih.file_exists("state.json"):
    # Loads the JSON and launches the app
    ih.load_state()
    ih.launch_app(ih.state['run'])

    # Passes the the graphics object from the launcher to the app
    ih.app.graphics = graphics
    ih.app.WIDTH = WIDTH
    ih.app.HEIGHT = HEIGHT

try:
    from secrets import WIFI_SSID, WIFI_PASSWORD
    ih.network_connect(WIFI_SSID, WIFI_PASSWORD)
except ImportError:
    print("Create secrets.py with your WiFi credentials")

# Get some memory back, we really need it!
gc.collect()

# The main loop executes the update and draw function from the imported app,
# and then goes to sleep ZzzzZZz

file = ih.file_exists("state.json")

while True:
    global token_data
    global sleep_mode_toggle
    ih.led_warn.on()
    
    ret_time_bool, ret_time = ih.app.time_update()
    if ret_time_bool == True:
        token_data = ih.app.get_access_token()
        
        ret_value_group, events = ih.app.get_group_events(token_data['access_token'])
        gc.collect()
        if ret_value_group != False:
            sleep_mode, current_meeting, next_meeting = ih.app.sort_and_filter_events(events, ret_time)
            if sleep_mode_toggle == False:
                ih.app.draw_frame(ret_time,current_meeting,next_meeting)
            sleep_mode_toggle = sleep_mode
    else:
        ih.app.draw_frame_error()

    ih.led_warn.off()
    ih.sleep(ih.app.UPDATE_INTERVAL)

