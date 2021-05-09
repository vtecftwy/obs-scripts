"""
Simple script to show a countdown in OBS

Starts and stops a timer countdown
Implemented by writing the countdown value twice per second in file timer.txt

To use the timer countdown, create a TXT (GDI+) source and configure it to get its text from the file timer.txt
in OBS assets.

Action Points:
to do: add script docstring
to do: refactor hotkey to create a single list/dict to configure all hotkeys and utility functions to (un)register them
to do: create the three color bg objects automatically so that it does not require preconfigures scene
to do: refactor color item so that the names of the items is based on those selected in script property
to do: create a timer.txt file in p2assets automatically. content: 00:00

Technical Reference:
https://obsproject.com/docs/scripting.html
https://github.com/obsproject/obs-studio/wiki/Getting-Started-With-OBS-Scripting

"""

import obspython as obs
import os
from datetime import datetime, time, timedelta
from pathlib import Path

durations = {'short': 5, 'long': 15}
active_duration = 'short'
color_bgs = {'normal': 'CLR: Title Green Bg', 'last_min': 'CLR: Title Orange Bg', 'elapsed': 'CLR: Title Red Bg'}

p2appdata = Path(os.environ['APPDATA'])
p2obs = p2appdata / 'obs-studio'
p2assets = p2obs/'assets-for-scripts'
os.makedirs(p2assets, exist_ok=True)

time_end = None
global_settings = None

scene_color_sources = []

# Set variables to handle OBS hotkeys
hotkey_id_start = None
HOTKEY_NAME_START = 'countdown_timer.start'
HOTKEY_DESC_START = 'Timer: Start New Countdown'

hotkey_id_end = None
HOTKEY_NAME_END = 'countdown_timer.end'
HOTKEY_DESC_END = 'Timer: Stop and Reset to 0'

hotkey_id_time_toggle = None
HOTKEY_NAME_TIME_TOGGLE = 'countdown_timer.time_toggle'
HOTKEY_DESC_TIME_TOGGLE = 'Timer: Toggle start time between short and long'


# ------------------------------------------------------------
# Global script functions call sequence for other operations
# https://github.com/obsproject/obs-studio/wiki/Getting-Started-With-OBS-Scripting#global-script-functions-call-sequence-for-other-operations
#
# Adding a script:
#
# Initialization steps like in OBS startup except that values of data settings are not available:
#
#   First script execution
#    1. script_defaults(settings)
#    2. script_description()
#    3. script_load(settings)
#    4. script_update(settings)
# Then, as the script is selected in the Scripts window, the properties are initialized and displayed:
#    5. script_properties()
#    6. Call to all "modified callbacks", for all properties (with data settings still not available)
#    7. script_properties() again
#    8. script_update(settings) with data settings available
#    9. Call to the "modified callbacks" of properties actually changed in previous steps
#
#  Removing a script just triggers a call to script_unload() (not script_save(settings)).
# ------------------------------------------------------------
def script_defaults(settings):
    """
    Technical Reference: https://obsproject.com/docs/reference-settings.html#obs-data-default-funcs
    """
    print('>> script_defaults')
    # todo: figure out how to get default values into GUI
    obs.obs_data_set_default_int(settings, 'timer_start_value_s', 5)
    obs.obs_data_set_default_int(settings, 'timer_start_value_l', 15)


def script_description():
    print('>> script_description')
    return """<h1><center>Timer Countdown</center></h1>
              <p><b>Start</b> and <b>stop</b> a timer countdown.</p>
              <p>Implemented by writing the countdown value twice per second in file <em>timer.txt</em></p>
              <p>To use the timer countdown, create a TXT (GDI+) source and configure it to get its text from the 
              file <em>timer.txt</em> in OBS assets.</p>
              """


def script_load(settings):
    print('>> script_load')
    global hotkey_id_start, hotkey_id_end, hotkey_id_time_toggle, global_settings

    # register all functions with hotkeys
    hotkey_id_start = register_and_load_hotkey(settings, HOTKEY_NAME_START, HOTKEY_DESC_START, start_countdown)
    hotkey_id_end = register_and_load_hotkey(settings, HOTKEY_NAME_END, HOTKEY_DESC_END, stop_countdown)
    hotkey_id_time_toggle = register_and_load_hotkey(settings, HOTKEY_NAME_TIME_TOGGLE, HOTKEY_DESC_TIME_TOGGLE,
                                                     toggle_initial_time)

    obs.obs_frontend_add_event_callback(on_frontend_load)
    global_settings = settings


def script_update(settings):
    # Create list of all color sources
    update_color_sources(settings)


def script_save(settings):
    print('>> script_save')
    save_hotkey(settings, HOTKEY_NAME_START, hotkey_id_start)
    save_hotkey(settings, HOTKEY_NAME_END, hotkey_id_end)
    save_hotkey(settings, HOTKEY_NAME_TIME_TOGGLE, hotkey_id_time_toggle)


def script_unload():
    print('>> script_unload')
    obs.obs_hotkey_unregister(start_countdown)
    obs.obs_hotkey_unregister(stop_countdown)
    obs.obs_hotkey_unregister(toggle_initial_time)


def register_and_load_hotkey(settings, name, description, callback):
    print('>> register_and_load_hotkey', name)
    hotkey_id = obs.obs_hotkey_register_frontend(name, description, callback)
    hotkey_save_array = obs.obs_data_get_array(settings, name)
    obs.obs_hotkey_load(hotkey_id, hotkey_save_array)
    obs.obs_data_array_release(hotkey_save_array)
    return hotkey_id


def update_color_sources(settings):
    global scene_color_sources
    print('>> update_color_sources')
    srcs = obs.obs_enum_sources()
    for src in srcs:
        if obs.obs_source_get_id(src) == "color_source_v3":
            scene_color_sources.append(src)


def on_property_change_callback(props, prop, settings):
    global durations
    print('>> on_property_change_callback')
    durations['short'] = obs.obs_data_get_int(settings, "timer_start_value_s")
    durations['long'] = obs.obs_data_get_int(settings, "timer_start_value_l")
    print('**Property changed: ', durations['short'], durations['long'])
    return True


def script_properties():
    """
    Technical note:
    1. Create props with obs.obs_properties_create(), to which all properties are added and then is returned

    For each property available in GUI:
    - create the property with obs.obs_properties_add_xxxx(props, ....). Catch pointer that is returned
    - create a callback function to catch changes made by user un GUI
    https://github.com/vtecftwy/OBS-Studio-Python-Scripting-Cheatsheet#additional-input

    """
    print('>> script_properties')
    props = obs.obs_properties_create()

    time_s = obs.obs_properties_add_int(props, "timer_start_value_s", "Timer Start Value Short (min)", 0, 59, 1)
    obs.obs_property_set_modified_callback(time_s, on_property_change_callback)

    time_l = obs.obs_properties_add_int(props, "timer_start_value_l", "Timer Start Value Long (min)", 0, 59, 1)
    obs.obs_property_set_modified_callback(time_l, on_property_change_callback)

    drop_list = obs.obs_properties_add_list(props, "color_normal", "Color Bg for Normal",
                                            obs.OBS_COMBO_TYPE_LIST,
                                            obs.OBS_COMBO_FORMAT_STRING)
    obs.obs_property_list_add_string(drop_list, "", "")
    for i, src in enumerate(scene_color_sources):
        obs.obs_property_list_add_string(drop_list, obs.obs_source_get_name(src), obs.obs_source_get_name(src))

    drop_list = obs.obs_properties_add_list(props, "color_last_minute", "Color Bg for Last Minute",
                                            obs.OBS_COMBO_TYPE_LIST,
                                            obs.OBS_COMBO_FORMAT_STRING)
    obs.obs_property_list_add_string(drop_list, "", "")
    for i, src in enumerate(scene_color_sources):
        obs.obs_property_list_add_string(drop_list, obs.obs_source_get_name(src), obs.obs_source_get_name(src))

    drop_list = obs.obs_properties_add_list(props, "color_elapsed", "Color Bg for Elapsed",
                                            obs.OBS_COMBO_TYPE_LIST,
                                            obs.OBS_COMBO_FORMAT_STRING)
    obs.obs_property_list_add_string(drop_list, "", "")
    for i, src in enumerate(scene_color_sources):
        obs.obs_property_list_add_string(drop_list, obs.obs_source_get_name(src), obs.obs_source_get_name(src))

    return props


def save_hotkey(settings, name, hotkey_id):
    print('>> save_hotkey', name)
    hotkey_save_array = obs.obs_hotkey_save(hotkey_id)
    obs.obs_data_set_array(settings, name, hotkey_save_array)
    obs.obs_data_array_release(hotkey_save_array)


def on_frontend_load(event):
    print('>> on_frontend_load')
    if event is not obs.OBS_FRONTEND_EVENT_FINISHED_LOADING:
        return

    print("frontend load")
    update_color_sources(global_settings)
    script_properties()


# ------------------------------------------------------------
# functions for handling timer countdown
# ------------------------------------------------------------
def update_countdown():
    # print('>> update_countdown')
    global time_end

    p2txt_file = p2assets / 'timer.txt'
    remaining_time = time_end - datetime.now()

    with open(p2txt_file, 'w') as f:
        if remaining_time.days < 0:
            # Timer value is negative
            change_visibility(color_bgs['normal'], False)
            change_visibility(color_bgs['last_min'], False)
            change_visibility(color_bgs['elapsed'], True)
            # change_visibility('CLR: Title Green Bg', False)
            # change_visibility('CLR: Title Orange Bg', False)
            # change_visibility('CLR: Title Red Bg', True)
            timer_txt = f"{0:02d}:{0:02d}"
            f.write(timer_txt)
            obs.timer_remove(update_countdown)
        else:
            mins, secs = divmod(remaining_time.seconds, 60)
            timer_txt = f"{mins:02d}:{secs:02d}"
            f.write(timer_txt)

            # Handle color background
            if mins < 1:
                # Last minute
                change_visibility(color_bgs['normal'], False)
                change_visibility(color_bgs['last_min'], True)
                change_visibility(color_bgs['elapsed'], True)
            else:
                # Normal countdown
                change_visibility(color_bgs['normal'], True)
                change_visibility(color_bgs['last_min'], True)
                change_visibility(color_bgs['elapsed'], True)


def start_countdown(pressed):
    global time_end

    if pressed:
        print('>> start_countdown')
        change_visibility('CLR: Title Green Bg', True)
        duration = timedelta(minutes=durations[active_duration])
        time_end = datetime.now() + duration
        update_countdown()
        obs.timer_add(update_countdown, 500)


def stop_countdown(pressed):
    if pressed:
        print('>> stop_countdown')
        p2txt_file = p2assets / 'timer.txt'
        change_visibility('CLR: Title Green Bg', True)
        with open(p2txt_file, 'w') as f:
            timer_txt = f"{0:02d}:{0:02d}"
            f.write(timer_txt)
        obs.timer_remove(update_countdown)


def change_visibility(src_name, is_visible=None):
    """Toggle or Set the item visibility based on the item's source name"""
    # print('change_visibility')
    current_scene = obs.obs_scene_from_source(obs.obs_frontend_get_current_scene())
    # to do: update this to focus on Timer scene, even if not the current scene

    scene_item = obs.obs_scene_find_source(current_scene, src_name)
    if is_visible is None:
        boolean = not obs.obs_sceneitem_visible(scene_item)
    else:
        boolean = is_visible
    obs.obs_sceneitem_set_visible(scene_item, boolean)


def toggle_initial_time(pressed):
    global active_duration
    if pressed:
        print('>> toggle_initial_time')
        if active_duration == 'short':
            active_duration = 'long'
        else:
            active_duration = 'short'
        print(active_duration)


if __name__ == '__main__':
    t = durations[active_duration]
    start_countdown()
