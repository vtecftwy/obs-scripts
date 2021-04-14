# Script to handle PPT Slideshow commands from within OBS
#
# Downloaded from https://obsproject.com/forum/resources/powerpoint-slide-window-navigation-using-obs-hotkey.938/
# Modified to correct ssv = ssw[0].View which fails + additional improvements
#
# Technical References:
# MS Component Object Model (COM) https://docs.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal
# Powerpoint Object Model: https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint/object-model
#

# todo: debug following items:
# 1.
# Upon entering HOME key to open a presentation
# [ppt_slide.py] Starting script
# [ppt_slide.py] Navigate Powerpoint Slides from within OBS.
# [ppt_slide.py] entering get_slidesshow_view
# [ppt_slide.py] Application Name:  Not Opened
# [ppt_slide.py] Opening MS PPT
# [ppt_slide.py] Opening a  presentation from file dialog
# [ppt_slide.py] user:  etien
# [ppt_slide.py] user_profile:  C:\Users\etien
# [ppt_slide.py] path:  D:\etien\Documents\*
# [ppt_slide.py] Traceback (most recent call last):
# [ppt_slide.py]   File "D:/PyProjects/obs/scripts\ppt_slide.py", line 168, in slideshow_view_first
# [ppt_slide.py]     ssv = get_slideshow_view()
# [ppt_slide.py]   File "D:/PyProjects/obs/scripts\ppt_slide.py", line 154, in get_slideshow_view
# [ppt_slide.py]     target_ppt = target_ppt % prez_count
# [ppt_slide.py] ZeroDivisionError: integer division or modulo by zero



import obspython as obs
import win32com.client as win32

powerpoint = None
target_ppt = 1


hotkey_id_frst = None
HOTKEY_NAME_FRST = 'powerpoint_slides.first'
HOTKEY_DESC_FRST = 'Go to the first slide of active Powerpoint Presentation.'

hotkey_id_prev = None
HOTKEY_NAME_PREV = 'powerpoint_slides.previous'
HOTKEY_DESC_PREV = 'Go to the previous slide of active Powerpoint Presentation.'

hotkey_id_next = None
HOTKEY_NAME_NEXT = 'powerpoint_slides.next'
HOTKEY_DESC_NEXT = 'Go to the next slide of active Powerpoint Presentation.'

hotkey_id_last = None
HOTKEY_NAME_LAST = 'powerpoint_slides.last'
HOTKEY_DESC_LAST = 'Go to the last slide of active Powerpoint Presentation.'

hotkey_id_next_ppt = None
HOTKEY_NAME_NEXT_PPT = 'powerpoint_slides.switch_ppt'
HOTKEY_DESC_NEXT_PPT = 'Switch to the next opened presentation.'


# ------------------------------------------------------------
# global functions for script plugins
# ------------------------------------------------------------
def script_load(settings):
    global hotkey_id_frst
    global hotkey_id_prev
    global hotkey_id_next
    global hotkey_id_last
    global hotkey_id_next_ppt

    hotkey_id_frst = register_and_load_hotkey(settings, HOTKEY_NAME_FRST, HOTKEY_DESC_FRST, slideshow_view_first)
    hotkey_id_prev = register_and_load_hotkey(settings, HOTKEY_NAME_PREV, HOTKEY_DESC_PREV, slideshow_view_previous)
    hotkey_id_next = register_and_load_hotkey(settings, HOTKEY_NAME_NEXT, HOTKEY_DESC_NEXT, slideshow_view_next)
    hotkey_id_last = register_and_load_hotkey(settings, HOTKEY_NAME_LAST, HOTKEY_DESC_LAST, slideshow_view_last)
    hotkey_id_next_ppt = register_and_load_hotkey(settings, HOTKEY_NAME_NEXT_PPT, HOTKEY_DESC_NEXT_PPT, switch_to_next_ppt)


def script_unload():
    obs.obs_hotkey_unregister(slideshow_view_first)
    obs.obs_hotkey_unregister(slideshow_view_previous)
    obs.obs_hotkey_unregister(slideshow_view_next)
    obs.obs_hotkey_unregister(slideshow_view_last)
    obs.obs_hotkey_unregister(switch_to_next_ppt)


def script_save(settings):
    save_hotkey(settings, HOTKEY_NAME_FRST, hotkey_id_frst)
    save_hotkey(settings, HOTKEY_NAME_PREV, hotkey_id_prev)
    save_hotkey(settings, HOTKEY_NAME_NEXT, hotkey_id_next)
    save_hotkey(settings, HOTKEY_NAME_LAST, hotkey_id_last)
    save_hotkey(settings, HOTKEY_NAME_NEXT_PPT, hotkey_id_next_ppt)


def script_description():
    return 'Navigate Powerpoint Slides from within OBS.'


def script_defaults(settings):
    obs.obs_data_set_default_int(settings, 'interval', 10)
    obs.obs_data_set_default_string(settings, 'source', '')


def register_and_load_hotkey(settings, name, description, callback):
    """Utility function to register a hotkey"""
    hotkey_id = obs.obs_hotkey_register_frontend(name, description, callback)
    hotkey_save_array = obs.obs_data_get_array(settings, name)
    obs.obs_hotkey_load(hotkey_id, hotkey_save_array)
    obs.obs_data_array_release(hotkey_save_array)
    return hotkey_id


def save_hotkey(settings, name, hotkey_id):
    """Utility function to save a hotkey"""
    hotkey_save_array = obs.obs_hotkey_save(hotkey_id)
    obs.obs_data_set_array(settings, name, hotkey_save_array)
    obs.obs_data_array_release(hotkey_save_array)


print('Starting script')
print(script_description())


# ------------------------------------------------------------
# functions for handling powerpoint
# ------------------------------------------------------------
def open_file_with_dialog(app):
    """https://docs.microsoft.com/en-us/office/vba/api/powerpoint.application.filedialog"""
    # to do: update docstring

    # Define used MSO constants
    msoFileDialogOpen = 0x1
    msoFileDialogFilePicker = 0x3

    objShell = win32.Dispatch("WScript.Shell")

    main_drive = 'D:'
    user = objShell.ExpandEnvironmentStrings("%UserName%")
    user_profile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
    path2initialfolder = f"{main_drive}\\{user}\\Documents\\*"

    dlgOpen = app.FileDialog(msoFileDialogOpen)
    dlgOpen.AllowMultiselect = False
    dlgOpen.InitialFileName = path2initialfolder
    dlgOpen.Show()
    dlgOpen.Execute()

    print('user: ',user)
    print('user_profile: ', user_profile)
    print('path: ', path2initialfolder)


def get_slideshow_view():
    """Retrieve the view for the active slideshow"""
    # to do: update docstring
    global powerpoint
    global target_ppt
    print('entering get_slidesshow_view')
    print('Application Name: ', powerpoint.Name if powerpoint is not None else 'Not Opened')

    if powerpoint is None:
        print('Opening MS PPT')
        powerpoint = win32.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

    prez_count = powerpoint.Presentations.Count
    if prez_count == 0:
        print('Opening a  presentation from file dialog')
        open_file_with_dialog(powerpoint)

    if target_ppt > prez_count:
        target_ppt = target_ppt % prez_count
    prez = powerpoint.Presentations.Item(target_ppt)
    print(f"Active PPT is: {prez.Name}")

    # Check that a slideshow is opened for the active presentation, else, run
    if powerpoint.SlideShowWindows.Count == 0:
        print(f'No slideshow currently opened for {prez.Name}. Starting one.')
        prez.SlideShowSettings.Run()

    return prez.SlideShowWindow.View


def slideshow_view_first(pressed):
    if pressed:
        ssv = get_slideshow_view()
        if ssv:
            ssv.First()


def slideshow_view_previous(pressed):
    if pressed:
        ssv = get_slideshow_view()
        if ssv:
            ssv.Previous()


def slideshow_view_next(pressed):
    if pressed:
        ssv = get_slideshow_view()
        if ssv:
            ssv.Next()


def slideshow_view_last(pressed):
    if pressed:
        ssv = get_slideshow_view()
        if ssv:
            ssv.Last()


def switch_to_next_ppt(pressed):
    global target_ppt
    if pressed:
        print('Entering switch_to_next_ppt')
        # Close current slideshow
        ssv = get_slideshow_view()
        if ssv:
            ssv.Exit()

        # Increment target_ppt
        target_ppt = target_ppt + 1
        ssv = get_slideshow_view()
        if ssv:
            ssv.First()


if __name__ == '__main__':
    get_slideshow_view().First()
    get_slideshow_view().Last()
    switch_to_next_ppt(1)
