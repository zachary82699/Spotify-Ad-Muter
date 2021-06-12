#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import time
import psutil
import win32gui

from subprocess import Popen
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# mute process 
def mute(process_name):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and session.Process.name() == process_name:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(1, None)


# unmute process
def unmute(process_name):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and session.Process.name() == process_name:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(0, None)


# Line clear
def line_clear():
    sys.stdout.write('\033[2K\033[1G')
    return


# check for if a process is currently running
def is_process_running(processName):
    # Iterate over the all the running process
    for proc in psutil.process_iter():
        try:
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False


# Check for the specific titles. works better than ctypes
def find_window(title):
    hwnd = win32gui.FindWindowEx(0, 0, 0, title)
    return hwnd  # returns 0 if nothing found


# Platform safe clear console method
def clear():
    if os.name == 'nt':
        Popen('cls', shell=True)
    else:
        Popen('clear', shell=True)
        

# Main entry point of the program
def main():
    
    clear()
    
    Popen('title Spotify Ad Muter', shell=True)
    Popen('mode con cols=50 lines=20', shell=True)
    
    while True:
        ad_titles = [
            # some ads set these titles
            r"Spotify Free",
            r"Spotify",
            # ads that needed manually added to list
            r"Advertisement",
            r"😈 Become a Blue Devil 😈",
            r"Sign then Drive is Back!",
            r"An Alternative To A College",
            r"Nominate Your Company",
            r"We`re on a path to better healthcare.",
            r"You're welcome here.",
            r"Online MBA, Nursing, Education",
            r"Cedar Fair",
            r"Become a Montessori teacher",
            r"Fall 2021 Classes Enrolling Now"
        ]
        ad_detected = False

        for title in ad_titles:

            time.sleep(0.1)

            if find_window(title):
                mute('Spotify.exe')

                line_clear()
                print(r'>> ads were detected or spotify is paused! <<',
                      end='\r', flush=True)
                ad_detected = True
                continue

            if ad_detected:
                break

        else:
            unmute('Spotify.exe')

            line_clear()
            print(r'>> no ads detected! <<', end='\r', flush=True)

            ad_detected = False
            
            if is_process_running('Spotify'):
                continue
            else:
                print(r'>> spotify is closed! <<', end='\r', flush=True)


if __name__ == "__main__":
    main()
