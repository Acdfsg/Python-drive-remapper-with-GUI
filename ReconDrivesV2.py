"""
Drivemapping Wrapper:
    The following program acts as a wrapper to enable and obscure the drive mapping and removal process from endusers behind a GUI while keeping a 
    log of the last error found, if any.
    
    The current version of this program includes a flag-launch option "--not_clicked", which could be used in future as part of running the 
    Drivemapping Wrapper in a specified timeframe as part of a Windows service. This functionality has not been setup, and is pending approval.
    
    Program logic:
    User needs to reconnect a drive, or is having trouble with a drive, and launches the .exe:
        Program tests connection to [server IP]:
            Pass?
                GUI prompt with drive to select, prompt to attempt remapping of selected drive, prompt to remove the selected drive 
                (something to try as basic troubleshooting), prompt to see contact information for the helpdesk, and a prompt to quit.
                    Remap?
                        Attempt remapping for selected drive and notify user of result.
                    Remove?
                        Attempt removal of the selected drive and notify user of result.
                    Contact info?
                        New window with our Contact information.
                    Quit?
                        Exit program.
            Fail?
                GUI loop prompting user to troubleshoot and try again, contact helpdesk, or quit.
                    Try again?
                        Go back to testing connection to server [server IP],and proceed from there.
                    Contact helpdesk?
                        New window with our contact information.
                    Quit?
                        Exit program.
"""

import os
import subprocess
import argparse
import datetime
import csv
from tkinter import *
from tkinter import messagebox
import tkinter.font as tkfont

#Global variables, to be called from within functions that track or modify the current program state
states = ["testing","fail","pass"]
state = states[0]

dont_try = False #clunky varialble for stopping main and tkinter loops

def main(in_str):
    if in_str == "clicked":
        #Execute gui-logic and associated functions (User ran program)
        global states
        global state
        #perform a simple network test before doing anything else:
        while True:
            #while loop controls which gui-state
            if state == states[0]:
                network_check = simple_connection_test()
                if network_check == "server reachable":
                    state = states[2]
                else:
                    state = states[1]
            elif state == states[1]:
                global dont_try 
                if dont_try == False:
                    tk_main("fail")
                else:
                    #ensure user is still able to quit if they want!
                    break
            elif state == states[2]:
                tk_main("pass")
                #print(state)
                break #doesn't run until after tkinter window closes!
        #print(state) #runs after tkinter window closes! script continues!!!
    else:
        write_log("Program expected 'clicked' (no flags or additional input) and received --not_clicked or other unexpected input for main().")
        print("You've attempted to run this program using the optional command flag. This functionality is pending.")


def tk_main(pass_fail):
    #handle tkinter logic
    window = Tk() #start tkinter loop
    window.title("Reconnect Drives")
    app_font = tkfont.Font(family="Times New Roman",size=18)
    if pass_fail == "fail":
        user_info = Label(text="Failed to connect!  Please check your Internet and VPN connection and press 'Restart' to try again, \
select 'Help' for how to contact the helpdesk, or select 'Quit'.",font=app_font,wraplength=400).pack()
        state = states[0]
        retry_connection = Button(window,text="Restart",font=app_font,command=lambda: refresh_state(window)).pack()
        contact_helpdesk =Button(window,text="Help",font=app_font,command=lambda: get_help(app_font)).pack()
        exit_out = Button(window,text="Quit",font=app_font,command=lambda: no_re(window)).pack()
    if pass_fail == "pass":
        full_drives = pull_drive_details() #extract drive paths and drive letters
        drive_options = strip_driveletter(full_drives) #extract only drive letters
        user_info = Label(text="Please select the drive you wish to remap and 'Remap Drive', try 'Clear Drive' if you're having trouble, \
select 'Help' for how to contact the helpdesk, or select Quit to exit.",font=app_font,wraplength=400).pack()
        #define dropdown selection for enduser
        selected = StringVar() #tkinter requires specially type def
        selected.set(list(drive_options)[0]) #set default for drive mapping selection. 
        drive_selection = OptionMenu(window,selected,*drive_options.keys()) #create dropdown
        drive_selection.config(font=app_font) #need to set the button font specifically
        drive_menu = drive_selection.nametowidget(drive_selection.menuname) #select dropdown menu element
        drive_menu.config(font=app_font) #need to set the font of the options in menu specifically
        drive_selection.pack()
        #window interactables
        remap_prompt = Button(window,text="Remap Drive",font=app_font,command=lambda: remap_drives(app_font,drive_options,selected.get())).pack() #trigger attempt to remap drives
        wipe_known_drives = Button(window,text="Clear Drive",font=app_font,command=lambda: wipe_drive(app_font,selected.get())).pack() #wipe selected drive
        contact_helpdesk =Button(window,text="Help",font=app_font,command=lambda: get_help(app_font)).pack()
        exit_out = Button(window,text="Quit",font=app_font,command=window.destroy).pack() #close the window. CANNOT CALL destory(), or will close before open!!
    window.mainloop()


def no_re(main):
    #tkinter can't handle multiple actions on a button press unless calling 1 function. Set exit variable and kill tkinter loop here.
    global dont_try 
    dont_try = True
    main.destroy()


def get_help(cust_font):
    #information for calling us if big problems pop up
    new_window = Toplevel()
    contact_label = Label(new_window,text="<Insert helpdesk contact information here>",font=cust_font,wraplength=400).pack()


def wipe_drive(cust_font,drive_letter):
    #wipes drive, then pops up information for the user. If there is any failure, such as the drive having never been connected, a warning will
    #pop up instead
    new_window = Toplevel()
    try_remove = subprocess.run('net use /del ' + drive_letter,stdout=subprocess.PIPE)
    if try_remove.returncode == 0:
        status_label = Label(new_window,text="Drive "+drive_letter+" removed.",font=cust_font).pack()
    else:
        write_log("Failure to remove "+drive_letter + " drive.")
        status_error = Label(new_window,text="Drive "+drive_letter+" not removed! Are you sure you're connected to it? If so, \
consider contacting the helpdesk:\n",font=cust_font,wraplength=400).pack()


def remap_drives(cust_font,full_dict,drive_letter):
    new_window = Toplevel()
    to_assign = ""
    if drive_letter in full_dict.keys():
        #Uses drive letter to pull drive letter + path from dictionary
        to_assign = full_dict.get(drive_letter)
    if to_assign == "":
        write_log("Drive to assign returned empty!")
    add_drives = subprocess.run('net use ' + to_assign + ' /persistent:no',shell=True,stdout=subprocess.PIPE)
    #running with shell = True, can be corrected if this should be a security issue.
    #print(add_drives.returncode)
    if add_drives.returncode == 0: #returning text not working for some reason, so check for return code instead
        job_complete = Label(new_window,text="Drive "+drive_letter+" connected.",font=cust_font,wraplength=400).pack()
    else:
        write_log("Failure to map drive, likely lost connection to VPN or Internet after initial connection test.")
        job_complete = Label(new_window,text="Drive "+drive_letter+" failed to connect! \
Check your connection and try again, or call or email the helpdesk for assistance.",font=cust_font,wraplength=400).pack()


def refresh_state(main_window):
    #move out of tkinter logic and back to main function
    global state
    state = states[0] #ensure while loop starts at 0
    main_window.destroy() #exit tkinter loop


def strip_driveletter(in_list):
    char_list = {} #opted to go with a dictionary to maintain the letter drive dropdown relation
    #strip and return an array of drive letters for enduser convience
    for child in in_list:
        char_list[child[:2]] = child #add drive letter and everything needed for mapping to dictionary
    return char_list


def pull_drive_details():
    #takes input from file of 'driveletter: folderpath' lines, splits up the lines, and adds the lines into a list for a drop down
    paths = []
    path_source = open('C:\\mapdrive_path.txt','r') #path to a file of drives to map
    paths_raw = path_source.read() #automatically adds extra \ where needed
    for line in paths_raw.splitlines():
        paths.append(line)
    path_source.close()
    return paths


def simple_connection_test():
    #Test connection to server
    test = subprocess.run("ping [insert server IP] -n 1",stdout=subprocess.PIPE)
    test_str = test.stdout.decode('utf-8')
    positive = test_str.find("Lost = 0 (0% loss)") #prints location in str where match is found, or -1
    if positive > 0:
        return "server reachable"
    else:
        write_log(test_str)
        return "failure"


def write_log(last_error):
    #log designed only to retain the last error passed to it, along with current time if error not from Windows. 
    # Functions within this program are expected to fail individually or as a result of Windows, 
    # so extended logging is not currently seen as necessary. 
    raw_log = open('C:\\drive_mapping_errors.txt','w')
    raw_log.write(last_error + "\n" + str(datetime.datetime.now()))
    raw_log.close()


if __name__ == "__main__":
    #The following tests for additional arguments, which will be provided by the monitoring service, before running main(input)
    parser = argparse.ArgumentParser()
    #below checks for presence of --not_clicked. If present, store_true=true and if fires. If not present, false and else fires.
    #should allow running from service with flag vs user double-clicking .exe
    parser.add_argument("--not_clicked", help="use flag --not_clicked when setting up as a service",action="store_true")
    args = parser.parse_args()
    if args.not_clicked:
        main("not_clicked")
    else:
        main("clicked")
