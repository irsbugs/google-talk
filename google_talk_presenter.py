#!/usr/bin/env python3
#
# File: google_talk_presenter_gst.py
#
# Python script that will launch an impress presentation using the socket
# connection and then deliver the slide show. Text to speech is provided
# by google.
#
# Authors: Ian Stewart and Lawrence D'Oliveiro
# Release: 2017 Apr 05
# Test Platform: LibreOffice Version: 5.1.6., python 3.5.2, Ubuntu 16-04
# Version: 2.0
# Enhancements: V2. Clear out some of the development/text code.
# Presented at Hamilton Pytohn User Group meeting Monday, 8 May 2017
#
# TODO: Check program on Windows platform !
# TODO: Add sys.argv commands
#
# PREREQUISITES:
# 1. Ensure you have python3 installed.
# 2. Ensure that Python3-Uno-Bridge is installed. 
#    Check: >>> import uno
#           >>> dir(uno) ...should output about a list of about 50 items.
#    To install: sudo apt-get install uno-python3
#
# LAUNCHING THIS PROGRAM:
# 1. Ensure that there are no libreoffice/openoffce applications running.
# 2. Open a terminal window and enter the command:
#    soffice "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
# 3. Open another terminal window and run this python program:
#    $ python3 google_talk_presenter_v1.py
# 
# References on using pyuno bridge: 
# https://wiki.openoffice.org/wiki/Danny%27s_Python_Modules
# https://wiki.openoffice.org/wiki/Danny.OOo.DrawLib.py
# http://www.openoffice.org/udk/python/python-bridge.html
#
# Importing...
import os
import sys
import time
import subprocess
import urllib.parse
import urllib.request


import gi
gi.require_version('Gst', '1.0')
from gi.repository import GObject, Gst


try: 
    import uno
except ImportError:
    sys.exit("ImportError:\nInstall uno: sudo apt-get install python3-uno")
import unohelper

# initialize variables
text_file = "google_talk_presentation.txt"

mp3_player_list = ["mplayer", "ffplay"]
mp3_player = mp3_player_list[0]

default_language = "English"
default_language_code = "en"
slide_start = 0

# Languages that google can perform text to speech as of 2017-03-30
language_code_dict = {
'albanian': 'sq', 'arabic': 'ar', 'armenian': 'hy', 'basque': 'eu', 
'bosnian': 'bs', 'catalan': 'ca', 'chinese (simplified)': 'zh-CN', 
'chinese (traditional)': 'zh-TW', 'corsican': 'co', 'croatian': 'hr', 
'czech': 'cs', 'danish': 'da', 'dutch': 'nl', 'english': 'en', 
'esperanto': 'eo', 'finnish': 'fi', 'french': 'fr', 'german': 'de', 
'greek': 'el', 'hindi': 'hi', 'hungarian': 'hu', 'icelandic': 'is', 
'indonesian': 'id', 'italian': 'it', 'japanese': 'ja', 'khmer': 'km', 
'korean': 'ko', 'latin': 'la', 'latvian': 'lv', 'macedonian': 'mk', 
'nepali': 'ne', 'norwegian': 'no', 'polish': 'pl', 'portuguese': 'pt',
'romanian': 'ro', 'russian': 'ru', 'serbian': 'sr', 'sinhala': 'si', 
'slovak': 'sk', 'spanish': 'es', 'swahili': 'sw', 'swedish': 'sv', 
'tamil': 'ta', 'thai': 'th', 'turkish': 'tr', 'ukrainian': 'uk', 
'vietnamese': 'vi', 'welsh': 'cy'}

# Exit if python less than version 3
if sys.version_info.major < 3:
    string = ('\nPython3 is required. Please restart using Python3 \n'
              'Python {0} is not supported. Exiting...'
              #.format(sys.version_info.major))
              .format(sys.version[0:sys.version.find(" ")]))  
    sys.exit(string)

#------------------------------------------------------------------------------
#   Main procedural controlling function
#------------------------------------------------------------------------------
def main(control_dict, oDoc, oControl, mp3_player, slide_start=0):
    """
    Main function to control the flow of the program.
    The control_dict is a dictionary with keys from 0 to the number of 
    slides -1 to display.
    The value for each key is a list.
    Each list is a series of lists that comprize of a key and a value.
    Calls text_to_speech() and time.sleep() function
    Sends commands through pyuno bridge to Impress presentation

    Call the initialization and return player and loop
    For each item in the audio list call play_audio.
    """
    player, loop = initialize()

    for i in range(slide_start, len(control_dict)):
        slide_list = control_dict[i]
        for item in slide_list:
            if len(item) == 2:
                #print(item[0])
                key = item[0]
                #print(item[1])
                value = item[1]

                if key == "slide":
                    # Change slide
                    #print("Changing to next slide: {}".format(value-1))
                    oControl.gotoSlideIndex(value-1)
                    continue

                if key == "pause":   
                    time.sleep(value)
                    continue

                if key == "music":
                    # Value is a mp3 file name                     
                    #print("Music:", value)    
                    play_audio(value, player, loop)                    

                if len(key) == 2:
                    # key = language, value = paragraph of text. 
                    #text_to_speech(value, key, mp3_player)

                    #for message in message_list:
                    #if ".mp3" in message:
                    # pass the mp3 filename
                    #arg_list.append(message)
                    #else:
                    # Create the uri for google  
                    s = 'https://translate.google.com/translate_tts?'
                    s = s + 'ie=UTF-8&client=tw-ob&tl={}&q={}'
                    audio_string = s.format(key,value)
                    #print(audio_string, len(audio_string)-73)
                    play_audio(audio_string, player, loop) 
                    continue

                if key == "zh-TW" or key == "zh-CN":
                    # Two chinese languages have 5 x characters. 
                    #text_to_speech(value, key, mp3_player)
                    # Create the uri for google  
                    s = 'https://translate.google.com/translate_tts?'
                    s = s + 'ie=UTF-8&client=tw-ob&tl={}&q={}'
                    audio_string = s.format(key,value)
                    #print(audio_string)
                    play_audio(audio_string, player, loop) 
                    continue
    #oDoc.Presentation.dispose()


def initialize():
    """
    Initialize GObject.threads, Gst, player, loop and bus.
    Create a fakesink to bury any video.
    Bus is set up to perfom a call back to def bus_call(bus, message, loop)
    every time a playbin message is generated.
    """
    # Init
    GObject.threads_init()
    Gst.init(None)
    # Instantiate    
    player = Gst.ElementFactory.make("playbin", 'player')
    if not player:
        sys.stderr.write("'playbin' gstreamer plugin missing\n")
        sys.exit(1)
    fakesink = Gst.ElementFactory.make("fakesink", "fakesink")
    player.set_property("video-sink", fakesink)

    # Instantiate the event loop .
    loop = GObject.MainLoop()
    # Instantiate and initialize the bus call-back 
    bus = player.get_bus()
    bus.add_signal_watch()
    bus.connect ("message", bus_call, loop)

    return player, loop


def play_audio(audio_source, player, loop):
    # value, key,
    """
    Determine is audio source is uri for text-to speech. E.g.
    https://translate.google.com/translate_tts?ie=UTF-8&client=tw-ob&tl=
    en&q=Message one
    or convert mp3 filename to uri format. E.g. 
    file:///home/ian1/Desktop/audio/hello.mp3
    set playbin property and state to playing.
    Enter loop waiting for audio to finish.
    Set the state to Null. 
    """
    if Gst.uri_is_valid(audio_source):
        uri = audio_source
    else:
        uri = Gst.filename_to_uri(audio_source)

    player.set_property('uri', uri)

    # Start streaming the audio.
    player.set_state(Gst.State.PLAYING)
    # Loop while waiting for audio to finish. 
    loop.run()
    # On exiting loop() set playbin state to Null.
    player.set_state(Gst.State.NULL)


def bus_call(bus, message, loop):
    """
    Call back for messages generated when playbin is playing.
    The End-of-Stream, EOS, message indicates the audio is complete and the
    waiting loop is quit.
    """
    t = message.type
    if t == Gst.MessageType.EOS:
        # End-of-Stream therefore quit loop which executes playbin state Null
        #sys.stdout.write("End-of-stream\n")
        loop.quit()

    elif t == Gst.MessageType.ERROR:
        err, debug = message.parse_error()
        # TODO: If error then re-try last message
        sys.stderr.write("Error: %s: %s\n" % (err, debug))
        loop.quit()
    return True


#------------------------------------------------------------------------------ 
#   Launch functions - A collection of functions used before passing to main()
#------------------------------------------------------------------------------
def select_audio_player(mp3_player_list, mp3_player):
    """
    Provide selection of the mp3 player application to use.
    Based on items in mp3_player_list. Set mp3_player.
    """
    while True:
        print("Select the mp3 player application...")
        for index, item in enumerate(mp3_player_list):
            print("\t{}. {}".format(index, item))
        response = input("Select desired mp3 player [{}]: ".format(mp3_player))
        if response == "":
            response = mp3_player
            break
        else:
            try: 
                response = abs(int(response))
                mp3_player = mp3_player_list[response]
                break
            except:
                print("Invalid response. Please re-enter...")
                continue            
    print("MP3 data will be reproduced as audio with the application {}."
          .format(mp3_player))
    return mp3_player

 
def audio_test():
    """ 
    Test audio. 
    """
    while True:
        print("Audio level check...")   
        message = "Audio level check." # One, two, three. Check."
        text_to_speech(message, 'en')
        response = input("Audio level OK? [Yes]")
        if response == "":
            break
        else:
            if response[0].lower() in ["y", "1", "t"]:
                break
            else:
                continue

 
def connection_to_libreoffice():
    """Establish python connection to LibreOffice and return the desktop"""
    
    localContext = uno.getComponentContext()
				       
    resolver = localContext.ServiceManager.createInstanceWithContext(
	    "com.sun.star.bridge.UnoUrlResolver", localContext)
    
    smgr = resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager")

    # Alternative port 8100
    #smgr = resolver.resolve(
    #    "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager")
    
    remoteContext = smgr.getPropertyValue("DefaultContext")

    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",
            remoteContext)

    return desktop
        
    #Alternative variable naming:
    #mycontext = uno.getComponentContext()
    #resolver = mycontext.ServiceManager.createInstanceWithContext
    #    ("com.sun.star.bridge.UnoUrlResolver", mycontext)
    #myapi = resolver.resolve
    #    ("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")


def open_impress_document(desktop, presentation): 
    """Open an existing Impress document."""
    return desktop.loadComponentFromURL(presentation,"_blank", 0, () )
    

def read_text_file(text_file_path):
    """
    Read the text file for the control of the presentation.
    """
    try:
        with open(text_file_path,"r") as f:
            # Read to a list without the \n's at the end of lines.
            return f.read().splitlines()
    except FileNotFoundError as e:
        print("Attempted to open text file for the slide presentation.")
        print("FileNotFoundError: {}".format(e.strerror))
        print("The path or file is invalid: {}".format(e.filename))
        sys.exit("Exiting...")


def get_slide_show_filename(slide_data_list):
    """
    Get the path and filename for the presentation.
    The path is expected to not be supplied.
    The current working directory is used as a path.
    Be able to accept whitespace and uppercase.
    Skip comment lines and blank lines.
    """
    for item in slide_data_list:
        # Ignore blank and comment lines
        if len(item) == 0:
            continue
        if len(item) > 0 and item[0] == "#":
            continue

        if ("slide_show_file" in item and ":" in item and "[" in item and
             "]" in item):
            temp = item.strip(' \t\n\r[]')
            filename = temp.split(":")[1].strip()
            #print("|" + filename + "|")
            if filename == "":
                print("No file name supplied for [slide_show_file: ].")
                sys.exit("Exiting...")
            filename_path = "{}{}{}".format(os.getcwd(), os.sep, filename)
            # Test if file can be opened. Return file path and name if it can...
            try:
                with open(filename_path,"r") as f:
                    return filename, filename_path
            except FileNotFoundError as e:
                print("Attempted to open the slide presentation file.")
                print("FileNotFoundError: {}".format(e.strerror))
                print("The path or file is invalid: {}".format(e.filename))
                sys.exit("Exiting...")

    print("[slide_show_file: ] command not found." )
    sys.exit("Exiting...")


def check_slide_command(slide_data_list, max_slide, oDoc, text_file):
    """
    Knowing the total number of slides in the presentation.
    Check the [slide:x] commands are valid. i.e. Is an integer in range.
    If error, provide error message and exit.
    Some slides may be displayed more that once.
    Return the number of slides that will be displayed.
    """
    line_number = 0
    counter = 0
    for item in slide_data_list:
        line_number +=1
        # Ignore blank lines
        if len(item) == 0:
            continue
        # Ignore comment and text. In case command found in the comments.
        if not item[0] == "[":
            continue

        if "[slide:" in item.lower() and "]" in item:
            temp = item.strip(' \t\n\r[]')
            slide_number = temp.split(":")[1].strip()
            #print("|" + slide_number + "|")
            try:
                slide_number = int(slide_number)
            except ValueError as e:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("Slide number {} is not an integer".format(slide_number))
                oDoc.dispose()
                sys.exit("Exiting...")

            if slide_number > max_slide:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("Slide number {} exceeds total slides of {}"
                      .format(slide_number, max_slide))
                oDoc.dispose()
                sys.exit("Exiting...")

            if slide_number < 1:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("Slide number {} is less than first slide number of 1."
                      .format(slide_number))
                oDoc.dispose()
                sys.exit("Exiting...") 

            counter +=1
    return counter


def check_music_command(slide_data_list, max_slide, oDoc, text_file):
    """
    Knowing the total number of slides in the presentation.
    Check the [music:x] commands are valid. i.e. Is a file with mp3 extension.
    If error, provide error message and exit.
    Some slides may be displayed more that once.
    Return the number of slides that will be displayed.
    """
    line_number = 0
    counter = 0
    for item in slide_data_list:
        line_number +=1
        # Ignore blank lines
        if len(item) == 0:
            continue
        # Ignore comment and text. In case command found in the comments.
        if not item[0] == "[":
            continue

        if "[music:" in item.lower() and "]" in item:
            temp = item.strip(' \t\n\r[]')
            filename = temp.split(":")[1].strip()
            #print("|" + slide_number + "|")
            #filename = hello.mp3
            filename_list = filename.split(".")
            #print(filename_list)
            #print(filename_list[-1])

            try:
                # Try opening the file
                f = open(filename, "r")
                f.close()

            except ValueError as e:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("music file {} is not valie".format(slide_number))
                oDoc.dispose()
                sys.exit("Exiting...")

            if filename_list[-1] != "mp3" and filename_list[-1] != "wav":
                print("\nError in file {} at line number: {}"
                      .format(filename, line_number))
                print("{}".format(item))
                print("Music file {} does not have .mp3 or .wav extension"
                      .format(filename))
                oDoc.dispose()
                sys.exit("Exiting...")

            counter +=1
    return counter


def check_pause_command(slide_data_list, max_slide, oDoc, text_file):
    """
    Check the [pause:x] commands are valid. i.e. Is an integer in range.
    If error, provide error message and exit.
    Return the number of pause commands.
    """
    line_number = 0
    counter = 0
    for item in slide_data_list:
        line_number +=1
        # Ignore blank lines
        if len(item) == 0:
            continue
        # Ignore comment and text. In case command found in the comments.
        if not item[0] == "[":
            continue

        if "[pause:" in item.lower() and "]" in item:
            temp = item.strip(' \t\n\r[]')
            pause_value = temp.split(":")[1].strip()
            #print("|" + pause_value + "|")
            try:
                pause_value = float(pause_value)
            except ValueError as e:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("Pause value {} is not a float".format(pause_value))
                oDoc.dispose()
                sys.exit("Exiting...")

            if pause_value < 0:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("Pause {} is negative value."
                      .format(pause_value))
                oDoc.dispose()
                sys.exit("Exiting...") 

            counter +=1
    return counter


def check_language_command(slide_data_list, language_code_dict, oDoc,
                           text_file):
    """
    Check the [language:x] commands are valid.
    If error, provide error message and exit.
    Return the number of language commands.
    """
    line_number = 0
    counter = 0
    for index, item in enumerate(slide_data_list):
        line_number +=1
        # Ignore blank lines
        if len(item) == 0:
            continue
        # Ignore comment and text. In case command found in the comments.
        if not item[0] == "[":
            continue

        if "[language:" in item.lower() and "]" in item:
            temp = item.strip(' \t\n\r[]')
            language_string = temp.split(":")[1].strip()
            #print("|" + pause_value + "|")
            try:
                # try to get the language code. If so, insert code in list
                language_code = language_code_dict[language_string]
                language_pop = slide_data_list.pop(index)
                temp = "[language:{}]".format(language_code)
                slide_data_list.insert(index, temp )

            except KeyError:
                print("\nError in file {} at line number: {}"
                      .format(text_file, line_number))
                print("{}".format(item))
                print("Language of '{}' is not valid".format(language_string))
                oDoc.dispose()
                sys.exit("Exiting...")
  
            counter +=1
    return counter, slide_data_list


def built_control_dict_template(slide_total_displayed):
    """
    Build the template for the control dictionary.
    Dictionary has a integer key for each slide displayed 
    {0: [], 1: [], 2: [], 3: []}
    """
    control_dict = {}
    for i in range(slide_total_displayed):
        control_dict.update({i: []})
    return control_dict
    

def built_control(slide_data_list, control_dict, slide_total_displayed, oDoc):
    """
    Insert the control data and text into the control_dict.
    For each slide in the dictionary, use lists within lists.
    Key value lists will be like this...
    [['slide':1], ['en': 'This is slide 1'], ['pause':2], ['fr': 'bonjour']...] 
    A paragraph of text may span multiple lines. A blank line or a command 
    terminates the paragraph.
    """
    line_number = 0
    counter = -1
    boo_text =  False
    language_code = "en"
    temp_text = ""
    for index, item in enumerate(slide_data_list):
        line_number +=1
        # Clear spaces from front and rear of line
        # When determined to be text then add a space at end of line.
        item = item.strip(' \t\n\r')

        if len(item) == 0:
            if boo_text:    
                # update the text with language code as key.
                control_dict[counter].append([language_code, temp_text])
                boo_text = False
                temp_text = ""
                continue

            else:
                #Extra blank line - Do nothing"
                continue
        
        # Lines that have content.
        # Ignore comment lines
        if item[0] == "#":
            #print(item)
            continue
        
        # Slide command
        if item[0] =="[" and "[slide:" in item.lower() and "]" in item:
            if boo_text:    
                # update the text with language code as key.
                control_dict[counter].append([language_code, temp_text])
                boo_text = False
                temp_text = ""
            temp = item.strip(' \t\n\r[]')
            slide_number = int(temp.split(":")[1].strip())
            # Update the main key of the control_dict
            counter +=1
            control_dict[counter].append(['slide', slide_number])
            continue

        # music command
        if item[0] =="[" and "[music:" in item.lower() and "]" in item:
            if boo_text:    
                # update the text with language code as key.
                control_dict[counter].append([language_code, temp_text])
                boo_text = False
                temp_text = ""
            temp = item.strip(' \t\n\r[]')
            filename = temp.split(":")[1].strip()
            # Update the main key of the control_dict
            #print(filename)
            control_dict[counter].append(['music', filename])
            continue

        # Pause command
        if item[0] =="[" and "[pause:" in item.lower() and "]" in item:
            if boo_text:    
                # update the text with language code as key.
                control_dict[counter].append([language_code, temp_text])
                boo_text = False
                temp_text = ""
            temp = item.strip(' \t\n\r[]')
            pause_value = float(temp.split(":")[1].strip())
            control_dict[counter].append(['pause', pause_value])
            continue

        # Language command
        if item[0] =="[" and "[language:" in item.lower() and "]" in item:
            if boo_text:    
                # update the text with language code as key.
                control_dict[counter].append([language_code, temp_text])
                boo_text = False
                temp_text = ""
            temp = item.strip(' \t\n\r[]')
            language_code = temp.split(":")[1].strip()
            #print("Language code is now: {}".format(language_code))
            continue           
    
        # slide_show_file command
        if item[0] =="[" and "[slide_show_file:" in item.lower():
            # Ignore this command.
            continue

        # Miscellaneous. Junked command. 
        if item[0] =="[" and "]" in item:
            # Ignore this command.
            continue

        # Item must be text. Build/Append text string.       
        temp_text = temp_text + item + " "
        boo_text = True
        
    return control_dict

#------------------------------------------------------------------------------
# OLD To be deleted. This is the version using urllib.
#------------------------------------------------------------------------------
#   Audio 
#------------------------------------------------------------------------------
def text_to_speech(message='Hello World', language='en', mp3=mp3_player):
    """
    Use google translate to do text to speech translation.
    Use mplayer to play the mp3 data.
    message = text to be converted to speech
    language = en is English, fr is French, de is German, etc.
    """
    # Build the url string.
    url = 'https://translate.google.com/translate_tts'
    user_agent = 'Mozilla'
    values = {'tl' : language,
              'client' : 'tw-ob',
              'ie' : 'UTF-8',
              'q' : message }
    data = urllib.parse.urlencode(values)
    headers = { 'User-Agent' : user_agent }

    req = urllib.request.Request(url + "?" + data, None, headers)

    # Select the mp3 player to use...
    if mp3 == "mplayer":
        player = subprocess.Popen \
          (
            args = ("mplayer", "-cache", "1024", "-really-quiet", "/dev/stdin"),
            stdin = subprocess.PIPE
          )
    if mp3 == "ffplay":
        player = subprocess.Popen \
          (
            args = ("ffplay", "-autoexit", "-"),
            stdin = subprocess.PIPE
          )   

    # Send the request to google, and send mp3 data to mp3 player.
    try: 
        with urllib.request.urlopen(req) as response:
            mp3_data = response.read()
            player.stdin.write(mp3_data)

    except urllib.error.URLError as e:
        print(e)
        print(e.reason)
        print(e.read())

    player.stdin.close()
    player.wait() # fixme: should check return status

     
#------------------------------------------------------------------------------ 
#   Start
#------------------------------------------------------------------------------ 
if __name__ == "__main__":
    """
    Select the mp3 application. TODO: Pass as sys.argv[]
    Test audio level.
    Call function to establish connection to LibreOffice
    Check the text/control file exists.
    Read text/control to a list without the \n's at the end of lines.
    Get slide show default path and filename. Check file exists.
    Open Impress slide show and return object oDoc.
    Display the slide show default language.
    Provide information about the slide show. No of slides.
    Check the slide commands have valid values
    Check pause commands have a valid amount.
    Check language commands are valid.
    Initialize the dictionary to control the presentation.
    Load data into the dictionary to control the presentation.
    Start the slide show and instantiate the presentation control object
    Call the main() function to run the slide show.
    Dispose of the slide show after it has finnished.
    """

    child = subprocess.Popen(args = ("soffice", "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"))

    # TODO: Add sys.argv command line to input the control and text file.
    if len(sys.argv) > 1:
        slide_start = int(sys.argv[1])
 
    # select the mp3_player to use.
    #mp3_player = select_audio_player(mp3_player_list, mp3_player)
    mp3_player = "mplayer"
    # TODO: Remove mp3_player selection

    # Provide and audio test to set the volume.
    audio_test()

    # Establish connection to LibreOffice. 
    oDesktop = connection_to_libreoffice()

    # Open the text/control file. Check for file not found.
    text_file_path = "{}{}{}".format(os.getcwd(), os.sep, text_file)
    #print(text_file_path)

    # Read text/control to a list without the \n's at the end of lines.
    slide_data_list = read_text_file(text_file_path)
    #print(slide_data_list)
  
    # Get slide show default path and filename and Check file exists. e.g.
    # [slide_show_file:slide_test_v1.odp]
    impress_file, impress_file_path = get_slide_show_filename(slide_data_list)
    print("Impress presentation: {}".format(impress_file_path))
 
    # Open existing Impress slide show and return object oDoc.
    presentation_url = "file:///{}".format(impress_file_path)
    oDoc = open_impress_document(oDesktop, presentation_url)
    # If there are errors after this point, dispose of the Impress document.
    # oDoc.dispose()

    # Display the slide show default language.
    print("Default language code: {}".format(default_language_code))

    # Provide information about the slide show. No of slides.
    total_slide = oDoc.DrawPages.Count
    print("Total slides in {}: {}".format(impress_file, total_slide))

    # Check the slide commands are valid for the slides in the presentation.
    # Count the total calls for slides to be shown.
    # Of all the slides some may not be shown others may be shown twice, etc.
    slide_total_displayed = check_slide_command(slide_data_list, total_slide,
                                                oDoc, text_file)
    print("Total slides to be displayed: {}".format(slide_total_displayed))

    music_total = check_music_command(slide_data_list, total_slide, oDoc,
                                      text_file)
    print("Total music files to be played: {}".format(music_total))

    # Check pause commands have a valid amount. Integer or float.
    pause_total = check_pause_command(slide_data_list, total_slide, oDoc,
                                      text_file)
    print("Total Pause commands: {}".format(pause_total))

    # Check language commands are valid.
    # Convert to language code: '[language:French]' becomes '[language:fr]'
    # All use two letter code en, fr, etc., except chinese zh-TW, zh-CN.
    # Update: slide_data_list with modified language commands.
    #print(slide_data_list)
    language_total, slide_data_list = check_language_command(slide_data_list,
                                                          language_code_dict,
                                                          oDoc,
                                                          text_file)        
    print("Total Language commands: {}".format(language_total))
    #print(slide_data_list)
    # Command data has been verified as OK. 

    # Build the dictionary to control the presentation.
    control_dict = built_control_dict_template(slide_total_displayed)
    #print(control_dict) #{0: [], 1: [], 2: [], 3: []}

    control_dict = built_control(slide_data_list, control_dict,
                                 slide_total_displayed, oDoc)
    #print(control_dict)

    #response = input("Paused. Hit return to start slide show.")
    # Start the slide show and instantiate the presentation control object
    oDoc.Presentation.start()
    # Allow time to launch
    time.sleep(3)
    oControl = oDoc.Presentation.getController()

    print("Slide Show is running: {}".format(oDoc.Presentation.isRunning()))



    # Call the main() function to run the slide show.
    main(control_dict, oDoc, oControl, mp3_player, slide_start)

    # Dispose of the slide show after it has finnished.
    oDoc.dispose()
    child.kill()
    sys.exit()


