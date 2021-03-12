import time 
from pathlib import Path
import logging
from subprocess import Popen
from io import BytesIO

import keyboard
import win32clipboard as cb
from PIL import Image
from selenium import webdriver
from pyOpenRPA.Robot import UIDesktop

from settings import SEARCH_PHRASE 

logger = logging.getLogger(__name__) # create separate logger for current module

chrome_exe_path = r'..\resources\GoogleChromePortable\App\Chrome-bin\chrome.exe'
chrome_driver_path = r'..\resources\SeleniumWebDrivers\Chrome\chromedriver_win32 v84.0.4147.30\chromedriver.exe'
screenshot_dir_path = r'..\docs'


def init_webdriver(chrome_exe_path, chrome_driver_path = None): # launch Chrome through Selenium
    # Set full path to exe of the chrome
    options_instance = webdriver.ChromeOptions() # create object options_instance of ChromeOptions() Class
    options_instance.binary_location = chrome_exe_path # allocate attribute .binary_location to options_instance object

    # Run chrome instance
    driver_instance = None
    if chrome_driver_path:
        # Initialize webdriver with specified webdriver path
        driver_instance = webdriver.Chrome(executable_path = chrome_driver_path, options = options_instance) # initialize webdriver
    else:
        driver_instance = webdriver.Chrome(options = options_instance) # webdriver.chrome() should look for path to chromedriver.exe by itself
    # Return the result
    return driver_instance


def extract_search_results(driver, request):

    driver.get(request) # Open the URL
    driver.implicitly_wait(15) # wait for 15 sec for results to load
    
    found_elements = driver.find_elements_by_class_name('serp-item') # class name 'serp-item' dedicated to search results
         
    results = []
    for i, element in enumerate(found_elements):

        title_link_element = None
        title_text_element = None

        try:
            title_link_element = element.find_element_by_class_name('OrganicTitle-Link').get_attribute('href')
            title_text_element = element.find_element_by_class_name('OrganicTitle-LinkText').text.replace('\n','')
        except Exception as e: # ensure the program will proceed even if one of result's title text was not fetched correctly
            logger.warning(f'Missing critical selectors:{e}')
            continue
        
        title_desc_element = ''
        try:
            title_desc_element = element.find_element_by_class_name('extended-text__short').text.replace('\n','').replace('Читать ещё','')
        except Exception as ex:
            logger.warning(f'Missing description:{ex}')

        results.append((title_text_element, title_desc_element, title_link_element)) # save results

        print(i, title_link_element)
        print('*' * 100) 
        print(i, title_text_element)
        print('*' * 100) 
        print(i, title_desc_element)
        print('*' * 100) 

    return results


def save_screenshots(driver, results, path):

    for i, result in enumerate(results):

        link = result[2] # link to site from search result is saved in tuple under index 2

        driver.get(link)
        driver.implicitly_wait(10) # waiting either until page is fully loaded or 10 seconds, whichever comes earlier
        
        filename = Path(path) / f'{i+1}.png' # concatenates directory path with the screenshot filename
        driver.save_screenshot(str(filename)) # driver saves screenshots w/ appropriate extension and filename


def copy_screenshot(path):

    image = Image.open(path) # using library PIL for opening image
    output = BytesIO()  # create buffer of bytes in RAM
    image.convert('RGB').save(output, 'BMP') # convert image to RGB format (3 bytes per pixel) and save bytes in buffer in BMP format
    data = output.getvalue()[14:] # first 14 bytes are for BPM file header, the rest of bytes belongs to picture
    output.close()
    cb.OpenClipboard() 
    cb.EmptyClipboard()
    cb.SetClipboardData(cb.CF_DIB, data) # insert bytes in buffer and tell Windows that those bytes are DIB (BITMAPINFO - bitmap bits)
    cb.CloseClipboard()
    image.close() # close file


def compose_wordpad_document(results, path):

    f = Path('results.rtf') # create Path object which purpose will be defined later
    f.touch() # create file 'results.rtf'

    p = None
    try:
        p = Popen(['wordpad.exe', str(f)]) # attempt to open WordPad according to hypothesis that wordpad.exe is in system32 folder
    except FileNotFoundError as e:
        logger.warning(f'WordPad was not found in default path: {e}')
        try:                                # in case of aforementioned hypothesis is wrong, considered that Windows NT installed on machine
            p = Popen(['C:\\Program Files\\Windows NT\\Accessories\\wordpad.exe', str(f)]) # define full path to wordpad.exe and open results.rtf in wordpad.exe
        except:
            logger.error('Could not run WordPad!')
            exit(1)

    time.sleep(5) 

    SelectorWordPad = [{"title":"results.rtf - WordPad","class_name":"WordPadClass","backend":"win32"}] # Create WordPad Selector
        
    UIDesktop.UIOSelector_FocusHighlight(SelectorWordPad) # Focus on WordPad UIO Item

    for i, result in enumerate(results):
        print('result:', i, result)

        keyboard.press_and_release('ctrl+b') # activate bold
        keyboard.write(f'{str(i+1)}. {str(result[0])}') # insert result title from 'results' tuple
        keyboard.press_and_release('ctrl+b') # deactivate bold
        keyboard.press_and_release('enter')

        keyboard.write(str(result[1])) # insert description from 'results' tuple
        keyboard.press_and_release('enter')

        filename = Path(path) / f'{i+1}.png' # concatenates directory path with the screenshot filename
        try:
            copy_screenshot(filename) # call copy_screenshot method for opening and copying pic to buffer
            keyboard.press_and_release('ctrl+v')
        except Exception as e:
            logger.warning(f'Error while pasting picture: {e}')

        keyboard.press_and_release('enter')

        time.sleep(5) # allow keyboard commands to be processed in queue 

    keyboard.press_and_release('ctrl+s') # save file results.rtf

    time.sleep(5) # give time to save results

    p.terminate() # close wordpad



if __name__ == '__main__':  
    driver = init_webdriver(
        chrome_exe_path = chrome_exe_path, 
        chrome_driver_path = chrome_driver_path, 
    )

    try:
        request = f'https://yandex.ru/search/?text={SEARCH_PHRASE.replace(" ","%20")}' # concatenate url for search results using settings.py configuration
        results = extract_search_results(driver, request) # writes search results in the empty list 'results'
        save_screenshots(driver, results, screenshot_dir_path) # saves screenshots to folder
        compose_wordpad_document(results, screenshot_dir_path) # opens empty wordpad file, focuses on wordpad window, extracts titles and description and inserts it w/ pictures 
    except Exception as e:
        print(e)
    finally:
        driver.quit()

    
