import pyautogui
import time
import webbrowser
from python_imagesearch.imagesearch import imagesearch
import os
import pythoncom
import threading
import pyWinhook as pyHook
import pyautogui
import time
import argparse
        
# Create a function to terminate the script
def terminate():
    print("Script terminated")
    os._exit(1)

    # Create a new thread
def new_thread(function, *args):
    print("New thread created")
    t = threading.Thread(target=function, args=args)
    t.start()

def find_image(filename, offset=(0,0), threshold=0.8):
    pos = imagesearch(filename, precision=threshold)
    if pos[0] != -1:
        pyautogui.moveTo(pos[0] + offset[0], pos[1] + offset[1])
        return(True, (pos[0], pos[1]))
    else:
        print("image not found")
        return(False, None)
    
def open_chrome(url):
    webbrowser.open(url)

def move_to_window():
    window = pyautogui.getWindowsWithTitle("Google Chrome")[0]
    window.activate()
    x, y, width, height = window.left, window.top, window.width, window.height

    # Calculate the center of the window
    center_x = x + width // 2
    center_y = y + height // 2

    # Move the mouse to the center of the window
    pyautogui.moveTo(center_x, center_y, duration=0.25)

def scrape_coupons(website_details):
    url = website_details['url']
    clip_button = website_details['clip_button']
    next_page_button = website_details.get('next_page_button', None)
    threshold = website_details['threshold']

    open_chrome(url)
    time.sleep(5)
    move_to_window()


    image_not_found_counter = 0
    coupons_clipped = 0

    while True:
        # Find the clip button
        result = find_image(clip_button, threshold=threshold)
        
        # If we didn't find the image, check for the next page button or scroll down
        if result[0] == False:
            if image_not_found_counter > 10:
                break

            print("image not found")

            # If there's a next page button, try to click it
            if next_page_button:
                result = find_image(next_page_button, threshold=threshold)
                if result[0]:
                    pyautogui.click(button='left', clicks=1, interval=0.1, duration=0.1)
                    time.sleep(3)
                    if website_details.get('reset'):
                        pyautogui.scroll(10000)

            # Scroll down to load more coupons
            pyautogui.scroll(-500)
            if website_details.get('delay'):
                time.sleep(website_details['delay'])

            image_not_found_counter += 1
        else:
            print("image found")
            coupons_clipped += 1

            # Reset exit counter
            image_not_found_counter = 0

            # Click the clip button
            pyautogui.click(button='left', clicks=1, interval=0.1, duration=0.1)
    
    print(f"Finished, clipped {coupons_clipped} coupons")

def onKeyboardEvent(event):
    if event.Key == 'P':
        new_thread(terminate)
    return True

def parse_args():
    parser = argparse.ArgumentParser(description="Parse --open flag and --store argument.")
    parser.add_argument('--open', action='store_true', help='Set this flag to open the websites in order to log in before running the script')
    parser.add_argument('--store', type=str, help='Specify the store name')

    return parser.parse_args()

# Create a function to represent the script
def main(args):
    
    website_map = {
        'cvs': {
            'url': 'https://www.cvs.com/extracare/home',
            'clip_button': 'pics/cvs/clip.png',
            'threshold': 0.95
        },
        'vons': {
            'url': 'https://www.vons.com/foru/coupons-deals.html',
            'clip_button': 'pics/vons/clip.png',
            'next_page_button': 'pics/vons/next_page.png',
            'threshold': 0.85
        },
        'total_wine': {
            'url': 'https://www.totalwine.com/offers/coupons',
            'clip_button': 'pics/total_wine/clip.png',
            'next_page_button': 'pics/total_wine/next_page.png',
            'threshold': 0.85
        },
        'ralphs': {
            'url': 'https://www.ralphs.com/savings/cl/coupons/',
            'clip_button': 'pics/ralphs/clip.png',
            'threshold': 0.85
        },
        'food4less': {
            'url': 'https://www.food4less.com/savings/cl/coupons/',
            'clip_button': 'pics/ralphs/clip.png',
            'threshold': 0.85
        },
        'smart_and_final': {
            'url': 'https://www.smartandfinal.com/sm/planning/rsid/519/coupon-gallery?utm_source=homepage&utm_medium=smallicon2&utm_campaign=coupongallery',
            'clip_button': 'pics/smart_and_final/clip.png',
            'next_page_button': 'pics/smart_and_final/next_page.png',
            'threshold': 0.95,
            'reset': True
        },
        'walmart': {
            'url': 'https://www.walmart.com/walmartcash',
            'clip_button': 'pics/walmart/clip.png',
            'next_page_button': 'pics/walmart/next_page.png',
            'threshold': 0.80
        },
        'target': {
            'url': 'https://www.target.com/l/target-circle-deals/-/N-j1p80',
            'clip_button': 'pics/target/clip.png',
            'next_page_button': 'pics/target/next_page.png',
            'threshold': 0.90
        },
        'walgreens': {
            'url': 'https://www.walgreens.com/offers/offers.jsp/weeklyad?ban=dl_dlsp_MegaMenu_WeeklyAd',
            'clip_button': 'pics/walgreens/clip.png',
            'threshold': 0.90,
            'delay': 3
        }
    }
    if args.store:
        print(f"Store: {args.store}")
        website_map = {args.store: website_map[args.store]}

    if args.open:
        for key in website_map:
            print(f"Opening {key}")
            open_chrome(website_map[key]['url'])
        terminate() 

    for key in website_map:
        print(f"Clipping coupons for {key}")
        scrape_coupons(website_map[key])
        print('------------------------------------------')
    terminate()
        
if __name__ == "__main__":
    args = parse_args()

    hm = pyHook.HookManager()
    hm.SubscribeKeyDown(onKeyboardEvent)
    hm.HookKeyboard()
    new_thread(main, args)
    pythoncom.PumpMessages()


