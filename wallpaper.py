import ctypes as c
import os
from os import listdir
from os.path import isfile, join
import time
import random
import json
import codecs
import aiohttp
import asyncio
import async_timeout
from bs4 import BeautifulSoup
import re
import shutil
import time
from tqdm import tqdm
import os
import win32com.client

# Folder Check
directory = "images"

if not os.path.exists(directory):
    os.makedirs(directory)
if not os.path.exists("used"):
    os.makedirs("used")
# Settings Check
fn = "settings.json"
data = {"time": 60,"startup": False, "page": 1, "files": 0}

try:
    file = open(fn, 'r')
except IOError:
    with open(fn, 'w') as outfile:
        json.dump(data, outfile, indent=4)

# Main Program
dirp = join(os.path.dirname(os.path.abspath(__file__)), "images")
output_json = json.load(codecs.open('settings.json', 'r', 'utf-8-sig'))
dictonary = {}

# Start up check
mypath =os.getenv('APPDATA')+ "\Microsoft\Windows\Start Menu\Programs\Startup"
if output_json["startup"] is True:
    if os.path.exists(mypath+'/Automatic-Wallpaper.lnk'):
        pass
    else:
        ws = win32com.client.Dispatch("wscript.shell")
        scut = ws.CreateShortcut(mypath+'\Automatic-Wallpaper.lnk')
        scut.TargetPath = __file__.replace(".py", ".exe")
        scut.Save()
else:
    if output_json["startup"] is False:
        if os.path.exists(mypath+'/Automatic-Wallpaper.lnk'):
            os.remove(mypath+'/Automatic-Wallpaper.lnk')
    else:
        pass
def jsonwrite(data, file):
    with open(file, 'w') as outfile:
        json.dump(data, outfile, indent=4)

def set_as_wallpaper(path):
    SPI = 20
    SPIF = 2
    return c.windll.user32.SystemParametersInfoA(SPI, 0, path.encode("us-ascii"), SPIF)


async def fetch(session, url):
    with async_timeout.timeout(10):
        async with session.get(url) as response:
            return await response.text()


async def listupdate(loop):
    print("Updating List of Wallpapers...")
    lists = []
    for x in range(output_json["page"], output_json["page"]+1):
        lists.append("https://www.pexels.com/popular-photos.js?page={}".format(x))
    async with aiohttp.ClientSession(loop=loop) as session:
        for url in tqdm(lists):
            #print(url)
            html = await fetch(session, url)
            soup = BeautifulSoup(html, "lxml")
            typelist = []
            linklist = []

            for item in soup.find_all('a', {'href': True}):
                types = item["href"].replace("\\", "").replace('"', "")
                if "/photo" not in types:
                    pass
                else:
                    types = re.search('\w/(.+?)/', types).group(1)
                    typelist.append(types)
            for a in soup.find_all('img', class_='\\"photo-item__img\\"'):
                link = a["data-pin-media"].split("?", 1)[0].replace("\\", "").replace('"', "")
                linklist.append(link)
            for f, b in zip(typelist, linklist):
                dictonary[f] = b



async def download_coroutine(session, url):
    with async_timeout.timeout(10):
        async with session.get(url) as response:
            #print(url)
            filename = os.path.basename(url)
            if not os.path.exists("images/"+filename):
                with open("images/"+filename, 'wb') as f_handle:
                    while True:
                        chunk = await response.content.read(1024)
                        if not chunk:
                            break
                        f_handle.write(chunk)
                return await response.release()
            else:
                print(" [{}] already exists Skipping!.".format(filename))
                return await response.release()

async def main(loop):
    while True:
        if output_json["files"] == 0:
            await listupdate(loop)
            print("List Update Completed!")
            print("Starting to Download Wallpapers!")
            async with aiohttp.ClientSession(loop=loop) as session:
                for url in tqdm(dictonary):
                    print(url)
                    if "adult" in url:
                        pass
                    elif "woman" in url:
                        pass
                    else:
                        await download_coroutine(session, dictonary[url])
            print("Finished Downloading Wallpapers!")
            output_json["files"] = 12
            jsonwrite(output_json, "settings.json")
        else:
            onlyfiles = [f for f in listdir(directory) if isfile(join(directory, f))]
            if output_json["page"] != 5:
                if len(onlyfiles) == 0:
                    output_json["page"] +=1
            jsonwrite(output_json, "settings.json")
            output_json["files"] = len([f for f in listdir(directory) if isfile(join(directory, f))])
            jsonwrite(output_json, "settings.json")
            shuffledlist = onlyfiles[:]
            shuffled = sorted(shuffledlist, key=lambda k: random.random())
            for xx in shuffled:
                wallpaper = join(dirp, xx)
                set_as_wallpaper(wallpaper)
                time.sleep(3)
                try:
                    os.rename(wallpaper, "used/"+str(xx))
                except FileExistsError:
                    print(" {} already exists in used folder..deleting..".format(wallpaper))
                    os.remove(wallpaper)
                print("Wallpaper Set!")
                time.sleep(output_json["time"]*60)

loop = asyncio.get_event_loop()
loop.run_until_complete(main(loop))
