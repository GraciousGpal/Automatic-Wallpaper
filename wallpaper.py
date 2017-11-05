import ctypes as c
import os
from os import listdir
from os.path import isfile, join
import time
import random
import json
import codecs

# Folder Check
directory = "images"
if not os.path.exists(directory):
    os.makedirs(directory)

# Settings Check
fn = "settings.json"
data = {"time": 60*60}
try:
    file = open(fn, 'r')
except IOError:
    with open(fn, 'w') as outfile:
        json.dump(data, outfile)

# Main Program
dirp = join(os.path.dirname(os.path.abspath(__file__)), "images")
onlyfiles = [f for f in listdir(directory) if isfile(join(directory, f))]
shuffledlist = onlyfiles[:]
shuffled = sorted(shuffledlist, key=lambda k: random.random())


def set_as_wallpaper(path):
    SPI = 20
    SPIF = 2
    return c.windll.user32.SystemParametersInfoA(SPI, 0, path.encode("us-ascii"), SPIF)


while True:
    for file in shuffled:
        set_as_wallpaper(join(dirp, file))
        output_json = json.load(codecs.open('settings.json', 'r', 'utf-8-sig'))
        time.sleep(output_json["time"])
