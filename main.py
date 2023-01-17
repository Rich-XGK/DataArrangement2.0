"""
    Author:Jack Xu
    Gmail:jack2919048985@gmail.com
"""
import os
from pprint import pprint

from data_handle.disposal import Town
from data_handle.utils import clean_region_dir


def main():
    root_path = "G:\\python\\DataArrangement2.0\\data"
    log_path = "G:\\python\\DataArrangement2.0\\log"
    region_dirs = [d for d in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, d))]
    for region_dir in region_dirs:
        clean_region_dir(os.path.join(root_path, region_dir))
    region_names = [d for d in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, d))]
    for region_name in region_names:
        town_names = [d for d in os.listdir(os.path.join(root_path, region_name)) if os.path.isdir(os.path.join(root_path, region_name, d))]
        for town_name in town_names:
            town = Town(root_path, region_name, town_name, log_path)
            town.villages_handle()
            town.word02_handle()
            town.excel01_handle()


if __name__ == "__main__":
    main()
