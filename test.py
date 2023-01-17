"""
    Author:Jack Xu
    Gmail:jack2919048985@gmail.com
"""
from pprint import pprint

from data_handle import disposal


def main():
    village = disposal.Village("G:\\python\\DataArrangement2.0\\data", "宝坻区", "大口屯镇", "西堼村", "./log")
    village.word01_handle()
    village.word02_handle()
    village.excel01_handle()
    village.excel02_handle()
    village.photos_handle()
    village.log_write()
    village.clean_cache()
    pprint(village.states)


if __name__ == '__main__':
    main()
