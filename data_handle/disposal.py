"""
    Author:Jack Xu
    Gmail:jack2919048985@gmail.com
"""
import os
import os.path as path
import shutil
import time
from pprint import pformat
from typing import List, Dict

from . import utils


class Village:

    def __init__(self, root_path: str, region_name: str, town_name: str, village_name: str, log_path: str):
        self.__states = {
            "word01_handled": False,
            "word02_handled": False,
            "photos_handled": False,
            "excel01_handled": False,
            "excel02_handled": False,
        }
        self.__root_path = root_path
        self.__region_name = region_name
        self.__town_name = town_name
        self.__village_name = village_name
        self.__log_path = log_path
        self.__path = path.join(root_path, region_name, town_name, village_name)
        self.__substances = self.__scan__()
        # self.check_all()

    def __scan__(self):
        substances = {
            "photos": [],
            "word01": [],
            "word02": [],
            "excel01": [],
            "excel02": [],
            "cache": [],
        }
        files, dst_path = utils.get_filepath(self.path, []), path.join(self.path, "暂存")
        file_names = [file.split('\\')[-1] for file in files]
        if not path.exists(dst_path):
            os.mkdir(dst_path)
        file_names_dic: Dict[str, int] = {}  # 避免文件名重复
        for i in range(len(files)):
            file = files[i]
            file_name = file_names[i]
            # 将所有文件移动至暂存文件夹
            if file_name not in file_names_dic.keys():
                file_names_dic[file_name] = 1
                renamed_file = path.join(dst_path, file_name)
                os.rename(file, renamed_file)
                files[i] = renamed_file
            else:
                file_names_dic[file_name] += 1
                renamed_file = path.join(dst_path, f"{file_names_dic[file_name]:02d}-{file_name}")
                os.rename(file, renamed_file)
                files[i] = renamed_file
        for old_dir in os.listdir(self.path):
            # 将除"暂存"之外的文件夹全部删除
            if old_dir != "暂存":
                shutil.rmtree(path.join(self.path, old_dir))
        for file in files:
            # 将所有文件进行分类
            if not (file.startswith('~$') or file.startswith('.')) and (file.endswith('doc') or file.endswith('docx') or file.endswith('.wps')):
                if file.endswith('doc'):
                    try:
                        file = utils.doc_to_docx(file)
                    except Exception:
                        if file not in substances["cache"]:
                            substances["cache"].append(file)
                elif file.endswith('wps'):
                    try:
                        file = utils.wps_to_docx(file)
                    except Exception:
                        if file not in substances["cache"]:
                            substances["cache"].append(file)
                docx_serial = utils.docx01_or_docx02(file)
                if docx_serial == 1:
                    if file not in substances["word01"]:
                        substances["word01"].append(file)
                elif docx_serial == 2:
                    if file not in substances["word02"]:
                        substances["word02"].append(file)
                elif docx_serial == 3:
                    if file not in substances["cache"]:
                        substances["cache"].append(file)
            elif not (file.startswith('~$') or file.startswith('.')) and (file.endswith('.xls') or file.endswith('.xlsx')):
                if file.endswith('.xls'):
                    try:
                        file = utils.xls_to_xlsx(file)
                    except Exception:
                        if file not in substances["cache"]:
                            substances["cache"].append(file)
                if utils.xlsx01_or_xlsx02(file):
                    if file not in substances["excel01"]:
                        substances["excel01"].append(file)
                else:
                    if file not in substances["excel02"]:
                        substances["excel02"].append(file)
            elif file.endswith('.jpg') or file.endswith('png'):
                if file not in substances["photos"]:
                    substances["photos"].append(file)
            elif file.endswith('.zip'):
                # TODO
                if file not in substances["cache"]:
                    substances["cache"].append(file)
            else:
                if file not in substances["cache"]:
                    substances["cache"].append(file)
        return substances

    def word01_handle(self):
        """
        处理 附件1-单体抗震性能调查表.docx
        """
        path_to_store = path.join(self.path, "附件1-单体抗震性能调查表")
        if not path.exists(path_to_store):
            os.mkdir(path_to_store)
        word01_ls = self.substances["word01"]
        if len(word01_ls) == 0:
            # 未找到 word01
            try:
                utils.VillageWord01Handle.case00()
                return
            except Exception:
                return
        elif len(word01_ls) == 1:
            # 所有 word01 统一存在一个文档中
            try:
                utils.VillageWord01Handle.case01(path_to_store, word01_ls[0])
            except Exception:
                return
        else:
            # 所有 word01 分别存在一个单独的文档中
            try:
                utils.VillageWord01Handle.case02(path_to_store, word01_files_ls=word01_ls)
            except Exception:
                return
        self.states["word01_handled"] = True
        self.substances["word01"] = [path.join(path_to_store, d) for d in os.listdir(path_to_store)]
        if len(word01_ls) == 1:
            self.substances["word01"].append(word01_ls[0])

    def word02_handle(self):
        """
        处理 附件2-整体抗震性能统计表.docx
        """
        path_to_store = self.path
        if not path.exists(path_to_store):
            os.mkdir(path_to_store)
        word02_ls = self.substances["word02"]
        if len(word02_ls) == 0:
            # 未找到 word02
            # TODO 附件二可能和所有附件一存在一个word文档中
            try:
                utils.VillageWord02Handle.case00(path_to_store, self.substances["word01"][-1])
                self.substances["word01"].pop(-1)
            except Exception:
                return
        elif len(word02_ls) == 1:
            # 找到 word02
            try:
                utils.VillageWord02Handle.case01(path_to_store, word02_file=word02_ls[0])
            except Exception:
                return
        else:
            pass
        self.states["word02_handled"] = True
        self.substances["word02"] = [path.join(path_to_store, "附件2-整体抗震性能统计表.docx")]

    def excel01_handle(self):
        """
        处理 单体抗震性能调查表.xlsx
        """
        path_to_store = self.path
        if not path.exists(path_to_store):
            os.mkdir(path_to_store)
        excel01_ls = self.substances["excel01"]
        if len(excel01_ls) == 0:
            # 未找到 excel01
            if self.states["word01_handled"]:
                try:
                    utils.VillageExcel01Handle.case00(
                        path_to_store,
                        self.region_name,
                        self.town_name,
                        self.village_name,
                        path.join(self.path, "附件1-单体抗震性能调查表")
                    )
                except Exception:
                    return
            else:
                return
        elif len(excel01_ls) == 1:
            # 找到 excel01
            try:
                utils.VillageExcel01Handle.case01(path_to_store, excel01_file=excel01_ls[0])
            except Exception:
                return
        else:
            pass
        self.states["excel01_handled"] = True
        self.substances["excel01"] = [path.join(path_to_store, "单体抗震性能调查表.xlsx")]

    def excel02_handle(self):
        """
        处理 整体抗震性能统计表.xlsx
        """
        path_to_store = self.path
        if not path.exists(path_to_store):
            os.mkdir(path_to_store)
        excel02_ls = self.substances["excel02"]
        if len(excel02_ls) == 0:
            # 未找到 excel02
            if self.states["word02_handled"]:
                try:
                    utils.VillageExcel02Handle.case00(path_to_store, self.village_name,
                                                      excel02_file=path.join(self.path, "附件2-整体抗震性能统计表.docx"))
                except Exception:
                    return
            else:
                return
        elif len(excel02_ls) == 1:
            # 找到 excel02
            try:
                utils.VillageExcel02Handle.case01(path_to_store, excel02_file=excel02_ls[0])
            except Exception:
                return
        else:
            pass
        self.states["excel02_handled"] = True
        self.substances["excel02"] = [path.join(path_to_store, "整体抗震性能统计表.xlsx")]

    def photos_handle(self):
        path_to_store = path.join(self.path, "照片")
        if not path.exists(path_to_store):
            os.mkdir(path_to_store)
        photos_ls = self.substances["photos"]
        if len(photos_ls) == 0:
            # 未找到任何照片
            try:
                utils.VillagePhotosHandle.case00()
                return
            except Exception:
                return
        else:
            # 找到照片
            try:
                utils.VillagePhotosHandle.case01(path_to_store, photo_files_ls=photos_ls)
            except Exception:
                return
        self.states["photos_handled"] = True
        self.substances["photos"] = [path.join(path_to_store, d) for d in os.listdir(path_to_store)]

    def log_write(self):
        now_time = time.strftime("%Y-%m-%d-%Hh%Mm%Ss-", time.localtime())
        if self.states == {
            "word01_handled": True,
            "word02_handled": True,
            "photos_handled": True,
            "excel01_handled": True,
            "excel02_handled": True,
        }:
            if "未知" in os.listdir(path.join(self.path, "照片")):
                log_file_name = path.join(self.log_path, f"{self.region_name}-{self.town_name}-{self.village_name}-{now_time}log.txt")
                log_file_name = log_file_name.replace('\\', '/')
                with open(log_file_name, "w") as log_file:
                    log_file.write("********************************************************************\n")
                    log_file.write(f"区名：{self.region_name}\n")
                    log_file.write(f"镇名：{self.town_name}\n")
                    log_file.write(f"村名：{self.village_name}\n")
                    log_file.write("********************************************************************\n")
                    log_file.write("--------------\n")
                    log_file.write("照片命名不规范!\n")
                    log_file.write("照片：\n")
                    log_file.write("{}\n".format(pformat(os.listdir(path.join(self.path, "照片", "未知")))))
                    log_file.write("--------------\n")
                    log_file.write("********************************************************************\n")
        else:
            log_file_name = path.join(self.log_path, f"{self.region_name}-{self.town_name}-{self.village_name}-{now_time}log.txt")
            log_file_name = log_file_name.replace('\\', '/')
            with open(log_file_name, "w") as log_file:
                log_file.write("********************************************************************\n")
                log_file.write(f"区名：{self.region_name}\n")
                log_file.write(f"镇名：{self.town_name}\n")
                log_file.write(f"村名：{self.village_name}\n")
                log_file.write("********************************************************************\n")
                log_file.write(
                    "暂存：\n{}\n".format(pformat(os.listdir(path.join(self.path, "暂存"))))
                )
                log_file.write("********************************************************************\n")
                if not self.states["word01_handled"]:
                    log_file.write("--------------\n")
                    log_file.write("附件1未完成整理\n")
                    log_file.write("--------------\n")
                elif not self.states["word02_handled"]:
                    log_file.write("--------------\n")
                    log_file.write("附件2未完成整理\n")
                    log_file.write("--------------\n")
                elif not self.states["excel01_handled"]:
                    log_file.write("--------------\n")
                    log_file.write("表格1未完成整理\n")
                    log_file.write("--------------\n")
                elif not self.states["excel02_handled"]:
                    log_file.write("--------------\n")
                    log_file.write("表格2未完成整理\n")
                    log_file.write("--------------\n")
                elif not self.states["photos_handled"]:
                    log_file.write("--------------\n")
                    log_file.write("照片未完成整理\n")
                    log_file.write("--------------\n")
                log_file.write("********************************************************************\n")

    def clean_cache(self) -> None:
        """
        如果该村整理内容全部整理完成，则清除暂存文件夹
        """
        if self.states == {
            "word01_handled": True,
            "word02_handled": True,
            "photos_handled": True,
            "excel01_handled": True,
            "excel02_handled": True,
        }:
            shutil.rmtree(path.join(self.path, "暂存"))

    @property
    def states(self):
        return self.__states

    @property
    def root_path(self):
        return self.__root_path

    @property
    def region_name(self):
        return self.__region_name

    @property
    def town_name(self):
        return self.__town_name

    @property
    def village_name(self):
        return self.__village_name

    @property
    def log_path(self):
        return self.__log_path

    @property
    def path(self):
        return self.__path

    @property
    def substances(self):
        return self.__substances


class Town(object):
    def __init__(self, root_path: str, region_name: str, town_name: str, log_path: str):
        self.__states = {
            "word02_handled": False,
            "excel01_handled": False,
            # "villages_handled": False
        }
        self.__root_path = root_path
        self.__region_name = region_name
        self.__town_name = town_name
        self.__log_path = log_path
        self.__path = path.join(root_path, region_name, town_name)
        self.__villages: List[Village] = []
        self.__substances = self.__scan__()
        # -------------------------------------------------

    def __scan__(self):
        substances = {
            "village_names": [],
            "word02": [],
            "excel01": [],
            "cache": [],
        }
        for ele_dir in os.listdir(self.path):
            ele_path = path.join(self.path, ele_dir)
            # 格式纠正
            if path.isdir(ele_path):
                substances["village_names"].append(ele_dir)
            elif path.isfile(ele_path):
                file = ele_path
                if file.endswith('.doc') or file.endswith('.docx') or file.endswith('.wps'):
                    if file.endswith('.doc'):
                        try:
                            file = utils.doc_to_docx(file)
                        except Exception:
                            print(f"can not turn {file}")
                            substances['cache'].append(file)
                    elif file.endswith('.wps'):
                        try:
                            file = utils.wps_to_docx(file)
                        except Exception:
                            print(f"can not turn {file}")
                            substances["cache"].append(file)
                    docx_serial = utils.docx01_or_docx02(file)
                    if docx_serial == 2:
                        substances["word02"].append(file)
                    else:
                        substances["cache"].append(file)
                elif file.endswith('.xls') or file.endswith('.xlsx'):
                    if file.endswith('.xls'):
                        try:
                            file = utils.xls_to_xlsx(file)
                        except Exception:
                            print(f"can not turn {file}")
                            substances["cache"].append(file)
                    if utils.xlsx01_or_xlsx02(file):
                        substances["excel01"].append(file)
                    else:
                        substances["cache"].append(file)
                else:
                    substances["cache"].append(file)
        return substances

    def villages_handle(self):
        for village_name in self.village_names:
            village = Village(self.root_path, self.region_name, self.town_name, village_name, self.log_path)
            self.villages.append(village)
            village.word01_handle()
            village.word02_handle()
            village.excel01_handle()
            village.excel02_handle()
            village.photos_handle()
            village.log_write()
            village.clean_cache()

    def word02_handle(self):
        word02_ls = self.substances["word02"]
        path_to_store = self.path
        if len(word02_ls) == 0:
            try:
                utils.TownWord02Handle.case00()
                return
            except Exception:
                return
        elif len(word02_ls) == 1:
            try:
                utils.TownWord02Handle.case01(path_to_store, self.region_name, self.town_name, word02_ls[0])
            except Exception:
                return
        else:
            return
        self.states["word02_handled"] = True
        self.substances["word02"] = [path.join(path_to_store, f"{self.region_name}-{self.town_name}-整体抗震性能统计表.docx")]

    def excel01_handle(self):
        excel01_ls = self.substances["excel01"]
        path_to_store = self.path
        if len(excel01_ls) == 0:
            try:
                utils.TownExcel01Handle.case00()
                return
            except Exception:
                return
        elif len(excel01_ls) == 1:
            try:
                utils.TownExcel01Handle.case01(path_to_store, self.region_name, self.town_name, excel01_ls[0])
            except Exception:
                return
        else:
            return
        self.states["excel01_handled"] = True
        self.substances["excel01"] = [path.join(path_to_store, f"{self.region_name}-{self.town_name}-单体抗震性能调查表.xlsx")]

    @property
    def root_path(self):
        return self.__root_path

    @property
    def region_name(self):
        return self.__region_name

    @property
    def town_name(self):
        return self.__town_name

    @property
    def log_path(self):
        return self.__log_path

    @property
    def path(self):
        return self.__path

    @property
    def substances(self):
        return self.__substances

    @property
    def village_names(self):
        return self.substances["village_names"]

    @property
    def villages(self):
        return self.__villages

    @property
    def states(self):
        return self.__states

    def __repr__(self):
        return f"{self.region_name}-{self.town_name}"
