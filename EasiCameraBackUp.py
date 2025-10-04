# coding: utf-8

from watchdog.observers import Observer, api
from watchdog.events import DirCreatedEvent, FileCreatedEvent, FileSystemEventHandler
from os import path, mkdir
from shutil import copy
from time import sleep

'''
watchdog库用于检测文件变化
    watchdog.observers.Observer用于创建BaseObserver对象, 核心功能
    watchdog.observers.api用于提供BaseObserver类作为observer的类型静态检查
    watchdog.events.DirCreatedEvent, watchdog.events.FileCreatedEvent用于FolderHandler.on_created中event的类型提示
    watchdog.events.FileSystemEventHandler是FolderHandler的父类, 核心功能
os库用于和系统相关功能
    os.path.exists用于检测是否存在文件夹, 防止mkdir报错
    os.path.expandvars用于获取AppData的绝对路径
    os.mkdir用于创建文件夹
shutil.copy用于复制文件
time.sleep用于休眠
'''

__doc__ = '''用于自动复制Seewo展台缓存图片的程序
理论是Seewo的EasiCamera会把拍摄的图片放到%AppData%\\Seewo\\EasiCamera\\Temp\\下, 但是每次重新打开程序都会清空上次的缓存
本程序使用watchdog持续检测该文件夹下的文件变化, 并复制到指定的文件夹
使用时请先创建EasiCameraBackUp实例, 并传入保存地址, 检测间隔
>>> backUp = EasiCameraBackUp('bakUp', 1.0)
'''


class FolderHandler(FileSystemEventHandler):
    '''继承自FileSystemEventHandler
    用于处理有新文件时的处理操作'''

    def callInit(self, checkingTime: float, backUpPath: str) -> None:
        '''实在不知道怎么把参传入父类已有的函数, 所以用了这种本方法
        FolderHandler对象创建后需要使用本函数传参'''
        self.checkingTime: float = checkingTime
        self.backUpPath: str = backUpPath

    def on_created(self, event: DirCreatedEvent | FileCreatedEvent) -> None:
        '''父类函数, 在有文件被创建时调用'''
        self.path: str = str(event.src_path)  # event.src_path返回新增文件的path
        self.folderSplit: list[str] = self.path.split('\\')
        # 防止没有文件夹和有文件夹
        if self.folderSplit[-2] == 'Temp' and not path.exists(f'{self.backUpPath}\\{self.folderSplit[-1]}'):
            mkdir(f'{self.backUpPath}\\{self.folderSplit[-1]}')
        else:
            copy(self.path, self.backUpPath +
                 f'\\{self.folderSplit[-2]}\\{self.folderSplit[-1]}')


class EasiCameraBackUp:
    '''核心程序, 使用时请先创建EasiCameraBackUp实例, 并传入保存地址, 检测间隔
    >>> backUp = EasiCameraBackUp('bakUp', 1.0)'''

    def __init__(self, backUpPath: str, checkingTime: float = 1.0) -> None:
        self.backUpPath: str = backUpPath
        if not path.exists(self.backUpPath):  # 防止已有文件夹
            mkdir(self.backUpPath)
        self.tempPath: str = path.expandvars(
            "%AppData%") + '\\Seewo\\EasiCamera\\Temp\\'  # 获取当前用户的AppData地址并拼接
        self.checkingTime: float = checkingTime

        # 实例化Handler
        self.handler: FolderHandler = FolderHandler()
        self.handler.callInit(self.checkingTime, self.backUpPath)
        self.observe: api.BaseObserver = Observer()
        self.observe.schedule(self.handler, self.tempPath, recursive=True)
        self.observe.start()

        # 开始备份
        self.backUp()

    def backUp(self):
        try:
            while True:
                sleep(self.checkingTime)
        finally:
            self.observe.stop()
            self.observe.join()


# test
if __name__ == '__main__':
    easiCameraBackUp = EasiCameraBackUp('backUp', 1.0)
