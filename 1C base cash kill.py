# -*- coding: utf-8 -*-
import numbers
from binascii import a2b_hex
from configparser import ConfigParser
from ctypes import windll
from os import remove, path, stat, system, scandir
from subprocess import Popen, PIPE
from sys import exit, argv
from tempfile import TemporaryFile
from time import sleep
from shutil import rmtree
import win32com.shell.shell as shell
from win32con import SW_NORMAL


def _sleep(n):
    try:
        sleep(n)
    except KeyboardInterrupt:
        sleep(0)
    finally:
        sleep(0)


def _print(text, wait=0):
    print(text)
    _sleep(wait)


def get_maxlen_text(generallistbases, limittext, parametr):
    list = []
    for i in generallistbases:
        list.append(generallistbases[i][parametr])
    maxlenparametr = 0
    for i in list:
        val = i[:limittext]
        if len(val) > maxlenparametr:
            maxlenparametr = len(val)
    return maxlenparametr


def is_Admin(): return windll.shell32.IsUserAnAdmin() == 1


def showquest(text):
    try:
        n = input(text)
    except KeyboardInterrupt:
        n = None
    finally:
        if n == 'Q' or n == 'q' or n == None:
            _print('Завершение программы...', 2)
            exit(0)
        elif n == 'R' or n == 'r':
            _print('Очистка сообщений...', 1)
            system('cls')
            _print("Список баз:\n")
            showbases(generallistbases, ['base', 'connect'])
            return 'refresh'
        elif n == 'A' or n == 'a':
            if not is_Admin():
                _print('Перезапуск программы...', 2)

                # proc = Popen(
                #     'runas /trustlevel:0x40000 ' + argv[0],  # test.exe',
                #     shell=True,
                #     stdout=PIPE, stderr=PIPE
                # )
                # proc.wait()

                # shell.ShellExecuteEx(lpVerb='runas', lpFile=argv[0], lpParameters='/admin')

                # rv = None
                try:
                    # rv = \
                    shell.ShellExecuteEx(
                        lpFile=argv[0],
                        lpParameters='/admin',
                        nShow=SW_NORMAL,
                        lpVerb="runas")
                except:
                    _print("Произошла ошибка перезапуска", 2)
                # else:
                #     _print(str(rv), 2)
                finally:
                    exit(0)
        else:
            return n


def showbases(bases, points, outcommands=True):
    limittext = 40
    tmp_user = ''
    countbases = len(bases)
    _print('='*72)
    for numbase in bases:
        line = "  [" + numbase + "]"
        k = 1
        if countbases > 10 and int(numbase) < 10:
            k = 2
        elif countbases > 10 and int(numbase) > 10:
            k = 1

        line = line + ' ' * k

        for point in points:
            namebase = bases[numbase][point]
            maxlen = get_maxlen_text(bases, limittext, point)
            razd = (" " * (maxlen - len(namebase))) + "  "
            if len(namebase) > limittext:
                namebase = namebase[:(limittext - 3)] + "..."
            line = line + namebase + razd

        user = bases[numbase]['user']
        if user != tmp_user:
            tmp_user = user
            _print(user.upper())

        _print(line)

    if outcommands:
        _print('Команды для работы с программой:')
        listnum = list(generallistbases)
        diapazone = " " + listnum[0] + ".." + listnum[len(listnum) - 1]
        countP = len(diapazone) - 1
        _print(" Q" + " " * countP + "выход из программы")
        _print(" R" + " " * countP + "очистка сообщений")
        if not is_Admin():
            _print(" A" + " " * countP + 'получение списка баз всех пользователей (с правами администратора)')
        _print(diapazone + " номер базы или номера баз через пробел для удаления кэша")
        _print("=" * 72)

def delletecashe(listbasestokill, userspaths):
    for i in listbasestokill:
        isdellbase = False
        patch = userspaths[listbasestokill[i]['user']]['userpath']
        id_val = listbasestokill[i]['id']

        foldercasheroaming = path.join(patch, "AppData\\Roaming\\1C\\1cv8")
        for element in scandir(foldercasheroaming):
            if element.name == id_val and not element.is_file():
                isdellbase = deletecashepatch(element.path)

        foldercashelocal = path.join(patch, "AppData\\Local\\1C\\1cv8")
        if path.isdir(foldercashelocal):
            for element in scandir(foldercashelocal):
                if element.name == id_val and not element.is_file():
                    isdellbase = deletecashepatch(element.path)

        foldercasheroaming = path.join(patch, "AppData\\Roaming\\1C\\1cv8")
        if path.isdir(foldercasheroaming):
            for element in scandir(foldercasheroaming):
                if element.name == id_val and not element.is_file():
                    isdellbase = deletecashepatch(element.path)

        foldercashelocal1Cv82 = path.join(patch, "AppData\\Local\\1C\\1Cv82")
        if path.isdir(foldercashelocal1Cv82):
            for element in scandir(foldercashelocal1Cv82):
                if element.name == id_val and not element.is_file():
                    isdellbase = deletecashepatch(element.path)

        foldercasheroaming1Cv82 = path.join(patch, "AppData\\Roaming\\1C\\1Cv82")
        if path.isdir(foldercashelocal1Cv82):
            for element in scandir(foldercasheroaming1Cv82):
                if element.name == id_val and not element.is_file():
                    isdellbase = deletecashepatch(element.path)

    return isdellbase


def deletecashepatch(patch):
    isdellbase = False
    try:
        rmtree(patch)
        isdellbase = True
    except PermissionError:
        isdellbase = False
    finally:
        if isdellbase == False:
            _print("Не удалось удалить кэш базы ")
            _print("   Возможно база открыта.")
            _print("   ID кэша базы: " + patch, 1)
        else:
            _print("Удален каталог " + patch, 1)
        return isdellbase


global usersconfigfiles

if __name__ == '__main__':
    system('cls')
    system('color f9')

    """
    Сохраниение настроек пользователей во временный файл
    """

    _tempfilename = TemporaryFile().name
    proc = Popen(
        'regedit.exe -ea ' + _tempfilename + ' "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\"',
        shell=True,
        stdout=PIPE, stderr=PIPE
    )
    proc.wait(1, 5)
    proc.kill

    """
    Корректировка строк во временном файле настроек
    """

    readlines = []
    with open(path._getfullpathname(_tempfilename), 'r+') as fin:
        readlines = fin.readlines()
        readlines_copy = readlines.copy()
        for i in range(len(readlines_copy)):
            if len(readlines_copy[i].strip()) == 0 or readlines_copy[i].strip()[0] != '[':
                readlines_copy[i] = ''
            else:
                readlines = readlines_copy
                break

    if not len(readlines):
        _print("Нет данных для получения списка пользователей.\nВсего хорошего!\nЗавершение программы...", 2)
        exit(0)

    with open(_tempfilename, 'w') as F:
        F.writelines(readlines)

    """
    Получение данных из временного файла настроек
    """

    config = ConfigParser()
    try:
        config.read(_tempfilename, 'utf-8-sig')
    except:
        _print("Файл настроек по реестру не соответствует формату.", 2)
        exit(0)
    # finally:
    #     remove(_tempfilename)

    """
    Получение путей к файлам пользователей
    """

    usersconfigfiles = {}

    for profile in config.sections():
        keysconfig = config[profile].keys()
        if ('"fullprofile"' in keysconfig and '"profileimagepath"' in keysconfig) or '"refcount"' in keysconfig:
            hextext = config.get(profile, '"profileimagepath"')[7:-2]
            hextext = hextext.replace(',', '').replace('\\', '').replace('\n', '')
            try:
                userpath = str(a2b_hex(hextext).decode('utf-8'))
            except:
                _print("Ошибка получения параметра hextext: " + str(hextext), 5)
                continue
            userpathsplit = userpath.split('\\')
            username = userpathsplit[len(userpathsplit) - 1]
            pathconfig = path.join(userpath, 'AppData', 'Roaming', '1C', '1CEStart', 'ibases.v8i')
            try:
                st = stat(pathconfig)
            except OSError:
                continue
            else:
                usersconfigfiles.update({username: {'userpath': userpath, 'pathconfig': pathconfig}})
    del config
    if not len(usersconfigfiles):
        _print("Файлы настроек пользователей не найдены.", 2)
        exit(0)

    userspaths = usersconfigfiles.copy()

    """
    Подготовим структуру данных о базах пользователей
    """

    for user in usersconfigfiles:
        config = ConfigParser()
        fileconfig = usersconfigfiles[user]['pathconfig']
        try:
            config.read(fileconfig, 'utf-8-sig')
        except:
            _print("Ошибка чтения файла настроек " + str(fileconfig) + " пользователя " + user, 2)
            exit(0)
        else:
            usersbases = {}
            for thisbase in config.sections():
                if 'connect' not in config[thisbase]: continue
                tm_dict = {}
                for configbase in config[thisbase]:
                    tm_dict.update({configbase: config[thisbase][configbase]})
                usersbases.update({thisbase: tm_dict})
            usersconfigfiles[user] = usersbases

    """
    Получим структуру баз пользователей в виде
        generallistbases = {num:{'user':valuser,'base':valbase,'connect':valconnect,'id':valid}}
    для механизма удаления кэша
    """

    generallistbases = {}
    num = 1
    for user in usersconfigfiles:
        bases = usersconfigfiles[user]
        config = {}
        for base in bases:
            generallistbases.update(
                {str(num): {'user': user, 'base': base, 'connect': bases[base]['connect'], 'id': bases[base]['id']}})
            num += 1

    """
    Выведем список пользователей и баз. Номера баз должны соответствовать номерам в generallistbases
    """
    _print("Список баз:")
    showbases(generallistbases, ['base', 'connect'])

    """
    Выведем основной вопрос
    """

    while True:
        n = showquest("Что нужно сделать? Введите команду: ")

        if n == 'refresh':
            continue

        if str(n).find(' ') != -1:
            listN = n.split(' ')
        else:
            listN = [n]

        """
        Выведем выбранные базы и подтверждение на удаление кэшей
        """

        listbasestokill = {}

        for i in listN:
            if i in generallistbases:
                user = str(generallistbases[i]['user'])
                base = str(generallistbases[i]['base'])
                id = str(generallistbases[i]['id'])
                connect = str(generallistbases[i]['connect'])

                listbasestokill.update({i: {'user': user, 'base': base, 'id': id, 'connect': connect}})

        if len(listbasestokill) == 0:
            text = "!!! не распознана база по введенному номеру."
            if len(listN) > 1:
                text = "--- не распознаны базы по введенным номерам."
            _print(text + " Повторите еще раз...", 1)
            continue

        _print("\nВНИМАНИЕ! Будут удалены файлы кэши следующих баз:")
        showbases(listbasestokill, ['base', 'connect'], False)

        n = showquest("\nПодтвердите удаление кэшей [Y-да]: ")

        if n == 'Y' or n == 'y':
            if not delletecashe(listbasestokill, userspaths):
                _print('Кэш баз не удален. Возможно кэш уже был ранее очищен.', 2)
        else:
            _print('отмена\n', 1)
