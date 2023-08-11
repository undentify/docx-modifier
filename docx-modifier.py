import os
import re
import logging
import docx
import datetime
import win32_setfiletime as filetime
import json

timeobj=datetime.datetime(2001,1,1,12,15,0)
properties = {
    'author' : '', #str
    'category' : '', #str
    'comments' : '', #str
    'content_status' : '', #str
    'created' : timeobj, #datetime YYYY-MM-DD HH:MM:SS
    'identifier' : '', #str
    'keywords' : '', #str
    'language' : '', #str
    'last_modified_by' : '', #str
#    'last_printed' : '', #datetime YYYY-MM-DD HH:MM:SS
    'modified' : timeobj, #datetime YYYY-MM-DD HH:MM:SS
    'revision' : 1, #int
    'subject' : '', #str
    'title' : '', #str
    'version' : '' #str
}
#дерево каталога
def directoryStructure(path, tab):
    fileList=[]
    with os.scandir(path) as entries:
        tab = tab + '   '
        for entry in entries:
            if entry.is_dir():
#                print(tab+'\\'+entry.name)
                list=directoryStructure(path+'\\'+entry.name, tab)
                fileList.extend(list)
            if entry.is_file():
#                print(tab+entry.name)
                if entry.name[-4:] in ['docx']:
                    fileList.append(path+'\\'+entry.name)
    return fileList

def strToDT(str):
    dt=datetime.datetime(int(str[0:4]), int(str[5:7]), int(str[8:10]), int(str[11:13]), int(str[14:16]), int(str[17:19]))
    return dt

def docxDateUpdater(file, props):
    mydoc = docx.Document(file)
    '''
    print(mydoc.core_properties.author) #str авторы
    print(mydoc.core_properties.category)
    print(mydoc.core_properties.comments)
    print(mydoc.core_properties.content_status)
    print(mydoc.core_properties.created) #datetime дата создания
    print(mydoc.core_properties.identifier)
    print(mydoc.core_properties.keywords)
    print(mydoc.core_properties.language)
    print(mydoc.core_properties.last_modified_by) #str кем сохранен
    print(mydoc.core_properties.last_printed) #datetime посл вывод на печать
    print(mydoc.core_properties.modified) #datetime дата последнего сохранения
    print(mydoc.core_properties.revision) #int редакция
    print(mydoc.core_properties.subject)
    print(mydoc.core_properties.title)
    print(mydoc.core_properties.version) #str номер версии
    '''
    log = ''
    if 'author' in props:
        mydoc.core_properties.author=props['author']
        log = log + 'Автор: ' + mydoc.core_properties.author + '\n'
    if 'category' in props:
        mydoc.core_properties.category=props['category']
        log = log + 'Категория: ' + mydoc.core_properties.category + '\n'
    if 'comments' in props:
        mydoc.core_properties.comments=props['comments']
        log = log + 'Комментарии: ' + mydoc.core_properties.comments + '\n'
    if 'content_status' in props:
        mydoc.core_properties.content_status=props['content_status']
        log = log + 'Статус документа: ' + mydoc.core_properties.content_status + '\n'
    if 'created' in props:
        mydoc.core_properties.created=props['created']
        log = log + 'Дата создания содержимого: ' + mydoc.core_properties.created.strftime('%Y.%m.%d %H:%M:%S') + '\n'
    if 'identifier' in props:
        mydoc.core_properties.identifier=props['identifier']
        log = log + 'Идентификатор: ' + mydoc.core_properties.identifier + '\n'
    if 'keywords' in props:
        mydoc.core_properties.keywords=props['keywords']
        log = log + 'Теги: ' + mydoc.core_properties.keywords + '\n'
    if 'language' in props:
        mydoc.core_properties.language=props['language']
        log = log + 'Язык: ' + mydoc.core_properties.language + '\n'
    if 'last_modified_by' in props:
        mydoc.core_properties.last_modified_by=props['last_modified_by']
        log = log + 'Кем сохранен: ' + mydoc.core_properties.last_modified_by + '\n'
    if 'last_printed' in props:
        mydoc.core_properties.last_printed=props['last_printed']
        log = log + 'Последний вывод на печать: ' + mydoc.core_properties.last_printed.strftime('%Y.%m.%d %H:%M:%S') + '\n'
    if 'modified' in props:
        mydoc.core_properties.modified=props['modified']
        log = log + 'Дата последнего сохранения: ' + mydoc.core_properties.modified.strftime('%Y.%m.%d %H:%M:%S') + '\n'
    if 'revision' in props:
        mydoc.core_properties.revision=props['revision']
        log = log + 'Редакция: ' + str(mydoc.core_properties.revision) + '\n'
    if 'subject' in props:
        mydoc.core_properties.subject=props['subject']
        log = log + 'Тема: ' + mydoc.core_properties.subject + '\n'
    if 'title' in props:
        mydoc.core_properties.title=props['title']
        log = log + 'Название: ' + mydoc.core_properties.title + '\n'
    if 'version' in props:
        mydoc.core_properties.version=props['version']
        log = log + 'Номер версии: ' + mydoc.core_properties.version + '\n'

    #t.strftime('%m/%d/%Y')
    mydoc.save(file)
    #    os.utime(file, times=(int(round(props['file_modified'].timestamp())), int(round(props['file_modified'].timestamp()))))
    if 'file_created' in props:
        filetime.setctime(file, int(round(props['file_created'].timestamp())))
        log = log + 'Время создания файла: ' + props['file_created'].strftime('%Y.%m.%d %H:%M:%S') + '\n'
    if 'file_modified' in props:
        filetime.setmtime(file, int(round(props['file_modified'].timestamp())))
        log = log + 'Время изменения файла: ' + props['file_modified'].strftime('%Y.%m.%d %H:%M:%S') + '\n'
    if 'file_access' in props:
        filetime.setatime(file, int(round(props['file_access'].timestamp())))
        log = log + 'Время доступа к файлу: ' + props['file_access'].strftime('%Y.%m.%d %H:%M:%S') + '\n'
    logging.debug('Document properties was set to:\n'+log)
    return log

def readProps(file):
    props = {}
    f = open(file, 'r', encoding='utf-8')
    for line in f:
        variable = line[:re.search('=',line).start()]
        if variable in ['created', 'last_printed', 'modified', 'file_created', 'file_modified', 'file_access']:
            string = line[re.search('\".*\"', line).start() + 1:re.search('\".*\"', line).end() - 1]
            if string == '': pass
            else: props[variable] = strToDT(string)
        elif variable in ['revision']:
            props[variable] = int(line[re.search('\".*\"', line).start() + 1:re.search('\".*\"', line).end() - 1])
        elif variable in ['author', 'category', 'comments', 'content_status', 'identifier', 'keywords', 'language', 'last_modified_by', 'subject', 'title', 'version']:
            props[variable] = line[re.search('\".*\"',line).start()+1:re.search('\".*\"',line).end()-1]
        else:
            logging.info('Variable '+variable+' not supported. Skipped.')
    f.close()
    return props

if __name__=='__main__':
    logging.basicConfig(format="%(asctime)s.%(msecs)05d: %(message)s", level=logging.INFO, datefmt="%H:%M:%S")
    logging.info('Starting...')
    basePath = os.getcwd()
    basePath = basePath + '\\res'
    logging.info('Current Directory is: '+basePath)
    userProps = readProps('properties.ini')
    tab = ''
    fileList = directoryStructure(basePath, tab)
    logging.info(' - '+str(len(fileList))+' DOCX files found in directory:\n'+'\n'.join(fileList)+'\n')
    filecounter=0
    for file in fileList:
        resultLog=docxDateUpdater(file, userProps)
        logging.info('File "'+file.split('\\')[-1]+'" updated successfully')
        filecounter+=1
    logging.info(str(filecounter)+' files was updated.')
    logging.info('\nNew properties are: \n'+resultLog)
    logging.info('FINISHED')