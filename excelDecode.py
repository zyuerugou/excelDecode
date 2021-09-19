# -*- coding: utf-8 -*-
"""
@author: 193334
"""

'''
破解保护的excel
1. 将保护的xls另存为xlsx
2. 将xlsx文件的后缀修改为zip
3. 解压zip
4. 进入xl文件夹，修改workbook.xml
- 删除"<workbookProtection[^<>]*>"节点
- 保存workbook.xml
5. 进入worksheets文件夹，遍历所有xml文件
- 删除<sheetProtection[^<>]*>节点
- 保存所有xml
6. 压缩文件为zip
7. 将zip后缀修改为xlsx
8. 如果需要，再将xlsx另存为xls
'''


import os
import zipfile
import re
import io

'''
1. 将保护的xls另存为xlsx
手动处理
'''


'''
2. 将xlsx文件的后缀修改为zip
'''
filename = 'a'
os.rename(filename + '.xlsx', filename + '.zip')


'''
3. 解压zip
'''
zip_file = zipfile.ZipFile(filename + '.zip')

# 判断同名文件夹是否存在，若不存在则创建同名文件夹
if os.path.isdir(filename):
    rf_list = rf.namelist() # 得到压缩包里所有的文件 
    print('ZIP文件内容', rf_list) 
else:
    os.mkdir(filename)

for names in zip_file.namelist():
    zip_file.extract(names, filename + '/') # 解压文件
zip_file.close() # 关闭文件，必须有，释放内存


'''
4. 进入xl文件夹，修改workbook.xml
- 删除"<workbookProtection[^<>]*>"节点
- 保存workbook.xml
'''
def alter(file,old_str,new_str):
    with io.open(file, "r", encoding="utf-8") as f1, io.open("%s.bak" % file, "w", encoding="utf-8") as f2:
        for line in f1:
            f2.write(re.sub(old_str,new_str,line))
    os.remove(file)
    os.rename("%s.bak" % file, file)

base_dir = './{0}/xl/'.format(filename)
alter(base_dir + 'workbook.xml', '<workbookProtection[^<>]*>', '')


'''
5. 进入worksheets文件夹，遍历所有xml文件
- 删除<sheetProtection[^<>]*>节点
- 保存所有xml
'''
def getFiles(dir, suffix = None): # 遍历所有文件
    res = []
    for root, directory, files in os.walk(dir):
        for filenm in files:
            name, suf = os.path.splitext(filenm)
            if suffix is None or suf == suffix:
                res.append(os.path.join(root, filenm))
    return res

base_dir = './{0}/xl/worksheets/'.format(filename)
for file in getFiles(base_dir, '.xml'):
    alter(file, "<sheetProtection[^<>]*>", "")

'''
6. 压缩文件为zip
'''
def dfs_get_zip_file(input_path,result):
    '''
    dfs遍历所有文件
    '''
    files = os.listdir(input_path)
    for file in files:
        if os.path.isdir(input_path+'/'+file):
            dfs_get_zip_file(input_path+'/'+file,result)
        else:
            result.append(input_path+'/'+file)

# 需要压缩的目录
base_dir = './{0}'.format(filename)
# 压缩结果
target = './{0}_res.zip'.format(filename)


# 创建zip对象
f = zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED)

os.chdir(base_dir);

# 压缩base_dir下所有文件
filelists = []
dfs_get_zip_file('./',filelists)
for file in filelists:
    f.write(file)

os.chdir('../');
# 关闭zip对象
f.close()


'''
7. 将zip后缀修改为xlsx
'''
filename_res = './{0}_res'.format(filename)
os.rename(filename_res + '.zip', filename + '.xlsx')

'''
8. 如果需要，再将xlsx另存为xls

按需手动操作
'''

'''
清理临时文件
'''
for root, dirs, files in os.walk(filename, topdown=False):
    for name in files:
        os.remove(os.path.join(root, name))
    for name in dirs:
        os.rmdir(os.path.join(root, name))
os.rmdir(filename)
os.remove(filename + '.zip')


