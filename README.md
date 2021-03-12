# 破解保护的excel


## 1. 将保护的xls另存为xlsx

## 2. 将xlsx文件的后缀修改为zip

## 3. 解压zip

## 4. 进入xl文件夹，修改workbook.xml
- 删除"<workbookProtection[^<>]*>"节点
- 保存workbook.xml

## 5. 进入worksheets文件夹，遍历所有xml文件
- 删除<sheetProtection[^<>]*>节点
- 保存所有xml

## 6. 压缩文件为zip

## 7. 将zip后缀修改为xlsx

## 8. 如果需要，再将xlsx另存为xls

## 9. 清理临时文件

