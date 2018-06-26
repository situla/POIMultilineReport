# POI Multiline Report
    ###### en:
This project must help you to create the report from textual file with dividers. HSSF part of Apache POI (https://poi.apache.org) library is used. You may control report parameters in the textual settings file, with formatting like .ini file in Windows system.
    ###### ru:
Этот проект может помочь создать отчет из текстового файла с разделителями. В проекте используется часть библиотеки POI (HSSF). Вы можете управлять параметрами отчета с помощью текстового файла настроек, имеющего формат, такой как, например, ini-файлы в системе Windows.
    **en:**
To compile the project, go to the source directory and command:
```bash
javac -d . *.java
```
    **ru:**
Для компиляции проекта перейдите в директорию с исходниками и скомандуйте:
```bash
javac -d . *.java
```
    **en:**
see the directory ru/
    **ru:**
появится директория ru/
    **en:**
The following line is required to run the program:
```bash
java ru.learn2prog.poi.POIMultilineReport -f filename.csv -p property.file
```

where filename.csv - file with delimeters, with useful information (copy this file to directory with ru/ directory or define path to this file),

property.file - settings file.

    **ru:**
Для запуска программы требуется следующая строка:
```bash
java ru.learn2prog.poi.POIMultilineReport -f filename.csv -p property.file
```
где filename.csv - информационный файл с разделителями (скопируйте текстовый файл с разделителями в одну директорию с папкой ru/ или укажите полный путь к этому файлу),

property.file - файл настроек отчета (скопируйте текстовый файл с настройками в директорию с папкой ru/ или укажите полный путь к этому файлу)

## Settings (настройки)
### Formulas (формулы)
    **en:**
If you want to create formula in column, add to settings file this construction:
```bash
CellType4=formula
CellFormula4=$E$2*B?
```
where "4" - number of column, $E2$2 - cell number with absolute address, symbol "?" mean current number of excel line. All as in EXCEL, except for the character "?", which is replaced by the line number of this cell. This is necessary to place a repeating formula in EXCEL cells

    **ru:**
Если вы хотите создать формулу в столбце, добавьте в файл настроек такую конструкцию:
```bash
CellType4=formula
CellFormula4=$E$2*B?
```
где «4» - номер столбца, $E2$2 - номер ячейки с абсолютным адресом, символ «?» означает текущий номер строки. Все как в EXCEL, за исключением символа "?", который заменяется на номер строки данной ячейки. Это необходимо для размещения повторяющейся формулы в ячейках EXCEL.

### Line colors difference (Черезстрочное выделение)
###### en:

###### ru:
Для удобства четные строки отчета можно выделить цветом, для этого в файле настроек существуют опции:

```bash
LineColorsDifference=true
EvenLineColor=RED
```

Если установить опцию LineColorsDifference=true, то цвет заливки четных строк табличного отчета будет таким, как Вы укажете в EvenLineColor. Варианты цветов можно посмотреть в http://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/IndexedColors.html.

