# unprotect excel file v2
# pyinstaller --clean --onefile unprotectexcel.py --name unprotectexcel.exe --icon unprotectexcel_1.ico

import os
import re
import zipfile
import tempfile


# Funkcia vyhlada vsetky excel subory a vrati pole z nazvamy
def find_excels():
    ls = os.listdir()                       # list dir
    excels = []
    
    for file in ls:                         # prejde cely dir
        if file.split('.')[-1] == 'xlsx':   # posledny prvok oddeleny ciarkou musi byt excel format
            excels.append(file)             # vlozi do pola nazov najdeneneho excel subora
            
    return excels


# Funkcia vyhlada zamok pre workbook a vymaze ho.
def unprotect_workbook(file):

    patterns = [
        r'<workbookProtection .*workbookAlgorithmName="SHA-512"/>',
        r'<workbookProtection .*lockStructure="1"/>'
    ]
    
    for pattern in patterns:                    # prejde vsetky paterny
        tags = re.findall(pattern, file)        # vyhlada patern v subore pomocou modulu 're'
        if len(tags) == 1:
            #print(' [unlock workbook] {}'.format(tags[0]))
            break                               # zastavy for ak najde patern
    if len(tags) == 0:                      
        return file                             # vrati neupraveny subor ak nenajde patern
    
    file = file.replace(tags[0], '')            # vymera najdeny patern
    
    return file                                 # vrati upraveny subor


# Funkcia vyhlada zamok pre sheet a vymaze ho.
def unprotect_sheet(file):

    patterns = [
        r'<sheetProtection .*selectLockedCells="1"/>',
        r'<sheetProtection .*scenarios="1"/>',
        r'<sheetProtection .*autoFilter="0"/>',
        r'<sheetProtection .*formatColumns="0"/>',
        r'<sheetProtection .*formatRows="0"/>'
    ]
    
    for pattern in patterns:
        tags = re.findall(pattern, file)
        if len(tags) == 1:
            #print(' [unlock sheet] {}'.format(tags[0]))
            break
    if len(tags) == 0:
        return file
    
    file = file.replace(tags[0], '')
    return file


# Vstup je excel subor
def core(excel):

    path = 'xl/worksheets/sheet'                                    # cesta pre excel sheety
    path_w = 'xl/workbook.xml'                                      # cesta pre hlavny workbook subor

    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(excel))   # vytvory tmp subor
    os.close(tmpfd)

    with zipfile.ZipFile(excel, 'r') as sheet:                      # rozbali excel do zip archivu pre citanie
        with zipfile.ZipFile(tmpname, 'w') as sheetunprotect:       # rozbali tmp excel do zip archivu pre zapis

            for item in sheet.infolist():                           # prejde vsetky subory v origin zip excel
                
                if path not in item.filename:                       # ak subor nie je sheet a ani workbook
                    if path_w != item.filename:
                        sheetunprotect.writestr(item, sheet.read(item.filename)) # subor iba zapise do zip archivu
                    
                if path in item.filename:                           # ak najde subor sheet
                    #print(item.filename)
                    data = sheet.read(item.filename)                # precitane data ulozi to premenej
                    data = data.decode('UTF-8')                     # prekodovanie dat do utf-8 pre specialne znaky
                    unprotect = unprotect_sheet(data)               # funkcia pre vymazanie zamku, vrati upraveny subor
                    sheetunprotect.writestr(item, unprotect)        # zapise upraveny subor do zip archivu
                elif path_w == item.filename:                       # ak najde subor workbook tak to iste
                    #print(item.filename)
                    data = sheet.read(item.filename)
                    data = data.decode('UTF-8')
                    unprotect = unprotect_workbook(data)
                    sheetunprotect.writestr(item, unprotect)

    new_file = excel.split('.')                                     # nazov subora oddeli od bodiek, vytvory pole 
    new_file = new_file[:-1] + ['unprotect'] + new_file[-1:]        # posklada novy nazov, stale je to pole
    new_file = '.'.join(new_file)                                   # pospaja pole bodkamy, teraz je to string
    
    try:
        os.rename(tmpname, new_file)                                # modul os, premenuje subor tmp subor (zip archove) za novy
        print('[+] Save to {}'.format(new_file))                
    except:
        os.remove(tmpname)                                          # vymaze tmp subor
        print('[-] already unprotect!')
    print('[+] ...............{} OK'.format('.'*len(excel)))


if __name__ == '__main__':

    print('\n UNPROTECT EXCEL \n')
    print(' Unprotect locked workbooks and sheets.\n')
    
    excels = find_excels()                                          # funkcia najde vsetke excel subory 

    print('[+] Najdene subory {}'.format(len(excels)))
    if len(excels) > 0:
        vstup = input('\nOdomknut vsetky najdene subory? y/n (default = y) : ')

        if (vstup == 'y' or vstup == '' or vstup == 'Y'):
            for excel in excels:                                    # prejde vsetke excel subory
                print('[+] UNPROTECT {}.....'.format(excel))
                core(excel)                                         # spusti hlavnu funkciu
        else:
            pass

    input('\n\n Stlac ENTER pre ukoncenie.')
    
