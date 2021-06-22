# pharmacy_material_outflow

import re
import openpyxl

includes_Razem = re.compile(r'Razem')

opk_list = []
opk_sum_list = []
cost_list = []

rows_names = ("OPK", "Rodzaj kosztu", "Wartość")


print('-------'*7)
print('------- SORTOWANIE ROZCHODU MATERIAŁÓW - DZIAŁU FARMACJI -------')

# dokończyć wprowadzanie, utworzyć wyjątki i wyłapywanie błędów jeśli jest różny od cyfry
print('\nproszę podać miesiąc cyfrą a rok w całoście np. "07", "2019"')
month = input('proszę podać miesiąc do wykonania analizy: ')
year = input('proszę podać rok do wykonania analizy: ')
analisPeriod = str(month + '-' + year)

print('-------'*7)
print('------- Otwieranie wymaganych plików: -------')

while True:
    print('Czego dotyczy księgowanie ??')
    print("1. Szpital")
    print("2. Dary")    
    destiny  = str(input('proszę wpisać 1 lub 2: '))
    if destiny == "1":
        destiny_name = "Szpital"
        break
    if destiny == "2":
        destiny_name = "Dary"
        break
    else:
        print()
        print("zły wybór !!")
        print("prosze wpisać 1 lub 2")
        print()


print("-------")
print('1.Rozchud Leków - ' + destiny_name)
nameFile_1 = 'Raport rozchodów - ' + analisPeriod + ' ' + destiny_name + '.xlsx'
wbPharmacyMaterialOut = openpyxl.load_workbook(r'C:\Users\mrus\Analizy\Apteka\\' + nameFile_1, data_only=True)
sheetPharmacyMaterialOut = wbPharmacyMaterialOut['Sheet']


print('etap I - tworzenie listy OPK')
for i in range(13, sheetPharmacyMaterialOut.max_row + 1):
    if sheetPharmacyMaterialOut.cell(row=i, column=2).value != None and sheetPharmacyMaterialOut.cell(row=i, column=2).value != 'Oddział docelowy [ośr. kosztów]':


        # tworzenie listy OPK 
        # lista OPK służy jako wykaz kluczy do słowników utworzonych jako rodzaj rozchodowanych materiałów 
        # po stworzeniu listy tak samo przechodzi się plik excel w celu odnalezienia wszystkich OPK i tworzenia wg. zurzytych
        # przez nie materiałów slownikow dwzorowujacych zuzycie wg OPK
        opk = includes_Razem.search(sheetPharmacyMaterialOut.cell(row=i, column=2).value)

        if opk != None:
            # tworzenie listy_sum_opk
            opk_sum_list.append(sheetPharmacyMaterialOut.cell(row=i, column=2).value)        
        else:
            # tworzenie listy_opk
            opk_list.append(sheetPharmacyMaterialOut.cell(row=i, column=2).value)       


print('-------')

print('etap II - tworzenie slowników wg zurzycia grup materiałowych przez OPK')
for i in range(13, sheetPharmacyMaterialOut.max_row + 1):
    if sheetPharmacyMaterialOut.cell(row=i, column=3).value != None:
        if sheetPharmacyMaterialOut.cell(row=i, column=3).value != 'Rodzaj kosztu':
            # print(sheetPharmacyMaterialOut.cell(row=i, column=3).value)

            if sheetPharmacyMaterialOut.cell(row=i, column=3).value not in cost_list:
                cost_list.append(sheetPharmacyMaterialOut.cell(row=i, column=3).value)  




# zmiana ciagu tekstowego na liczbe(formatu float)
#
# numberStr_delRegex_char = re.compile(r'[ zł]')
# comma_dot_conversionRegex = re.compile(r'[,]')
#
# wartosc_text = ['7 777,77 zł', '7,70 zł', '70,77 zł', ' 77 777,77 zł']
#
#
#
# for i in wartosc_text:
#     number_str = numberStr_delRegex_char.sub('', i)
#     number = comma_dot_conversionRegex.sub('.', number_str)
#     number_float = float(number)
#     print(number_float)


# 2 space is \xa0 
numberStr_delRegex_char = re.compile(r'[ zł ]')
comma_dot_conversionRegex = re.compile(r'[,]')
is_number = re.compile(r'\d*')

for i in range(13, sheetPharmacyMaterialOut.max_row + 1):
    text_value = sheetPharmacyMaterialOut.cell(row=i, column=22).value
    if text_value != None:
        number_str = numberStr_delRegex_char.sub('', text_value)
        number = comma_dot_conversionRegex.sub('.', number_str)
        
        sheetPharmacyMaterialOut.cell(row=i, column=22).value = number_str
        sheetPharmacyMaterialOut.cell(row=i, column=22).number_format = '#,##0.00'

#         # if is_number.search(number):
#         #     number_float = float(number)
#         #     sheetPharmacyMaterialOut.cell(row=i, column=24).value = number_float



# 
# delete rows and cols
# delete empty rows 
for i in range(13, sheetPharmacyMaterialOut.max_row + 1):
    if sheetPharmacyMaterialOut.cell(row=i, column=2).value not in opk_list:

        if sheetPharmacyMaterialOut.cell(row=i, column=3).value not in cost_list:
            sheetPharmacyMaterialOut.delete_rows(i)

sheetPharmacyMaterialOut.delete_rows(1, 12)

# delete empty cols
sheetPharmacyMaterialOut.delete_cols(1)
sheetPharmacyMaterialOut.delete_cols(3,18)
sheetPharmacyMaterialOut.delete_cols(4,5)


for i in range(len(rows_names)):
    sheetPharmacyMaterialOut.cell(row=1, column=i+1).value = rows_names[i]




print('')
print('-------'*7)
print('Zapis do pliku')  
           
nameFile_Save = 'test_' + nameFile_1

wbPharmacyMaterialOut.save(r'C:\Users\mrus\Analizy\Apteka\\' + nameFile_Save)

