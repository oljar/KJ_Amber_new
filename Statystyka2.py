

#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3
import datetime
import time
import os

plik=open("sprawozdanie.txt","w")
os.system("cls")

print('Program listujący urządzenia KCX odebrane przez KJ Klimor Skowarcz \n\n')
plik.write('Program listujący urządzenia KCX odebrane przez KJ Klimor Skowarcz \n\n')


w1=input('Podaj datę 1 -szą  w formacie [YYYY  mm  dd]\n')
w2=input('Podaj datę 2 -gą  w formacie [YYYY  mm  dd]\n')

d1=datetime.datetime.strptime(w1,'%Y %m %d')



d2=datetime.datetime.strptime(w2,'%Y %m %d')


# utworzenie połączenia z bazą przechowywaną na dysku
# lub w pamięci (':memory:')
con = sqlite3.connect('baza_kcx_kco.db')

# dostęp do kolumn przez indeksy i przez nazwy
con.row_factory = sqlite3.Row

# utworzenie obiektu kursora
cur = con.cursor()
serial_nr=[]

cur.execute(
        """
        SELECT ID,nr_fabr,kod_prod, nr_zlec,identity,tab_znam_nakl,kla_en_oz_kr,zasl_raczki,est_silik_rysy,kapt_srub_uziem,\
            izol_uszcz,mon_filtr,tac_ociek,mon_bypas,went,czujn_temp,rozdz_mont,nagrz_dzial,drzwi,ent_wydatek,ent_nieszczel,\
            osl_rozdz,dtr_kart_prod,ozn_KJ,uwagi_bool,uwagi_txt,podpis,now_d,now_h FROM tab

        where now_d >= DATE(?) AND now_d <=DATE(?)   """,(d1,d2))

steps = cur.fetchall()

print(f'Sprawdzono następujące urządznia pomiędzy datami {d1} oraz {d2} \n\n')
plik.write(f'Sprawdzono następujące urządznia pomiędzy datami {d1} oraz {d2} \n\n')

for step in  steps:
    serial_nr.append(step[1])
serial_nr=list(set(serial_nr))

for one_nr in serial_nr:
    print(str(one_nr))
    plik.write(' \n')
    plik.write(str(one_nr))



input()
print('-'*50)
plik.write(' \n')
plik.write('-'*50)
plik.write(' \n')


d={1:'nr fabryczny',
   2:'kod produktu',
   3:'nr zlecenia',
   4:'identyfikacja',
   5:'tabliczka znamionowa , pozostałe naklejki',
   6:'klasa energetyczna , oznaczenie króćów',
   7:'zaŚlepki, rączki',
   8:'estetyka , silikonowanie, rysy',
   9:'kapturki śrub , znaczki uziemień',
   10:'izolacja , uszczelki',
   11:'montaż filtrów',
   12:'taca ociekowa',
   13:'bypass',
   14:'wentylatory',
   15:'czujniki temperatur',
   16:'montaż rozdzielnicy',
   17:'nagrzewnica ',
   18:'drzwi',
   19:'wydatek',
   20:'nieszczelność',
   21:'osłona rozdzielnicy',
   22:'dtr-ka , karta produktu',
   23:'oznaczenie KJ',
   24:'uwagi'


   }

for nr in serial_nr:

        cur.execute(
                """
                SELECT ID,nr_fabr,kod_prod, nr_zlec,identity,tab_znam_nakl,kla_en_oz_kr,zasl_raczki,est_silik_rysy,kapt_srub_uziem,\
                    izol_uszcz,mon_filtr,tac_ociek,mon_bypas,went,czujn_temp,rozdz_mont,nagrz_dzial,drzwi,ent_wydatek,ent_nieszczel,\
                    osl_rozdz,dtr_kart_prod,ozn_KJ,uwagi_bool,uwagi_txt,podpis,now_d,now_h FROM tab

                where nr_fabr=?""",(nr,))

        steps = cur.fetchall()



        nr_fabryczny=str(d[1])+': ' +str(steps[0][1])+' '
        kod_produktu=str(d[2])+': ' +str(steps[0][2])+' '
        nr_zlecenia=str(d[3])+': '+str(steps[0][3])+' '
        header=[nr_fabryczny,kod_produktu,nr_zlecenia]

        lista=[]
        print()
        plik.write("\n")
        for step in steps:
            for i in range(0,23):
                if step[i]=="Negatyw":

                    lista.append(d[i])
                    lista=list(set(lista))


            if step[2]=='KCX-300' and float(step[20])>=0.3:
                   lista.append('przekroczony przeciek wewnętrzny')

            if step[2]=='KCX-500' and float(step[20])>=0.4:
                   lista.append('przekroczony przeciek wewnętrzny')

            if step[2]=='KCX-800' and float(step[20])>=0.5:
                   lista.append('przekroczony przeciek wewnętrzny')

            if step[2]=='KCX-1200' and float(step[20])>=0.3:
                   lista.append('przekroczony przeciek wewnętrzny')





        lista=list(set(lista))
        lista_dod=[]



        for w in range(len(header)):
            print(header[w])
            plik.write(header[w])
            plik.write("\n")




        ilosc_blendow=int(len(lista))


        if ilosc_blendow>0 :
            print('Uwagi podstawowe :')
            plik.write('Uwagi podstawowe :')
            plik.write('\n')




        lista=list(set(lista))

        for z in range(len(lista)):
            print (lista[z])
            plik.write(lista[z])
            plik.write('\n')



        ilosc_blendow_dod_bool = False
        ilosc_blendow_dod = 0


        for step in steps:

            if step[24]=="Negatyw":
                ilosc_blendow_dod_bool =True
                lista_dod.append(f'Uwagi dodatkowe : {step[25]}')

        if ilosc_blendow_dod_bool ==True:
            ilosc_blendow_dod = 1

        lista_dod=list(set(lista_dod))

        ilosc_blendow=ilosc_blendow + ilosc_blendow_dod

        for n in lista_dod:
            print(n)
            plik.write(n)


        print('\n')
        plik.write("\n")



        print(f'Ilość błędów :{ilosc_blendow}')
        plik.write(f'Ilość błędów :{ilosc_blendow}')
        plik.write("\n")
        print('-'*50)
        plik.write('-'*50)





cur.close()

plik.write("\n")
plik.write('koniec')
print('koniec')
plik.close()
input()

