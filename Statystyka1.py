

#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import sqlite3

# utworzenie połączenia z bazą przechowywaną na dysku
# lub w pamięci (':memory:')
con = sqlite3.connect('baza_kcx_kco.db')

# dostęp do kolumn przez indeksy i przez nazwy
con.row_factory = sqlite3.Row

# utworzenie obiektu kursora
cur = con.cursor()
aa=input('Podaj numer fabryczny \n')

#definiowanie funkcji drukowania


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
   17:'nagrzewnica działanie',
   18:'drzwi',
   19:'wydatek',
   20:'nieszczelność',
   21:'osłona rozdzielnicy',
   22:'dtr-ka , karta produktu',
   23:'oznaczenie KJ',
   24:'uwagi'


   }







# pobieranie danych z bazy

cur.execute(
        """
        SELECT ID,nr_fabr,kod_prod, nr_zlec,identity,tab_znam_nakl,kla_en_oz_kr,zasl_raczki,est_silik_rysy,kapt_srub_uziem,\
            izol_uszcz,mon_filtr,tac_ociek,mon_bypas,went,czujn_temp,rozdz_mont,nagrz_dzial,drzwi,ent_wydatek,ent_nieszczel,\
            osl_rozdz,dtr_kart_prod,ozn_KJ,uwagi_bool,uwagi_txt,podpis,now_d,now_h FROM tab

        where nr_fabr=?""",(aa,))

steps = cur.fetchall()


nr_fabryczny=str(d[1])+': ' +str(steps[1][1])+' '
kod_produktu=str(d[2])+': ' +str(steps[1][2])+' '
nr_zlecenia=str(d[3])+': '+str(steps[1][3])+' '
header=[nr_fabryczny,kod_produktu,nr_zlecenia]
lista=[]
print()
for step in steps:
    for i in range(0,25):
        if step[i]=="Negatyw":

            lista.append(d[i])
            lista=list(set(lista))

ilosc_blendow=int(len(lista))

for w in range(len(header)):
    print(header[w])
for z in range(len(lista)):
    print (lista[z])
print(f'Ilość błędów :{ilosc_blendow}')



input()

cur.close()
