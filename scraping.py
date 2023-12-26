
import selenium.webdriver as webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter

dizi_adi_english=[]
dizi_adi_turkısh=[]
yapım_yli=[]
ort_suresi=[]
dizi_turu1=[]
dizi_turu2=[]
dizi_turu3=[]
blm_sayısı=[]
szn_sayısı=[]
senaryoo=[]
oyuncu1=[]
oyuncu2=[]
oyuncu3=[]
oyuncu4=[]
oyuncu5=[]
oyuncu6=[]
oyuncu7=[]
oyuncu8=[]
oyuncu9=[]
oyuncu10=[]
yapimci_sirkett=[]
yayinci1=[]
yayinci2=[]
ihrac=[]
yönetmenn=[]
senarist1=[]
senarist2=[]
senarist3=[]
senarist4=[]
senarist5=[]
yapımcı1=[]
yapımcı2=[]
yapımcı3=[]
müzisyen1=[]
müzisyen2=[]
müzisyen3=[]

workbook = xlsxwriter.Workbook( "dizi_verisi_2000i2006.xlsx" )
worksheet = workbook.add_worksheet()

for i in range(1,8):
    for k in range(1,31):
        try:
            if i <=10:
                driver = webdriver.Chrome()
                driver.get("https://www.imdb.com/search/title/?title_type=tv_series&release_date=200"+str(i-1)+"-01-01,200"+str(i-1)+"-12-31&sort=moviemeter,asc&countries=TR&languages=tr")
                sleep(4)
            else:
                driver = webdriver.Chrome()
                driver.get("https://www.imdb.com/search/title/?title_type=tv_series&release_date=20"+str(i-1)+"-01-01,20"+str(i-1)+"-12-31&sort=moviemeter,asc&countries=TR&languages=tr")
                sleep(4)

            #dizi sayfasına gir
            wait = WebDriverWait(driver, 10)

            sırala = driver.find_element(By.XPATH,
                                               "/html/body/div[2]/main/div[2]/div[3]/section/section/div/section/section/div[2]/div/section/div[2]/div[2]/div[1]/div[2]/div/button[3]")
            driver.execute_script("arguments[0].click();", sırala)

            dizi_syf_blgi=wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/main/div[2]/div[3]/section/section/div/section/section/div[2]/div/section/div[2]/div[2]/ul/li["+str(k)+"]/div[2]/div/div/div[1]/a")))
            driver.execute_script("arguments[0].click();", dizi_syf_blgi)
            sleep(3)

            #dizi adı
            dizi_adı_eng = driver.find_elements(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/h1/span')
            dizi_adi_english.append(str(dizi_adı_eng[0].text))
            print(dizi_adı_eng[0].text)

            try:
                dizi_adı_tr = driver.find_elements(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/div')
                dizi_adi_turkısh.append(str(dizi_adı_tr[0].text))
                print(dizi_adı_tr[0].text)
            except:
                dizi_adi_turkısh.append(" ")
            ##bitiş yılı
            #başlama yılı
            #yapım yılı
            yapım_yılı = driver.find_elements(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/ul/li[2]/a')
            yapım_yli.append(str(yapım_yılı[0].text))
            print(yapım_yılı[0].text)

            #ortalama süresi
            try:
                ort_sure = driver.find_elements(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/ul/li[4]')
                ort_suresi.append(str(ort_sure[0].text))
                print(ort_sure[0].text)
            except:
                try:
                    ort_sure = driver.find_elements(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[2]/div[1]/ul/li[3]')
                    ort_suresi.append(str(ort_sure[0].text))
                    print(ort_sure[0].text)
                except:
                    ort_suresi.append(" ")
            #türler:
            try:
                tür= driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[1]/div[2]/a[1]')
                dizi_turu1.append(str(tür[0].text))
                print(tür[0].text)
            except:
                dizi_turu1.append(" ")
            try:
                tür= driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[1]/div[2]/a[2]')
                print(tür[0].text)
                dizi_turu2.append(str(tür[0].text))
            except:
                dizi_turu2.append(" ")
            try:
                tür= driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/div[1]/div[2]/a[3]')
                print(tür[0].text)
                dizi_turu3.append(str(tür[0].text))
            except:
                dizi_turu3.append(" ")
            #bolum sayısı

            try:
                bolum_sayisi= driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section[1]/div/section/div/div[1]/section[2]/div[1]/a/h3/span[2]')
                blm_sayısı.append(str(bolum_sayisi[0].text))
                print(bolum_sayisi[0].text)
            except:
                try:
                    bolum_sayisi = driver.find_elements(By.XPATH,
                                                        '/html/body/div[2]/main/div/section[1]/div/section/div/div[1]/section[1]/div[1]/a/h3/span[2]')
                    blm_sayısı.append(str(bolum_sayisi[0].text))
                    print(bolum_sayisi[0].text)
                except:
                    blm_sayısı.append(" ")
            #sezon sayısı
            try:
                sezon_sayisi = driver.find_elements(By.XPATH,
                                                    '/html/body/div[2]/main/div/section[1]/div/section/div/div[1]/section[2]/div[2]/div[2]/div[2]/span[1]/span/select/option[2]')
                szn_sayısı.append(str(sezon_sayisi[0].text))
                print(sezon_sayisi[0].text)
            except:
                try:
                    sezon_sayisi = driver.find_elements(By.XPATH,
                                                        '/html/body/div[2]/main/div/section[1]/div/section/div/div[1]/section[1]/div[2]/div[2]/div[2]/span[1]/span/select/option[2]')
                    szn_sayısı.append(str(sezon_sayisi[0].text))
                except:
                    szn_sayısı.append(" ")
            #senaryo
            try:
                senaryo= driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section[1]/section/div[3]/section/section/div[3]/div[2]/div[1]/section/p')
                print(senaryo[0].text)
                senaryoo.append(str(senaryo[0].text))
            except:
                senaryoo.append(" ")

            #oyuncular
            def oyuncular(index, section_num):
                oyuncular = driver.find_elements(By.XPATH,
                                                 f'//*[@id="__next"]/main/div/section[1]/div/section/div/div[1]/section[{section_num}]/div[2]/div[2]/div[{index}]/div[2]/a')
                print(oyuncular[0].text)
                globals()[f"oyuncu{index}"].append(str(oyuncular[0].text))


            for p in range(1, 11):
                for section_num in ["4", "3", "2", "5"]:
                    try:
                        oyuncular(p, section_num)
                        break
                    except:
                        continue
                else:
                    globals()[f"oyuncu{p}"].append(" ")

            all_topics_button = driver.find_element(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[1]/div/div[2]/button')
            all_topics_button.click()
            sleep(3)

            #yapımcı şirketleri
            company_credits_btn=driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div[2]/a')
            company_credits_btn.click()
            try:
                yapımcı_sirket=driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section/div/section/div/div[1]/section[1]/div[2]/ul/li/a[1]')
                print(yapımcı_sirket[0].text)
                yapimci_sirkett.append(str(yapımcı_sirket[0].text))
            except:
                yapimci_sirkett.append(" ")

            #yayıncı kuruluşlar /Company credits
            try:
                yayinci_kurulus=driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section/div/section/div/div[1]/section[2]/div[2]/ul/li[1]/a[1]')
                yayinci1.append(str(yayinci_kurulus[0].text))
                print(yayinci_kurulus[0].text)
            except:
                yayinci1.append(" ")

            try:
                yayinci_kurulus=driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section/div/section/div/div[1]/section[2]/div[2]/ul/li[2]/a[1]')
                yayinci2.append(str(yayinci_kurulus[0].text))
                print(yayinci_kurulus[0].text)
            except:
                yayinci2.append(" ")

            ## back button
            back=driver.find_element(By.XPATH,'/html/body/div[2]/main/div/section/section/div[3]/section/section/div[1]/a')
            back.click()
            sleep(3)
            #back bitti
            all_topics_button = driver.find_element(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[1]/div/div[2]/button')
            all_topics_button.click()
            sleep(3)


            ##ihraç edildi mi
            release_date_Btn=driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/a')
            release_date_Btn.click()
            try:
                driver.find_elements(By.XPATH,'/html/body/div[2]/main/div/section/div/section/div/div[1]/section[2]/div[2]/ul/li[2]/span')
                ihrac.append("evet")
            except:
                ihrac.append("hayır")

            ### back button
            back=driver.find_element(By.XPATH,'/html/body/div[2]/main/div/section/section/div[3]/section/section/div[1]/a')
            back.click()
            sleep(3)
            #back bitti
            all_topics_button = driver.find_element(By.XPATH,'//*[@id="__next"]/main/div/section[1]/section/div[3]/section/section/div[1]/div/div[2]/button')
            all_topics_button.click()
            sleep(3)

            cast_sayfasına_git = driver.find_element(By.XPATH,'/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div[1]/ul/li[2]/a')
            cast_sayfasına_git.click()

            #yönetmen
            try:
                yönetmen = driver.find_elements(By.XPATH,'//*[@id="fullcredits_content"]/table[1]/tbody/tr/td[1]')
                yönetmenn.append(str(yönetmen[0].text))
                print(yönetmen[0].text)
            except:
                yönetmenn.append(" ")

            #senaristler
            try:
                senarists = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[2]/tbody/tr[1]/td[1]/a')
                print(senarists[0].text)
                senarist1.append(str(senarists[0].text))
            except:
                senarist1.append(" ")
            try:
                senarists = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[2]/tbody/tr[2]/td[1]/a')
                print(senarists[0].text)
                senarist2.append(str(senarists[0].text))
            except:
                senarist2.append(" ")
            try:
                senarists = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[2]/tbody/tr[3]/td[1]/a')
                print(senarists[0].text)
                senarist3.append(str(senarists[0].text))

            except:
                senarist3.append(" ")
            try:
                senarists = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[2]/tbody/tr[4]/td[1]/a')
                senarist4.append(str(senarists[0].text))
                print(senarists[0].text)
            except:
                senarist4.append(" ")
            try:
                senarists = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[2]/tbody/tr[5]/td[1]/a')
                print(senarists[0].text)
                senarist5.append(str(senarists[0].text))
            except:
                senarist5.append(" ")
            #yapımcı
            try:
                yapımcı = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[4]/tbody/tr[1]/td[1]/a')
                print(yapımcı[0].text)
                yapımcı1.append(str(yapımcı[0].text))
            except:
                yapımcı1.append("")
            try:
                yapımcı = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[4]/tbody/tr[2]/td[1]/a')
                print(yapımcı[0].text)
                yapımcı2.append(str(yapımcı[0].text))
            except:
                yapımcı2.append(" ")
            try:
                yapımcı = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[4]/tbody/tr[3]/td[1]/a')
                print(yapımcı[0].text)
                yapımcı3.append(str(yapımcı[0].text))
            except:
                yapımcı3.append("")
            #müzik
            try:
                müzisyen = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[5]/tbody/tr[1]/td[1]/a')
                print(müzisyen[0].text)
                müzisyen1.append(str(müzisyen[0].text))

            except:
                müzisyen1.append("")
            try:
                müzisyen = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[5]/tbody/tr[2]/td[1]/a')
                print(müzisyen[0].text)
                müzisyen2.append(str(müzisyen[0].text))
            except:
                müzisyen2.append(" ")
            try:
                müzisyen = driver.find_elements(By.XPATH,'/html/body/div[2]/div/div[2]/div[3]/div[1]/div[1]/div[2]/table[5]/tbody/tr[3]/td[1]/a')
                print(müzisyen[0].text)
                müzisyen3.append(str(müzisyen[0].text))
            except:
                müzisyen3.append(" ")
            print("dizi bitti")
        except:
            try:
                variables = [
                    dizi_adi_english, dizi_adi_turkısh, yapım_yli, ort_suresi, dizi_turu1, dizi_turu2, dizi_turu3,
                    blm_sayısı, szn_sayısı, senaryoo, oyuncu1, oyuncu2, oyuncu3, oyuncu4, oyuncu5, oyuncu6, oyuncu7,
                    oyuncu8, oyuncu9, oyuncu10, yapimci_sirkett, yayinci1, yayinci2, ihrac, yönetmenn, senarist1,
                    senarist2, senarist3, senarist4, senarist5, yapımcı1, yapımcı2, yapımcı3, müzisyen1, müzisyen2,
                    müzisyen3
                ]

                for m in range(len(variables) - 1):
                    if len(variables[m]) > len(variables[m + 1]):
                        for n in range(m + 1, len(variables)):
                            variables[n].append(" ")

                müzisyen3.append(" ") if len(müzisyen3) > len(müzisyen2) else None

                l1, l2, l3, l4, l5, l6, l7, l8, l9, l10, l11, l12, l13, l14, l15, l16, l17, l18, l19, l20, l21, l22, l23, l24, l25, l26, l27, l28, l29, l30, l31, l32, l33, l34, l35, l36 = map(
                    len, variables)
            except:
                pass

worksheet.write('A1', 'Dizi Adı English')
worksheet.write('B1', 'Dizi Adı Turkçe')
worksheet.write('C1', 'Başlama-Bitme Yılı')
worksheet.write('D1', 'Ortalama Süre')
worksheet.write('E1', 'Tür-1')
worksheet.write('F1', 'Tür-2')
worksheet.write('G1', 'Tür-3')
worksheet.write('H1', 'Bölüm Sayısı')
worksheet.write('I1', 'Sezon Sayısı')
worksheet.write('J1', 'Senaryo')
worksheet.write('K1', 'Oyuncu-1')
worksheet.write('L1', 'Oyuncu-2')
worksheet.write('M1', 'Oyuncu-3')
worksheet.write('N1', 'Oyuncu-4')
worksheet.write('O1', 'Oyuncu-5')
worksheet.write('P1', 'Oyuncu-6')
worksheet.write('Q1', 'Oyuncu-7')
worksheet.write('R1', 'Oyuncu-8')
worksheet.write('S1', 'Oyuncu-9')
worksheet.write('T1', 'Oyuncu-10')
worksheet.write('U1', 'Yapımcı Şirket')
worksheet.write('V1', 'Yayıncı Kuruluşlar-1')
worksheet.write('W1', 'Yayıncı Kuruluşlar-2')
worksheet.write('X1', 'İhraç Edildi Mi?')
worksheet.write('Y1', 'Yönetmen')
worksheet.write('Z1', 'Senarist-1')
worksheet.write('AA1', 'Senarist-2')
worksheet.write('AB1', 'Senarist-3')
worksheet.write('AC1', 'Senarist-4')
worksheet.write('AD1', 'Senarist-5')
worksheet.write('AE1', 'Yapımcı-1')
worksheet.write('AF1', 'Yapımcı-2')
worksheet.write('AG1', 'Yapımcı-3')
worksheet.write('AH1', 'Müzisyen-1')
worksheet.write('AI1', 'Müzisyen-2')
worksheet.write('AJ1', 'Müzisyen-3')

for satir,veri in enumerate(dizi_adi_english):    worksheet.write(satir+1,0,veri)
for satir,veri in enumerate(dizi_adi_turkısh):    worksheet.write(satir+1,1,veri)
for satir,veri in enumerate(yapım_yli):           worksheet.write(satir+1,2,veri)
for satir,veri in enumerate(ort_suresi):          worksheet.write(satir+1,3,veri)
for satir,veri in enumerate(dizi_turu1):          worksheet.write(satir+1,4,veri)
for satir,veri in enumerate(dizi_turu2):          worksheet.write(satir+1,5,veri)
for satir,veri in enumerate(dizi_turu3):          worksheet.write(satir+1,6,veri)
for satir,veri in enumerate(blm_sayısı):          worksheet.write(satir+1,7,veri)
for satir,veri in enumerate(szn_sayısı):          worksheet.write(satir+1,8,veri)
for satir,veri in enumerate(senaryoo):            worksheet.write(satir+1,9,veri)
for satir,veri in enumerate(oyuncu1):             worksheet.write(satir+1,10,veri)
for satir,veri in enumerate(oyuncu2):             worksheet.write(satir+1,11,veri)
for satir,veri in enumerate(oyuncu3):             worksheet.write(satir+1,12,veri)
for satir,veri in enumerate(oyuncu4):             worksheet.write(satir+1,13,veri)
for satir,veri in enumerate(oyuncu5):             worksheet.write(satir+1,14,veri)
for satir,veri in enumerate(oyuncu6):             worksheet.write(satir+1,15,veri)
for satir,veri in enumerate(oyuncu7):             worksheet.write(satir+1,16,veri)
for satir,veri in enumerate(oyuncu8):             worksheet.write(satir+1,17,veri)
for satir,veri in enumerate(oyuncu9):             worksheet.write(satir+1,18,veri)
for satir,veri in enumerate(oyuncu10):            worksheet.write(satir+1,19,veri)
for satir,veri in enumerate(yapimci_sirkett):     worksheet.write(satir+1,20,veri)
for satir,veri in enumerate(yayinci1):            worksheet.write(satir+1,21,veri)
for satir,veri in enumerate(yayinci2):            worksheet.write(satir+1,22,veri)
for satir,veri in enumerate(ihrac):               worksheet.write(satir+1,23,veri)
for satir,veri in enumerate(yönetmenn):           worksheet.write(satir+1,24,veri)
for satir,veri in enumerate(senarist1):           worksheet.write(satir+1,25,veri)
for satir,veri in enumerate(senarist2):           worksheet.write(satir+1,26,veri)
for satir,veri in enumerate(senarist3):           worksheet.write(satir+1,27,veri)
for satir,veri in enumerate(senarist4):           worksheet.write(satir+1,28,veri)
for satir,veri in enumerate(senarist5):           worksheet.write(satir+1,29,veri)
for satir,veri in enumerate(yapımcı1):            worksheet.write(satir+1,30,veri)
for satir,veri in enumerate(yapımcı2):            worksheet.write(satir+1,31,veri)
for satir,veri in enumerate(yapımcı3):            worksheet.write(satir+1,32,veri)
for satir,veri in enumerate(müzisyen1):           worksheet.write(satir+1,33,veri)
for satir,veri in enumerate(müzisyen2):           worksheet.write(satir+1,34,veri)
for satir,veri in enumerate(müzisyen3):           worksheet.write(satir+1,35,veri)

workbook.close()
