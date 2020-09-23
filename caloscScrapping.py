import time
import datetime, openpyxl, sys
from openpyxl.styles import borders, Font
from openpyxl.styles.fills import PatternFill
from bs4 import BeautifulSoup
from urllib.request import Request
import urllib

slownik = {}
aktualnaData = str(datetime.datetime.today().strftime('%Y-%m-%d'))
def pobieranieMorele():
    linki = ['https://www.morele.net/procesor-amd-ryzen-5-3600-3-6ghz-32-mb-box-100-100000031box-5938599/',
             '',
             'https://www.morele.net/karta-graficzna-sapphire-radeon-rx-5700-xt-pulse-8gb-gddr6-11293-01-20g-6276799/',
             '',
             'https://www.morele.net/plyta-glowna-asus-tuf-b450-pro-gaming-5566520/',
             'https://www.morele.net/plyta-glowna-msi-mpg-x570-gaming-edge-wifi-5938606/',
             '',
             'https://www.morele.net/dysk-ssd-adata-xpg-gammix-s11-pro-512-gb-m-2-2280-pci-e-x4-gen3-nvme-agammixs11p-512gt-c-5625912/',
             '',
             '',
             'https://www.morele.net/pamiec-goodram-irdm-ddr4-16-gb-3600mhz-cl17-irp-3600d4v64l17s-16gdc-6432278/',
             'https://www.morele.net/pamiec-crucial-ballistix-rgb-black-at-ddr4-3600-16gb-cl16-bl2k8g36c16u4bl-6492649/',
             '',
             'https://www.morele.net/zasilacz-silentiumpc-supremo-m2-gold-550w-spc140-774137/',
             'https://www.morele.net/zasilacz-silentiumpc-supremo-fm2-gold-650w-spc168-1243754/',
             'https://www.morele.net/chlodzenie-cpu-silentiumpc-fortis-3-rgb-spc245-5940755/',
             'https://www.morele.net/chlodzenie-cpu-silentiumpc-fortis-3-evo-argb-he1425-spc278-6575309/',
             '',
             'https://www.morele.net/obudowa-silentiumpc-armis-ar7x-evo-tg-argb-spc251-6535607/',
             'https://www.morele.net/obudowa-silentiumpc-armis-ar6q-evo-tg-argb-ze-szklanym-oknem-spc256-5941407/',
             'https://www.morele.net/obudowa-silentiumpc-signum-sg1q-evo-tg-argb-pure-black-spc253-6524283/',
             'https://www.morele.net/obudowa-silentiumpc-regnum-rg6v-tg-spc261-5941313/',
             'https://www.morele.net/obudowa-silentiumpc-regnum-rg6v-evo-tg-argb-spc262-5941314/',
             'https://www.morele.net/obudowa-msi-mag-vampiric-010-4145491/',
			 '',
			 'https://www.morele.net/monitor-aoc-agon-ag241qx-1060690/']
    i = 0
    for x in linki:
        if x != '':
            print(x)
            req = Request(x, headers = {"User-Agent": "Mozilla/5.0"})
            response = urllib.request.urlopen(req)
            html = response.read()
            if response.getcode() != 200:
                print('continue')
                continue
            soup = BeautifulSoup(html, 'html.parser')
            slownik[i, 0] = soup.find('h1', {'class': ['prod-name']}).text  # nazwa
            slownik[i, 1] = soup.find('div', {'class': ['price-new']}).text  # aktualnaCena
            slownik[i, 1] = slownik[i, 1].replace("z","")
            slownik[i, 1] = slownik[i, 1].replace("ł","")
            slownik[i, 1] = slownik[i, 1].replace(" ","")
            slownik[i, 2] = soup.find('div', {'class': ['price-old']})  # poprzedniaCena
            if slownik[i, 2]:
                slownik[i, 2] = True
            else:
                slownik[i, 2] = False
            slownik[i, 3] = soup.find('button',
                                      {'class': ['btn btn-grey btn-block btn-sidebar btn-disabled']})  # dostepnosc
            slownik[i, 3] = str(slownik[i, 3])
            if 'disabled' in slownik[i, 3] or 'NIEDOSTĘPNY' in slownik[i, 3]:
                slownik[i, 3] = False
            else:
                slownik[i, 3] = True
            slownik[i, 4] = x  #link
        else:
            slownik[i, 0] = ''
            slownik[i, 1] = ''
            slownik[i, 2] = ''
            slownik[i, 3] = ''
            slownik[i, 4] = ''
        i = i + 1
    return i - 1


def zapisDatyMorele(p):
    plik = openpyxl.load_workbook(p)
    odl = 0
    for x in range(5, 1000):
        if plik["Arkusz1"].cell(row=1, column=x).value == aktualnaData:
            return odl
        elif plik["Arkusz1"].cell(row=1, column=x).value is None:
            plik["Arkusz1"].cell(row=1, column=x).value = aktualnaData
            plik["Arkusz1"].cell(row=1, column=x).fill = PatternFill(fgColor="C6E0B4", fill_type="solid")
            odl = x
            plik.save(p)
            plik.close()
            return odl
    plik.save(p)
    plik.close()
    return odl


def zapisDanychMorele(odl, p):
    plik = openpyxl.load_workbook(p)
    i = 0
    ile = 0
    for x in slownik:
        #print(f'{x}  {ile}  {i}')
        #print(ile)
        if ile == 0:
            plik["Arkusz1"].cell(row=i + 2, column=2).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=2).value = slownik[i, 0]
            ile += 1
        elif ile == 1:
            if str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) > slownik[i, 1] and str(plik["Arkusz1"].cell(row=i + 2, column=odl - 1).value) != None:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="FFE699", fill_type="solid")
            elif str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) < slownik[i, 1]:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="F4B084", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=odl).value = slownik[i, 1]
            ile += 1
        elif ile == 2:
            if slownik[i, 2]:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=4).value = slownik[i, 2]
            ile += 1
        elif ile == 3:
            if slownik[i, 3]:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=3).value = slownik[i, 3]
            ile += 1

        else:
            plik["Arkusz1"].cell(row=i + 2, column=1).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=1).value = slownik[i, 4]
            ile = 0
            i += 1

    plik.save(p)
    plik.close()

def pobieranieXkom():
    linki = ['https://www.x-kom.pl/p/500085-procesor-amd-ryzen-5-amd-ryzen-5-3600.html',
             '',
             'https://www.x-kom.pl/p/521421-karta-graficzna-amd-sapphire-radeon-rx-5700-xt-pulse-8gb-gddr6.html',
             '',
             'https://www.x-kom.pl/p/493301-plyta-glowna-socket-am4-asus-tuf-b450-pro-gaming.html',
             'https://www.x-kom.pl/p/500398-plyta-glowna-socket-am4-msi-mpg-x570-gaming-edge-wifi.html',
             '',
             'https://www.x-kom.pl/p/474474-dysk-ssd-adata-512gb-m2-pcie-nvme-xpg-gammix-s11-pro.html',
             '',
             '',
             'https://www.x-kom.pl/p/531223-pamiec-ram-ddr4-goodram-16gb-2x8gb-3600mhz-cl17-irdm-pro.html',
             'https://www.x-kom.pl/p/550277-pamiec-ram-ddr4-crucial-16gb-2x8gb-3600mhz-cl16-ballistix-black-rgb.html',
             '',
             'https://www.x-kom.pl/p/308096-zasilacz-do-komputera-silentiumpc-supremo-m2-550w-80-plus-gold.html',
             'https://www.x-kom.pl/p/363851-zasilacz-do-komputera-silentiumpc-supremo-fm2-650w-80-plus-gold.html',
             'https://www.x-kom.pl/p/529353-chlodzenie-procesora-silentiumpc-fortis-3-rgb-140mm.html',
             'https://www.x-kom.pl/p/550449-chlodzenie-procesora-silentiumpc-fortis-3-evo-argb-140mm.html',
             '',
             'https://www.x-kom.pl/p/546569-obudowa-do-komputera-silentiumpc-armis-ar7x-evo-tg-argb.html',
             'https://www.x-kom.pl/p/544983-obudowa-do-komputera-silentiumpc-armis-ar6q-evo-tg-argb.html',
             'https://www.x-kom.pl/p/548256-obudowa-do-komputera-silentiumpc-signum-sg1q-evo-tg-argb.html',
             'https://www.x-kom.pl/p/541990-obudowa-do-komputera-silentiumpc-regnum-rg6v-tg-pure-black.html',
             'https://www.x-kom.pl/p/542017-obudowa-do-komputera-silentiumpc-regnum-rg6v-evo-tg-argb.html',
             'https://www.x-kom.pl/p/491808-obudowa-do-komputera-msi-mag-vampiric-010.html',
			 '',
			 'https://www.x-kom.pl/p/333314-monitor-led-24-aoc-agon-ag241qx.html',
             'https://www.x-kom.pl/p/359197-sluchawki-bezprzewodowe-steelseries-arctis-7-czarne-bezprzewodowe.html']
    i = 0
    for x in linki:
        if x != '':
            print(x)
            req = Request(x)
            response = urllib.request.urlopen(req)
            html = response.read()
            if response.getcode() != 200:
                print('continue')
                continue
            soup = BeautifulSoup(html, 'html.parser')
            try:
                slownik[i, 0] = soup.find('h1', {'class': ['sc-1x6crnh-5 gOwOoL']}).text  # nazwa
            except:
                slownik[i, 0] = soup.find('h1', {'class': ['sc-1x6crnh-5 cYILyh']}).text  # nazwa
            slownik[i, 1] = soup.find('div', {'class': ['u7xnnm-4 iVazGO']}).text  # aktualnaCena
            slownik[i, 1] = slownik[i, 1][:8]  # formatowanie ceny
            slownik[i, 1] = slownik[i, 1].replace(" z","")
            slownik[i, 1] = slownik[i, 1].replace(" ","")
            slownik[i, 2] = soup.find('div', {'class': ['u7xnnm-3 gAOShm']})  # poprzedniaCena
            if slownik[i, 2]:
                slownik[i, 2] = True
            else:
                slownik[i, 2] = False
            slownik[i, 3] = soup.find('button',
                                      {'class': ['sc-15ih3hi-0 sc-1smss4h-3 fWpoQc sc-1hdxfw1-0 gqKXGR']})  # dostepnosc
            slownik[i, 3] = str(slownik[i, 3])
            if 'disabled' in slownik[i, 3]:
                slownik[i, 3] = False
            else:
                slownik[i, 3] = True
            slownik[i, 4] = x  #link
        else:
            slownik[i, 0] = ''
            slownik[i, 1] = ''
            slownik[i, 2] = ''
            slownik[i, 3] = ''
            slownik[i, 4] = ''
        i = i + 1
    return i - 1


def zapisDatyXkom(p):
    plik = openpyxl.load_workbook(p)
    odl = 0
    for x in range(5, 1000):
        if plik["Arkusz1"].cell(row=1, column=x).value == aktualnaData:
            return odl
        elif plik["Arkusz1"].cell(row=1, column=x).value is None:
            plik["Arkusz1"].cell(row=1, column=x).value = aktualnaData
            plik["Arkusz1"].cell(row=1, column=x).fill = PatternFill(fgColor="C6E0B4", fill_type="solid")
            odl = x
            break
    plik.save(p)
    plik.close()
    return odl


def zapisDanychXkom(odl, p):
    plik = openpyxl.load_workbook(p)
    i = 0
    ile = 0
    for x in slownik:
        #print(f'{x}  {ile}  {i}')
        #print(ile)
        if ile == 0:
            plik["Arkusz1"].cell(row=i + 2, column=2).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=2).value = slownik[i, 0]
            ile += 1
        elif ile == 1:
            if str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) > slownik[i, 1] and str(plik["Arkusz1"].cell(row=i + 2, column=odl - 1).value) != None:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="FFE699", fill_type="solid")
            elif str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) < slownik[i, 1]:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="F4B084", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=odl).value = slownik[i, 1]
            ile += 1
        elif ile == 2:
            if slownik[i, 2] == True:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=4).value = slownik[i, 2]
            ile += 1
        elif ile == 3:
            if slownik[i, 3] == True:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=3).value = slownik[i, 3]
            ile += 1

        else:
            plik["Arkusz1"].cell(row=i + 2, column=1).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=1).value = slownik[i, 4]
            ile = 0
            i += 1

    plik.save(p)
    plik.close()

def pobieranieRTV():
    linki = ['https://www.euro.com.pl/procesory/amd-procesor-amd-ryzen-5-3600.bhtml',
             '',
             'https://www.euro.com.pl/karty-graficzne/sapphire-karta-graf-sapphire-pulse-rx5700xt-8gb.bhtml',
             '',
             'https://www.euro.com.pl/plyty-glowne/asus-plyta-glowna-asus-tuf-b450-pro-gaming.bhtml',
             'https://www.euro.com.pl/plyty-glowne/msi-plyta-glowna-msi-mpg-x570-gam-edge-wifi_1.bhtml',
             '',
             'https://www.euro.com.pl/dyski-wewnetrzne-ssd/adata-dysk-adata-ssd-xpg-s11-512gb-pcie-m-2.bhtml',
             '',
             '',
             'https://www.euro.com.pl/pamieci-ram/goodram-pamiecpc-good-ddr4irdmpro16-36002x8cza.bhtml',
             '',
             '',
             'https://www.euro.com.pl/zasilacze-do-komputerow-pc/silentiumpc-supremo-m2-gold-550w.bhtml',
             'https://www.euro.com.pl/zasilacze-do-komputerow-pc/silentiumpc-supremo-fm2-gold-750w-80-gold.bhtml',
             '',
             'https://www.euro.com.pl/chlodzenie-procesory/silentiumpc-chlodzenie-cpu-sil-pc-fortis-3-evo-argb.bhtml',
             '',
             '',
             '',
             'https://www.euro.com.pl/obudowy-pc/silentiumpc-obudowa-pc-sile-pcsignum-sg1qevo-tgargb.bhtml',
             'https://www.euro.com.pl/obudowy-pc/silentiumpc-obudowapc-sile-pc-regnumrg6vtgpure-black.bhtml',
             'https://www.euro.com.pl/obudowy-pc/silentiumpc-obudowa-pc-sile-reg-rg6vevotgargbpurebla.bhtml',
             '',
			 '',
			 'https://www.euro.com.pl/monitory-led-i-lcd/aoc-agon-ag241qx.bhtml']
    i = 0
    for x in linki:
        if x != '':
            print(x)
            req = Request(x, headers = {"User-Agent": "Mozilla/5.0"})
            response = urllib.request.urlopen(req)
            html = response.read()
            if response.getcode() != 200:
                print('continue')
                continue
            soup = BeautifulSoup(html, 'html.parser')
            try:
                slownik[i, 0] = soup.find('h1', {'class': ['selenium-KP-product-name']}).text  # nazwa
            except:
                slownik[i, 0] = soup.find('title').text  # nazwa2
            try:
                slownik[i, 1] = soup.find('div', {'class': ['price-normal selenium-price-normal']}).text  # aktualnaCena
            except:
                slownik[i, 1] = ''
            slownik[i, 1] = slownik[i, 1].replace("z","")
            slownik[i, 1] = slownik[i, 1].replace("ł","")
            slownik[i, 1] = slownik[i, 1].replace(" ","")
            slownik[i, 2] = soup.find('div', {'class': ['price-old']})  # poprzedniaCena
            if slownik[i, 2]:
                slownik[i, 2] = True
            else:
                slownik[i, 2] = False
            slownik[i, 3] = soup.find('div',
                                      {'class': ['label-button label-UNAVAILABLE_AT_THE_MOMENT']})  # dostepnosc
            slownik[i, 3] = str(slownik[i, 3])
            if 'nie jest dostępny' in slownik[i, 3] or 'Niedostępny' in slownik[i, 3]:
                slownik[i, 3] = False
            else:
                slownik[i, 3] = True
            if slownik[i, 3]:
                slownik[i, 3] = soup.find('div',
                                          {'class': ['product-unavailable']})  # dostepnosc2
                slownik[i, 3] = str(slownik[i, 3])
                if 'nie jest dostępny' in slownik[i, 3] or 'niedostępny' in slownik[i, 3]:
                    slownik[i, 3] = False
                else:
                    slownik[i, 3] = True
            slownik[i, 4] = x  #link
        else:
            slownik[i, 0] = ''
            slownik[i, 1] = ''
            slownik[i, 2] = ''
            slownik[i, 3] = ''
            slownik[i, 4] = ''
        i = i + 1
    return i - 1


def zapisDatyRTV(p):
    plik = openpyxl.load_workbook(p)
    odl = 0
    for x in range(5, 1000):
        if plik["Arkusz1"].cell(row=1, column=x).value == aktualnaData:
            return odl
        elif plik["Arkusz1"].cell(row=1, column=x).value is None:
            plik["Arkusz1"].cell(row=1, column=x).value = aktualnaData
            plik["Arkusz1"].cell(row=1, column=x).fill = PatternFill(fgColor="C6E0B4", fill_type="solid")
            odl = x
            break
    plik.save(p)
    plik.close()
    return odl


def zapisDanychRTV(odl, p):
    plik = openpyxl.load_workbook(p)
    i = 0
    ile = 0
    for x in slownik:
        #print(f'{x}  {ile}  {i}')
        #print(ile)
        if ile == 0:
            plik["Arkusz1"].cell(row=i + 2, column=2).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=2).value = slownik[i, 0]
            ile += 1
        elif ile == 1:
            if str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) > slownik[i, 1] and str(plik["Arkusz1"].cell(row=i + 2, column=odl - 1).value) != None:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="FFE699", fill_type="solid")
            elif str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) < slownik[i, 1]:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="F4B084", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=odl).value = slownik[i, 1]
            ile += 1
        elif ile == 2:
            if slownik[i, 2]:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=4).value = slownik[i, 2]
            ile += 1
        elif ile == 3:
            if slownik[i, 3]:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=3).value = slownik[i, 3]
            ile += 1

        else:
            plik["Arkusz1"].cell(row=i + 2, column=1).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=1).value = slownik[i, 4]
            ile = 0
            i += 1

    plik.save(p)
    plik.close()

def pobieranieMedia():
    linki = ['https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/procesory/procesor-amd-ryzen-5-3600',
             '',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/karty-graficzne/sapphire-pulse-radeon-rx-5700-xt-8g-gddr6-hdmi-triple-dp-oc-w-bp-uefi-11293-01-20g',
             '',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/plyty-glowne/asus-tuf-b450-pro-gaming-am4-b450-ddr4-3533mhz-dual-m-2-dvi-d-hdmi-tuf-b450-pro-gaming',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/plyty-glowne/plyta-glowna-msi-mpg-x570-gaming-edge-wi-fi',
             '',
             '',
             '',
             '',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/pamieci-ram/goodram-irdm-pro-pamiec-ddr4-16gb-3600mhz-cl17-1-35v-czarna-irp-3600d4v64l17-16g',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/pamieci-ram/pamiec-ram-crucial-ballistix-16gb-3600mhz-ddr4-cl16-dimm-2x8-white-rgb',
             '',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/zasilacze/supremo-m2-550w-80-gold-psu-modular',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/zasilacze/supremo-fm2-gold-650w-modular',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/chlodzenie/chlodzenie-cpu-fortis-3-he1425',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/chlodzenie/wentylator-silentiumpc-fortis-3-evo-argb-he1425',
             '',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/obudowy/obudowa-komputerowa-silentium-pc-armis-ar7x-evo-tg-argb',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/obudowy/obudowa-komputerowa-silentium-pc-armis-ar6q-evo-tg-argb',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/obudowy/obudowa-pc-signum-sg1q-evo-tg-argb',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/obudowy/obudowa-komputerowa-silentium-pc-regnum-rg6v-tg',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/obudowy/obudowa-komputerowa-silentium-pc-regnum-rg6v-evo-argb',
             'https://www.mediaexpert.pl/komputery-i-tablety/podzespoly-komputerowe/obudowy/obudowa-komputerowa-msi-mag-vampiric-100',
			 '',
			 'https://www.mediaexpert.pl/komputery-i-tablety/monitory-led/monitor-aoc-ag241qx']
    i = 0
    for x in linki:
        if x != '':
            print(x)
            req = Request(x, headers = {"User-Agent": "Mozilla/5.0"})
            response = urllib.request.urlopen(req)
            html = response.read()
            if response.getcode() != 200:
                print('continue')
                continue
            soup = BeautifulSoup(html, 'html.parser')
            slownik[i, 0] = soup.find('h1', {'class': ['a-typo is-primary']}).text  # nazwa
            try:
                slownik[i, 1] = soup.find('span', {'class': ['a-price_price']}).text  # aktualnaCena
            except:
                slownik[i, 1] = ''
            slownik[i, 1] = slownik[i, 1].replace("z","")
            slownik[i, 1] = slownik[i, 1].replace("ł","")
            slownik[i, 1] = slownik[i, 1].replace(" ","")
            slownik[i, 2] = soup.find('div',
                                      {'class': ['c-offerBox_discount ']})  # poprzedniacena
            slownik[i, 2] = str(slownik[i, 2])
            if 'Taniej o' in slownik[i, 2]:
                slownik[i, 2] = True
            else:
                slownik[i, 2] = False
            if slownik[i, 2] is False:
                slownik[i, 2] = soup.find('p',
                                          {'class': ['is-firstRow']})  # poprzedniacena2
                slownik[i, 2] = str(slownik[i, 2])
                if 'Cena z kodem' in slownik[i, 2] or 'NOCNA' in slownik[i, 2]:
                    slownik[i, 2] = True
                else:
                    slownik[i, 2] = False
            slownik[i, 3] = soup.find('div',
                                      {'class': ['a-typo is-text']})  # dostepnosc
            slownik[i, 3] = str(slownik[i, 3])
            if 'w wybranych sklepach' in slownik[i, 3] or 'Produkt' in slownik[i, 3]:
                slownik[i, 3] = False
            else:
                slownik[i, 3] = True
            slownik[i, 4] = x  #link
        else:
            slownik[i, 0] = ''
            slownik[i, 1] = ''
            slownik[i, 2] = ''
            slownik[i, 3] = ''
            slownik[i, 4] = ''
        i = i + 1
    return i - 1


def zapisDatyMedia(p):
    plik = openpyxl.load_workbook(p)
    odl = 0
    for x in range(5, 1000):
        if plik["Arkusz1"].cell(row=1, column=x).value == aktualnaData:
            return odl
        elif plik["Arkusz1"].cell(row=1, column=x).value is None:
            plik["Arkusz1"].cell(row=1, column=x).value = aktualnaData
            plik["Arkusz1"].cell(row=1, column=x).fill = PatternFill(fgColor="C6E0B4", fill_type="solid")
            odl = x
            break
    plik.save(p)
    plik.close()
    return odl


def zapisDanychMedia(odl, p):
    plik = openpyxl.load_workbook(p)
    i = 0
    ile = 0
    for x in slownik:
        #print(f'{x}  {ile}  {i}')
        #print(ile)
        if ile == 0:
            plik["Arkusz1"].cell(row=i + 2, column=2).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=2).value = slownik[i, 0]
            ile += 1
        elif ile == 1:
            if str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) > slownik[i, 1] and str(plik["Arkusz1"].cell(row=i + 2, column=odl - 1).value) != None:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="FFE699", fill_type="solid")
            elif str(plik["Arkusz1"].cell(row=i + 2, column=odl-1).value) < slownik[i, 1]:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="F4B084", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=odl).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=odl).value = slownik[i, 1]
            ile += 1
        elif ile == 2:
            if slownik[i, 2]:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=4).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=4).value = slownik[i, 2]
            ile += 1
        elif ile == 3:
            if slownik[i, 3]:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="00B050", fill_type="solid")
            else:
                plik["Arkusz1"].cell(row=i + 2, column=3).fill = PatternFill(fgColor="FF0000", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=3).value = slownik[i, 3]
            ile += 1

        else:
            plik["Arkusz1"].cell(row=i + 2, column=1).fill = PatternFill(fgColor="E2EFDA", fill_type="solid")
            plik["Arkusz1"].cell(row=i + 2, column=1).value = slownik[i, 4]
            ile = 0
            i += 1

    plik.save(p)
    plik.close()

def main():
    path1 = 'C:\\Users\\piotr\\OneDrive - T-Mobile Polska S.A\\morele.xlsx'
    path2 = 'C:\\Users\\pkaniewski3\\OneDrive - T-Mobile Polska S.A\\morele.xlsx'
    pobieranieMorele()
    try:
        odl = zapisDatyMorele(path1)
        if odl != 0:
            zapisDanychMorele(odl, path1)
            print('zapis Morele1')
    except:
        odl = zapisDatyMorele(path2)
        if odl != 0:
            zapisDanychMorele(odl, path2)
            print('zapis Morele2')
    slownik.clear()
    path1 = 'C:\\Users\\piotr\\OneDrive - T-Mobile Polska S.A\\xkom.xlsx'
    path2 = 'C:\\Users\\pkaniewski3\\OneDrive - T-Mobile Polska S.A\\xkom.xlsx'
    pobieranieXkom()
    try:
        odl = zapisDatyXkom(path1)
        if odl != 0:
            zapisDanychXkom(odl, path1)
            print('zapis xKom1')
    except:
        odl = zapisDatyXkom(path2)
        if odl != 0:
            zapisDanychXkom(odl, path2)
            print('zapis xKom2')
    slownik.clear()
    path1 = 'C:\\Users\\piotr\\OneDrive - T-Mobile Polska S.A\\rtv.xlsx'
    path2 = 'C:\\Users\\pkaniewski3\\OneDrive - T-Mobile Polska S.A\\rtv.xlsx'
    pobieranieRTV()
    try:
        odl = zapisDatyRTV(path1)
        if odl != 0:
            zapisDanychRTV(odl, path1)
            print('zapis RTV1')
    except:
        odl = zapisDatyRTV(path2)
        if odl != 0:
            zapisDanychRTV(odl, path2)
            print('zapis RTV2')
    slownik.clear()
    path1 = 'C:\\Users\\piotr\\OneDrive - T-Mobile Polska S.A\\mediaexpert.xlsx'
    path2 = 'C:\\Users\\pkaniewski3\\OneDrive - T-Mobile Polska S.A\\mediaexpert.xlsx'
    pobieranieMedia()
    try:
        odl = zapisDatyMedia(path1)
        if odl != 0:
            zapisDanychMedia(odl, path1)
            print('zapis Media1')
    except:
        odl = zapisDatyMedia(path2)
        if odl != 0:
            zapisDanychMedia(odl, path2)
            print('zapis Media2')


if __name__ == "__main__":
    main()