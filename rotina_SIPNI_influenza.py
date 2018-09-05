# weekly code to download printing's informations from nddprint#
# selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
# unzip and read csv file
import fnmatch, os
import zipfile
# timesleep
import time
# shutil
import shutil
# store into database
import csv
import cx_Oracle
from unicodedata import normalize
"""
--- FIRST ---
1º - Mover todos os arquivos "Old" para as respectivas pastas Old. São eles:
1.1 - coberturaVacinalInfluenza_*.xls - From: "C:\\BI\\Arquivos\\INFLUENZA\\"	->	To:	"C:\\BI\\Arquivos\\INFLUENZA\\Old"

1.2 - dosesAplicadasInfluenzaGrupo_*.xls - 	From: "C:\\BI\\Arquivos\\INFLUENZA\\COM COMORBIDADE"	->	To:	"C:\\BI\\Arquivos\\INFLUENZA\\COM COMORBIDADE\\Old"

1.3 - dosesAplicadasInfluenzaGrupo_*.xls -	From: "C:\\BI\\Arquivos\\INFLUENZA\\PRIORITARIO"	->	To:	"C:\\BI\\Arquivos\\INFLUENZA\\PRIORITARIO\\Old"

1.4 - dosesAplicadasInfluenzaGrupo_*.xls -	From: "C:\\BI\\Arquivos\\INFLUENZA\\\SEM COMORBIDADE"	->	To:	"C:\\BI\\Arquivos\\INFLUENZA\\SEM COMORBIDADE\\Old"
"""
# -----------------------------------------
# Função para mover o arquivo 'coberturaVacinalInfluenza_2018' para a pasta Old e incrementar.
def moveToXLSOld(to_path, pattern):
    # copy old file to old folder and increment number
    old_file_path = to_path + "Old\\"
    for root, dirs, files in os.walk(to_path):
        for filename in fnmatch.filter(files, pattern):
            csv_file = os.path.join(root, filename)
            count_old_files = len(os.listdir(old_file_path))
            shutil.move(csv_file, old_file_path + filename[:-4] + " ({}).xls".format(count_old_files + 1))
            break
        break


# Chamando a função e movendo 'cobertura' para Old;
moveToXLSOld("C:\\BI\\Arquivos\\INFLUENZA\\", "cobertura*.xls")

# Chamando a função e movendo 'dosesAplicadas' (Doses Aplicadas Por Comorbidades) para Old;
moveToXLSOld("C:\\BI\\Arquivos\\INFLUENZA\\COM COMORBIDADE\\", "dosesAplicadas*.xls")

# Chamando a função e movendo 'dosesAplicadas' (Doses Aplicadas Grupos Prioritarios) para Old;
moveToXLSOld("C:\\BI\\Arquivos\\INFLUENZA\\PRIORITARIO\\", "dosesAplicadas*.xls")

# Chamando a função e movendo 'dosesAplicadas' (Doses Aplicadas Outros Grupos) para Old;
moveToXLSOld("C:\\BI\\Arquivos\\INFLUENZA\\SEM COMORBIDADE\\", "dosesAplicadas*.xls")
# -----------------------------------------
'''
--- SECOND ---
2º - Baixar os arquivos de 'Cobertura Vacinal' no SIPNI e movê-los para a pasta 'Arquivos\Influenza':
2.1 - O código abaixo fará a seguinte rotina:
2.2 - Download do arquivo 'coberturaVacinalInfluenza_*.xls' na pasta DOWNLOADS;
2.3 - Função move o arquivo de 'cobertura' para '\BI\Arquivos\Influenza'.
'''
# -----------------------------------------
# login no sipni
browser = webdriver.Chrome(executable_path="C:\\Python\\Selenium_WebDriver\\chromedriver.exe")
browser.get("http://sipni.datasus.gov.br/si-pni-web/faces/relatorio/consolidado/coberturaVacinalCampanhaInfluenza.jsf")
browser.maximize_window()
time.sleep(5)
access_mun_button = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:console2']/tbody/tr/td[6]/label")[0].click()
time.sleep(3)
dropdownUF = browser.find_element_by_xpath("//*[@id='relatorioEnvioForm:estadual']/div[3]/span")
dropdownUF.click()
pesquisa = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:estadual_filter']")
pesquisa[0].send_keys('RIO GRANDE DO SUL')
pesquisa[0].send_keys(Keys.ENTER)
time.sleep(5)
access_search = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:pesquisar']")[0].click()
# wait until the download is finished
time.sleep(25)
access_export_file = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:j_idt436_content']/a[1]")[0].click()

# wait until the download is finished
time.sleep(10)

# close browser
browser.close()
# -----------------------------------------

# -----------------------------------------
# Função para mover o arquivo.
def moveTo(from_path, to_path, pattern):
    # copy new file to process' folder
    for root, dirs, files in os.walk(from_path):
        for filename in fnmatch.filter(files, pattern):
            csv_file = os.path.join(root, filename)
            shutil.move(csv_file, to_path)
            break
        break


# Chamando a função e movendo 'cobertura' de 'Downloads' para 'Influenza':
moveTo("C:\\Users\\ses83971874053\\Downloads", "C:\\BI\\Arquivos\\INFLUENZA", "cobertura*.xls")
# -----------------------------------------

# -----------------------------------------
'''
--- THIRD ---
3º - Baixar os arquivos de 'dosesAplicadasCampanha' no SIPNI, MOVER UM POR VEZ para sua respectiva pasta e 
depois voltar a fazer o donwload do arquivo seguinte:
3.1 - O código abaixo fará a seguinte rotina:
3.2 - Download do arquivo 'dosesAplicadas' (Grupos Prioritários) na pasta DOWNLOADS;
3.3 - Função move o arquivo 'dosesAplicadas' para '\BI\Arquivos\Influenza\PRIORITARIO';
3.4 - Download do arquivo 'dosesAplicadas' (Grupos Com Comorbidades) na pasta DOWNLOADS;
3.5 - Função move o arquivo 'dosesAplicadas' para '\BI\Arquivos\Influenza\COM COMORBIDADE';
3.6 - Download do arquivo 'dosesAplicadas' (Outros Grupos) na pasta DOWNLOADS;
3.7 - Função move o arquivo 'dosesAplicadas' para '\BI\Arquivos\Influenza\SEM COMORBIDADE';
3.8 - Finaliza e fecha browser.
'''
# -----------------------------------------
# login no sipni
browser = webdriver.Chrome(executable_path="C:\\Python\\Selenium_WebDriver\\chromedriver.exe")
browser.get("http://sipni.datasus.gov.br/si-pni-web/faces/relatorio/consolidado/dosesAplicadasCampanhaInfluenza.jsf")
browser.maximize_window()
time.sleep(5)
access_mun_button = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:console']/tbody/tr/td[6]/label")[0].click()
time.sleep(3)
dropdownUF = browser.find_element_by_xpath("//*[@id='relatorioEnvioForm:estadual']/div[3]/span")
dropdownUF.click()
pesquisa = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:estadual_filter']")
pesquisa[0].send_keys('RIO GRANDE DO SUL')
pesquisa[0].send_keys(Keys.ENTER)
time.sleep(10)
access_search = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:pesquisar']")[0].click()

# wait until the download is finished
time.sleep(10)
access_export_file = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:j_idt427_content']/a[1]")[0].click()
#function moves the file and then starts the new download
time.sleep(20)
# -----------------------------------------
# Chamando a função e movendo 'dosesAplicadas*.xls' de 'Downloads' para 'Influenza\PRIORITARIO':
moveTo("C:\\Users\\ses83971874053\\Downloads", "C:\\BI\\Arquivos\\INFLUENZA\\PRIORITARIO", "dosesAplicadas*.xls")
# -----------------------------------------

# troca consulta tipo de relatório
time.sleep(3)
dropdowntipoRel = browser.find_element_by_xpath("//*[@id='relatorioEnvioForm:tipoRel']/div[3]/span")
dropdowntipoRel.click()
pesquisa = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:tipoRel_filter']")
pesquisa[0].send_keys('Grupos com Comorbidades')
pesquisa[0].send_keys(Keys.ENTER)

# wait until the download is finished
time.sleep(7)
access_export_file = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:j_idt427_content']/a[1]")[0].click()
#function moves the file and then starts the new download
time.sleep(20)
# -----------------------------------------
# Chamando a função e movendo 'dosesAplicadas*.xls' de 'Downloads' para 'Influenza\COM COMORBIDADE':
moveTo("C:\\Users\\ses83971874053\\Downloads", "C:\\BI\\Arquivos\\INFLUENZA\\COM COMORBIDADE", "dosesAplicadas*.xls")
# -----------------------------------------
# troca consulta tipo de relatório
time.sleep(5)
dropdowntipoRel = browser.find_element_by_xpath("//*[@id='relatorioEnvioForm:tipoRel']/div[3]/span")
dropdowntipoRel.click()
pesquisa = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:tipoRel_filter']")
pesquisa[0].clear()
pesquisa[0].send_keys('Outros Grupos')
pesquisa[0].send_keys(Keys.ENTER)

# download file
time.sleep(7)
access_export_file = browser.find_elements_by_xpath("//*[@id='relatorioEnvioForm:j_idt427_content']/a[1]")[0].click()
time.sleep(20)
# -----------------------------------------
# Chama a função e move 'dosesAplicadas*.xls' de 'Downloads' para 'Influenza\SEM COMORBIDADE':
moveTo("C:\\Users\\ses83971874053\\Downloads", "C:\\BI\\Arquivos\\INFLUENZA\\SEM COMORBIDADE", "dosesAplicadas*.xls")
# -----------------------------------------

# wait until the download is finished
time.sleep(5)

# close browser
browser.close()

