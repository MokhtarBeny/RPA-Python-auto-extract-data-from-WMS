'''-----------------------------------------------------------------------------------------------------------
                            Extraction pour extraire les performanaces des opérateurs
                            V0.0 : Développement de l'interface de base  
 -----------------------------------------------------------------------------------------------------------'''

# -*- coding: utf-8 -*-


import os
import glob
import pyodbc
import time
import sys
import pandas as pd
import configparser
from ast import Break
from urllib.request import urlopen
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date, timedelta
from bs4 import BeautifulSoup
from datetime import datetime, timedelta


'''---------------------------------------------------------------------------------
                    Fonction main/Initalisation
 ---------------------------------------------------------------------------------'''
 #On vérifie si on a bien tous nos arguments passés en paramètre 
def main():
    
    #print(sys.argv[1])
    #parameter_sci = sys.argv[2]
    #print(parameter_sci)
    parameter_sci = "Waves_and_not_waves"
        
    #------------------------------------------------
    #   Initialisation des variables
    #------------------------------------------------

    #On récupère les valeurs dans la config
    config = configparser.ConfigParser()
    config.read("\\\\frroyetec01p\\production$\\Extractions_Robots\\Envoi mail SCI\\Identifiant.ini")

    #Définition des variables globales
    username = config.get('Manhattan', 'username')
    password = config.get('Manhattan', 'password')
    #link = config.get('Performance_employee', 'link')
    link ="https://ulorp-sci.sce.manh.com/bi/?pathRef=.public_folders%2FCustom%2BContent%2FNew%2Borganisation%2BV2%2F8_LMS%2F1_REPLICABLE%2FREP.LMS.012%253A%2BEmployee%2Bdetail%2Bperformance%2Band%2BThroughput"
    server = config.get('Manhattan', 'server')
    database = config.get('Manhattan', 'database')
    data_table = config.get('Manhattan', 'data_table')


    #Paramètre de la base de données
    # Chaîne de connexion
    connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database}'
    data_name = config.get(parameter_sci, 'data_name')

    #Définition des variables pour les différents dossiers
    user_name = os.getlogin() #obtention de la session utilisateur au format prenom.nom  
    downloads_path = os.path.join("C:\\Users\\", user_name, "Downloads") # Chemin du répertoire téléchargement utilisateur
    destination_path = "U:\\Preparation\\Centrale\\06. Gestion Activité\\MOD\\SCI à coller" # Chemin du répertoire de destination

    file_pattern = "REP.LMS.012: Employee detail performance and Throughput*.xlsx" # Modèle du fichier à rechercher

    file_name ="REP.LMS.012: Employee detail performance and Throughput.xlsx"  #Nom du fichier à créer

    # Liste des concatenations repertoire/fichier
    file_list_download = glob.glob(os.path.join(downloads_path, file_pattern)) #tous les fichiers dans le dossier téléchargement
    file_create = os.path.join(downloads_path, file_name) #le fichier dans le dossier téléchargement
    file_list_dest = os.path.join(destination_path,file_name) #le fichier dans le dossier de destination



        
    try:
        '''---------------------------------------------------------------------------------
                                    Initialisation
        ---------------------------------------------------------------------------------'''

        #On obtient la date
        if len(sys.argv) > 3 and sys.argv[3]:  #on passe une date en paramètre
            dateJ=sys.argv[3]
        else:
            dateJ=DerniereDate(date.today()) #on calcul la date de la veille

        #  Si des fichiers sont à supprimer dans le fichier téléchargementalors on les supprime
        if file_list_download:
            for file_path_download in file_list_download:
                os.remove(file_path_download)
            print("Fichiers supprimés avec succès !")
        else:
            print("Aucun fichier trouvé !")

        
        # On vérifie si la version du pilote WebDriver Edge est la même que celui du navigateur Edge
        EdgeChromiumDriverManager().install()

        # On ajoute les différentes options 
        options = webdriver.EdgeOptions()
        options.add_argument("--disable-windows-desktop-identity")

        # Initialisation du pilote WebDriver Edge avec les options
        driver = webdriver.Edge(options=options)

        driver.maximize_window()

        driver.get(link)

        wait = WebDriverWait(driver, 60) #Sleep chargement max des pages : 30 secondes



        '''---------------------------------------------------------------------------------
                                    Page Login - Partie 1 
        ---------------------------------------------------------------------------------'''

        # On attend que la page se charge complètement
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='brandingText']")))

        # On veut sélectionner l'option "ULOR-MAAuthServer" dans la liste déroulante
        select_element = driver.find_element(By.NAME, "CAMNamespace")
        select = Select(select_element)
        select.select_by_value("ulor-sci")

        '''---------------------------------------------------------------------------------
                                    Page Login - Partie 2
        ---------------------------------------------------------------------------------'''

        # On attend que la page se charge complètement
        wait.until(EC.presence_of_element_located((By.ID, "login-username")))

        email_input = driver.find_element(By.CSS_SELECTOR, "input#login-username")
        email_input.send_keys(username)  

        submit_button = driver.find_element(By.ID, "discover-user-submit")
        submit_button.click()

       

        '''---------------------------------------------------------------------------------
                                    Page Login - Partie 3
        ---------------------------------------------------------------------------------'''

        # On attend que la page se charge complètement
        wait.until(EC.presence_of_element_located((By.ID, "login-password")))
        password_input = driver.find_element(By.ID, "login-password")
        password_input.send_keys(password) 

    
        submit_button = driver.find_element(By.ID, "login-submit")
        submit_button.click()

        '''---------------------------------------------------------------------------------
                               On rempli les prompt dans la prompt page du sci
        ---------------------------------------------------------------------------------'''


        # On attend que la page se charge complètement
        wait.until(EC.presence_of_element_located((By.ID, "v34__tblDateTextBox__txtInput")))

        

        exit

        wait.until(EC.element_to_be_clickable((By.ID, 'com.ibm.bi.contentApps.teamFoldersSlideout'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//*//div[@title="{balise_custom_content}"]'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//*//div[@title="{balise_new_orga_v2}"]'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//*//div[@title="{balise_type}"]'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//*//div[@title="{balise_fr50}"]'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//*//div[@title="{balise_fr50l}"]'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f'//*//div[@title="{balise_job}"]'))).click()

        # click on ellipse

        element = wait.until(EC.presence_of_element_located((By.XPATH, f'//*//tr[@data-name="{balise_name_job}"]/td[3]')))
        action = ActionChains(driver) 
        action.move_to_element(element).click().perform() 

        # click on "Exécuter en tant que"
        try:
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//*//ul[@role="menu"]/li[1]/a[1]')))
        finally:
            action = ActionChains(driver) 
            action.move_to_element(element).click().perform() 

        # click on "Exécuter"
        try:
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//*//button[@data-tid="RUNAS_FOOTER_BUTTON_OK"]')))
        finally:
            action = ActionChains(driver) 
            action.move_to_element(element).click().perform()

        time.sleep(3)

        
        '''---------------------------------------------------------------------------------
                                Action en db
        ---------------------------------------------------------------------------------'''
        

        #on insère une ligne dans la base de données
        # On établi la connexion à la DB
        conn = pyodbc.connect(connection_string)

        #On ouvre le curseur
        cursor = conn.cursor()

        try:
            # On insère une ligne dans la db suivi exploit
            cursor.execute(f"SET dateformat dmy; INSERT INTO {data_table} ([Code_Table], [date_traitement], [nbr_lignes]) VALUES ('{data_name}', GETDATE(), 0)")

            # Commit pour appliquer les changements dans la base de données
            conn.commit()

            print("Insertion réussie!")
        except Exception as e:
            # En cas d'erreur, on annule la transaction
            conn.rollback()
            print(f"Erreur lors de l'insertion : {e}")
        finally:
            # On ferme le curseur et la connexion
            cursor.close()
            conn.close()
    
    except Exception as e:
        print(e)


    try:
        alert = Alert(driver)
        alert.accept()  # Ferme la pop-up de sécurité dans Edge
    except:
        pass

def DerniereDate(dateP):

  if dateP.weekday() == 0:  # on est lundi
    delta_jours = 3 
  else:
    delta_jours = 1 #autre jour de la semaine

  return dateP - timedelta(days=delta_jours)

    


'''---------------------------------------------------------------------------------
    
 ---------------------------------------------------------------------------------'''
if __name__ == '__main__':
    main()