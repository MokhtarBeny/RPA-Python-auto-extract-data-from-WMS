'''-----------------------------------------------------------------------------------------------------------
                            Extraction pour envoyer le mail pour SCI Waves and not waves V7
                            V0.0 : Développement de l'interface de base  
 -----------------------------------------------------------------------------------------------------------'''

# -*- coding: utf-8 -*-


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
from bs4 import BeautifulSoup
from datetime import datetime, timedelta


'''---------------------------------------------------------------------------------
                    Fonction main/Initalisation
 ---------------------------------------------------------------------------------'''
 #On vérifie si on a bien tous nos arguments passés en paramètre 
def main():
    
    #print(sys.argv[1])
    parameter_sci = sys.argv[2]
    #print(parameter_sci)
    #parameter_sci = "Waves_and_not_waves"
        
    #------------------------------------------------
    #   Initialisation des variables
    #------------------------------------------------

    #On récupère les valeurs dans la config
    config = configparser.ConfigParser()
    config.read("\\\\frroyetec01p\\production$\\Extractions_Robots\\Envoi mail SCI\\Identifiant.ini")
    
    #user_name = os.getlogin() #obtention de la session utilisateur au format prenom.nom

    #Définition des variables globales
    username = config.get('Manhattan', 'username')
    password = config.get('Manhattan', 'password')
    link = config.get('Manhattan', 'link')
    server = config.get('Manhattan', 'server')
    database = config.get('Manhattan', 'database')
    data_table = config.get('Manhattan', 'data_table')
    balise_custom_content= config.get('Manhattan', 'balise_custom_content')
    balise_new_orga_v2= config.get('Manhattan', 'balise_new_orga_v2')
    balise_type = config.get(parameter_sci, 'balise_type')
    balise_fr50 = config.get('Manhattan', 'balise_fr50') 
    balise_fr50l = config.get('Manhattan', 'balise_fr50l')
    balise_job = config.get('Manhattan', 'balise_job')
    balise_name_job= config.get(parameter_sci, 'balise_name_job')


    #Paramètre de la base de données
    # Chaîne de connexion
    connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database}'
    data_name = config.get(parameter_sci, 'data_name')


        
    try:
        '''---------------------------------------------------------------------------------
                                    Initialisation
        ---------------------------------------------------------------------------------'''

        
        # On vérifie si la version du pilote WebDriver Edge est la même que celui du navigateur Edge
        EdgeChromiumDriverManager().install()

        # On ajoute les différentes options 
        options = webdriver.EdgeOptions()
        options.add_argument("--disable-windows-desktop-identity")

        # Initialisation du pilote WebDriver Edge avec les options
        driver = webdriver.Edge(options=options)

        driver.maximize_window()

        driver.get(link)

        wait = WebDriverWait(driver, 30) #Sleep chargement max des pages : 30 secondes



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

    


'''---------------------------------------------------------------------------------
    
 ---------------------------------------------------------------------------------'''
if __name__ == '__main__':
    main()