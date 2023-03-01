#paquage: 
import tabula
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog
import PyPDF2
import os
from datetime import date
from datetime import datetime, timedelta
import time
import asyncio




#calcul du temps d'execution
start_time = time.time()

#chargement des données
     #les pdfs des condidats
folder_path = "QCM_pdfs"

     #les reponses correctes
correct_answers_file = "QCM_correct.xlsx"

#fonction pour calculer le temps:
#def function_to_time():

#extraire les tables des pdfs
async def extract_tables_from_pdfs(folder_path):
    tables_list = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            tables =  tabula.read_pdf(file_path, pages="all")
            for table in tables:
                tables_list.append(table)
    return  tables_list

#appel à la fonction
# tables_list =  extract_tables_from_pdfs(folder_path)

tables_list = asyncio.run(extract_tables_from_pdfs(folder_path))

#creation d'un dossier excel des tables_list
#creation d'un dossier excel des tables_list
def create_excel_folder(excel_folder):
    if not os.path.exists(excel_folder):
        os.makedirs(excel_folder)
    return excel_folder

#appel à la fonction
excel_folder = create_excel_folder("QCM_excels")

#enregistrement des tables dans des fichiers excels
def save_tables_as_excel(tables_list, excel_folder):
    for i, table in enumerate(tables_list):
        file_path = os.path.join(excel_folder, 'QCM_table{}.xlsx'.format(i))
        table.to_excel(file_path, index=False)

#appel à la fonction
save_tables_as_excel(tables_list, excel_folder)

#comparer les reponses entre deux fichiers excel et compter le score final
def correct_excel_quiz(student_file, correct_answers_file):
    # lire les fichiers Excel avec Pandas
    student_answers = pd.read_excel(student_file)
    answers = pd.read_excel(correct_answers_file)
    
    # initialiser le compteur de réponses correctes
    correct_count = 0
    
    # boucle sur chaque ligne de la feuille de réponses de l'élève
    for i, row in student_answers.iterrows():
        answer = row['Answer']
        if answer == answers.iloc[i]['Correct Answer']:
            # incrémenter le compteur de réponses correctes
            correct_count += 1
    
    # calculer la note finale sur 20
    final_score = (correct_count) 
    
    return final_score

#appel de cette fonction sera dans la fonction qui suit

#extraire les noms et les notes et les stocker dans un fichier excel
#extraire les noms et les notes et les stocker dans un fichier excel
def extract_data_from_excel(excel_folder, correct_answers_file):
    notes = []
    names = []

    for filename in os.listdir(excel_folder):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(excel_folder, filename)
            workbook = openpyxl.load_workbook(file_path)
            # Sélection de la feuille de calcul active
            worksheet = workbook.active
            # Définition de la cellule à afficher
            cell = worksheet["C2"]
            # Affichage du contenu de la cellule
            cell.value
            final_score = correct_excel_quiz(file_path, correct_answers_file)

            #liste des noms
            new_name = cell.value
            names.append(new_name)

            #liste des notes
            new_note = final_score
            final_score_formatted = "{0:.2f}".format(final_score)
            notes.append(final_score_formatted)
            
    return names, notes

#appel à la fonction
names, notes = extract_data_from_excel(excel_folder, correct_answers_file)

#affichage des resultats
now = datetime.now()
one_hour_later = now + timedelta(hours=1)
print("Current date and time:", one_hour_later)
count = len(names)
print("Number of elements in the list:", count)
print("Noms:", names)
print("Notes:", notes)

#stockage des resultats dans un fichier excel
def store_results_in_excel(names, notes):
    # Check if the file exists
    if not os.path.exists("resultats.xlsx"):
        # Create a new workbook
        workbook = openpyxl.Workbook()

        # Select the active worksheet
        worksheet = workbook.active

        # Set column names
        worksheet["A1"] = "Noms"
        worksheet["B1"] = "Notes"

        # Save the workbook
        workbook.save("resultats.xlsx")

    # Read the existing Excel file
    df = pd.read_excel("resultats.xlsx")

    # Add data to a DataFrame
    new_data = {"Noms": names,
                "Notes": notes}
    new_df = pd.DataFrame(data=new_data)

    # Concatenate the dataframes
    df = pd.concat([df, new_df], ignore_index=True)

    # Save the data to the Excel file
    df.to_excel("resultats.xlsx", index=False)

#appel de la fonction
store_results_in_excel(names,notes)


#calcul du temps d'execution
end_time = time.time()

elapsed_time = end_time - start_time
print("Elapsed time:", elapsed_time, "seconds")

