import xml.etree.ElementTree as ET
import pandas as pd
import os
from datetime import datetime

if __name__ == "__main__":
    # Demande à l'utilisateur de fournir le chemin du fichier XML
    print("If the output excel is open please close it before launch the script")
    xmlROW_file_path = input("Entrez le chemin du fichier XML ROW: ")

    # Parsing XML
    tree = ET.parse(xmlROW_file_path)
    root = tree.getroot()
    membersToRemove = ["et4ae5", "due__", "Google", "OSF_", "twilio", "odigo", "Didenjoy", "IndividualEmailResult", 
                       "ServiceTerritory", "ServiceResource", "SocialPost", "ServiceAppointment", "ResourceAbsence"
                       ,"WorkType", "SiqUserBlacklist", "SocialPost", "StreamActivity", "TableauHost", "UserEmailPreferred",
                       "VideoCall", "VoiceCall", "Waitlist"]
    membersList = []
    finalGrandList = []
    finalGrandListMembers = []
    statusMembersList = []

    df = pd.DataFrame()

    # Chaque type est défini par un name
    for types in root:
        nameTypeList = ''
        membersList = []
        for elem in types:
            flag = 0
            elem_name = elem.tag.split('}')[-1]
            if elem_name == 'members':
                membersList.append(elem.text)
                for indesirables in membersToRemove:
                    if indesirables in elem.text:
                        flag = 1 
                statusMembersList.append("Expected OK" if flag == 0 else "Not supported in SFOA")
            if elem_name == 'name':
                nameTypeList = elem.text

        if nameTypeList == 'EmailTemplate':
            statusMembersList[-len(membersList):] = ["Data Scope"] * len(membersList)

        if nameTypeList == 'Report':
            statusMembersList[-len(membersList):] = ["Scope To Define"] * len(membersList)

        if nameTypeList == 'Dashboard':
            statusMembersList[-len(membersList):] = ["Scope To Define"] * len(membersList)

    # Construction de finalGrandList après que statusMembersList a été construit 
        for member in membersList:
            finalGrandList.append(member)
            finalGrandListMembers.append(nameTypeList)

    xmlSFOA_file_path = input("Entrez le chemin du fichier XML SFOA: ")
    # Parsing XML
    tree = ET.parse(xmlSFOA_file_path)
    root = tree.getroot()
    membersListSFOA = []
    SFOA_isInROW = []

    for types in root:
        for elem in types:
            flag = 0
            elem_name = elem.tag.split('}')[-1]
            if elem_name == 'members':
                    membersListSFOA.append(elem.text)

    for temp in finalGrandList:
        if temp in membersListSFOA:
            SFOA_isInROW.append("In SFOA")
        else:
            SFOA_isInROW.append("Not in SFOA")
            

    # Création d'un DataFrame Pandas
    print("finalGrandList: " + str(len(finalGrandList)))
    print("finalGrandListMembers: " + str(len(finalGrandListMembers)))
    print("statusMembersList: " + str(len(statusMembersList)))
    print("SFOA_isInROW: " + str(len(SFOA_isInROW)))
    
    #Create the frame excel
    df = pd.DataFrame({"Members": finalGrandList, "Type": finalGrandListMembers, "SFOA Status": statusMembersList, "SFOA package": SFOA_isInROW}) #, "Deployment Status": df['Deployment Status']

    # Comparaison entre SFOA Status et SFOA Package pour créer la nouvelle colonne
    df['Deployment Status'] = 'OK'  # Initialise la colonne à "OK"

    # Conditions pour mettre à jour la colonne en fonction des comparaisons
    df.loc[(df['SFOA Status'] == 'Expected OK') & (df['SFOA package'] == 'Not in SFOA'), 'Deployment Status'] = 'NOT OK'
    df.loc[(df['SFOA Status'] == 'Out of Scope') & (df['SFOA package'] == 'In SFOA'), 'Deployment Status'] = 'NOT OK'
    df.loc[(df['SFOA Status'] == 'Not supported in SFOA') & (df['SFOA package'] == 'In SFOA'), 'Deployment Status'] = 'NOT OK'
    df.loc[(df['SFOA Status'] == 'Data Scope') & (df['SFOA package'] == 'In SFOA'), 'Deployment Status'] = 'NOT OK'




    # Chemin du fichier Excel sur le bureau
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_excel_path = os.path.join(desktop_path, "CompareXML.xlsx")

    # Vérifier si le fichier Excel existe
    if os.path.exists(output_excel_path):
        # Charger le fichier Excel existant avec les feuilles existantes
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a') as writer:
            # Ajouter le DataFrame comme une nouvelle feuille
            unique_suffix = 'Sheet' + str(datetime.now().strftime("%d%H%M%S"))
            df.to_excel(writer, sheet_name=unique_suffix , index=False)
    else:
        # Si le fichier Excel n'existe pas, simplement écrire le DataFrame
        df.to_excel(output_excel_path, index=False)

    print(f"Excel file '{output_excel_path}' updated successfully.")
    # Ouvrir le fichier Excel
    os.system(output_excel_path)