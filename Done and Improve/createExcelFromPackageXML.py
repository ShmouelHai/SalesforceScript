# By Shmouel Illouz : Being lazy is knowing how to write scripts.

import xml.etree.ElementTree as ET
import pandas as pd
import os
from datetime import datetime

if __name__ == "__main__":
    print("If the output excel is open please close it before launch the script")
    xml_file_path = input("Please enter the path of the package XML : ")

    # Parsing XML
    tree = ET.parse(xml_file_path)
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

        # Construction of finalGrandList after that statusMembersList has been constructed
        for member in membersList:
            finalGrandList.append(member)
            finalGrandListMembers.append(nameTypeList)


    print("finalGrandList: " + str(len(finalGrandList)))
    print("finalGrandListMembers: " + str(len(finalGrandListMembers)))
    print("statusMembersList: " + str(len(statusMembersList)))
    
    df = pd.DataFrame({"Members": finalGrandList, "Type": finalGrandListMembers  , "SFOA Status": statusMembersList}) 


    # Output's Path to Desktop
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_excel_path = os.path.join(desktop_path, "Audit.xlsx")

    # Check before if the excel already existing
    if os.path.exists(output_excel_path):
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a') as writer:
            # Add into a new sheet
            unique_suffix = 'Sheet' + str(datetime.now().strftime("%d%H%M%S"))
            df.to_excel(writer, sheet_name=unique_suffix , index=False)
    else:
        # Create a new excel
        df.to_excel(output_excel_path, index=False)

    print(f"Excel file '{output_excel_path}' updated successfully.")
    # Open the Excel
    os.system(output_excel_path)
