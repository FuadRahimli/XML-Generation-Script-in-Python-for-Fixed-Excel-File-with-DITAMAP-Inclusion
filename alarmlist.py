from xml.dom.minidom import Document, DOMImplementation
import pandas as pd

def prepare_xml(row):
    imp = DOMImplementation()# DOM is Document Object Model (DOM). It can be used to create new documents, test the support of specific features and create new document types.
    doctype = imp.createDocumentType ("concept", "-//OASIS//DTD DITA Concept//EN", "concept.dtd") #Creating doctype element
    doc = Document() #Once you have a Document instance, you can add elements, attributes, and text to it to create a complete XML document.
    doc.appendChild (doctype) #Appending doctype to document
    concept = doc.createElement ('concept') #Creating concept element
    concept.setAttribute ("id", f"_{row['alarmId ZTS LMS']}") #Filling inside of the concept element with attributes
    title = doc.createElement ("title") #Creating title element
    title_text = doc.createTextNode(f"{row['alarmId ZTS LMS']}") #Filling inside of the title element with text
    title.appendChild(title_text) #Appending title_text to title element
    concept.appendChild(title) #Appending title element to concept element

    conbody = doc.createElement ("conbody") #Creating conbody element

    def create_ptag(key): #There are the same elements. Therefore, we create the function do not repeat each time. Key is the column names.
        p= doc.createElement ('p') #Creating p element
        p_text = doc.createTextNode(str(key)) #Creating p_text which is inside of the p tags
        #str(key) is the column names.Because the column names are string
        p.appendChild(p_text) #Appending p_text to p element
        conbody.appendChild(p) #Appending p element to conbody element

        codeblock = doc.createElement ("codeblock") #Creating codeblock element
        codeblock.appendChild(doc.createTextNode(str(row[key]) if not pd.isnull (row[key]) else '(empty)')) #Firstly creating the codeblock text and then appending to codeblock element 
        #str(row[key]) means the values of columns.
        conbody.appendChild(codeblock) #Appending codeblock element to conbody element


    create_ptag("alarmId NetAct") #Creating alarmID NetAct column
    create_ptag("alarmId ZTS LMS") #Creating alarmID ZTS LMS column
    create_ptag("alarmName") #Creating alarmName column
    create_ptag("alarmType") #Creating alarmType column
    create_ptag("alarmSeverity") #Creating alarmSeverity column
    create_ptag("alarmLongText") #Creating alarmLongText column
    create_ptag("alarmShortText") #alarmShortText column
    create_ptag("cancelling") #Creating cancelling column
    create_ptag("repairAction") #Creating repaitAction column
    create_ptag("resetAlarm") #Creating resetAlarm column
    create_ptag("userActionToBePerformedOnAlarm") #Creating userActionToBePerformedOnAlarm column

    concept.appendChild(conbody) #Appending conbody element to concept element

    doc.appendChild(concept) #Appending concept element to document


    f = open(f"alarm_{row['alarmId NetAct']}.xml", "w", encoding = "UTF-8")
    #The files' names are printed according to alarmID NetAct. 
    # 'w' means, It is possible to write something inside of the file. w means write mode.
    #encoding = 'UTF-8' is encoding the code. Without it, it will give charmap error.
    doc.writexml(f, indent="\t", encoding = "UTF-8", addindent = '\t', newl = '\n')
    f.close()

df = pd.read_excel("CNCS_22_8_Alarm_list.xlsx", sheet_name = "Alarm_list") 
#pd.read_excel is the function of pandas library. 
#sheet_name means, the data is obtained from Alarm_list sheet of excel file


df.apply(prepare_xml, axis = 1)