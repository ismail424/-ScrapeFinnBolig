
from bs4 import BeautifulSoup
import requests
import xlsxwriter

#File name
file_name = "apartments.xlsx"

#Range of how many pages it will loop through
pages = range(1,5)

#Dont change this :D
row = 1

#Creates the excel file and opens it
workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()

worksheet.write(0,0,'Url')
worksheet.write(0,1,'By')
worksheet.write(0,2,'Navn')
worksheet.write(0,3,'Bolig Address')
worksheet.write(0,4,'Prisantydning')
worksheet.write(0,5,'Fellesgjeld')
worksheet.write(0,6,'Fellesformue')
worksheet.write(0,7,'Felleskost/mnd.')
worksheet.write(0,8,'Totalpris')
worksheet.write(0,9,'Boligtype')
worksheet.write(0,10,'Eieform bolig')
worksheet.write(0,11,'Soverom')
worksheet.write(0,12,'Primærrom')
worksheet.write(0,13,'Bruksareal')
worksheet.write(0,14,'Byggeår')
worksheet.write(0,15,'Energimerking')
worksheet.write(0,16,'Tomteareal')
worksheet.write(0,17,'Bruttoareal')
worksheet.write(0,18,'Boligselgerforsikring')
worksheet.write(0,19,'Fellesformue')
worksheet.write(0,20,'Formuesverdi')
worksheet.write(0,21,"Balkong/Terrasse")
worksheet.write(0,22,"Barnevennlig")
worksheet.write(0,23,"Bredbåndstilknytning")
worksheet.write(0,24,"Garasje/P-plass")
worksheet.write(0,25,"Ingen")
worksheet.write(0,26,"gjenboere")
worksheet.write(0,27,"Kabel-TV")
worksheet.write(0,28,"Offentlig")
worksheet.write(0,29,"vann/kloakk")
worksheet.write(0,30,"Rolig")
worksheet.write(0,31,"Sentralt")
worksheet.write(0,32,"Utsikt")
worksheet.write(0,33,"Vaktmester-/vektertjeneste")
worksheet.write(0,34,"Turterreng")

def remove(string): 
    return "".join(string.split()) 

def get_Info_from_url(url):
    global row
    r = requests.get(url)
    site = BeautifulSoup(r.text, "html.parser")

    name_of_article = site.findAll("span",{"class":"u-t3 u-display-block"})   
    address =  site.findAll("p",{"class":"u-caption"})
    list_address = address[0].text.split()
    city = str(list_address[-1])
    list_address[-1] = ","+list_address[-1]
    new_address = " ".join(list_address)
    
    
    price = site.findAll("span",{"class":"u-t3"})
    
    names = site.findAll("dt")
    values = site.findAll("dd")
    

    listor = site.findAll("ul",{"class":"list list--bullets list--cols1to2 u-mb16"})
    listor = listor[0].text.split()

    for x in listor:
        value = x.strip()
        value = str(value)
        if value == 'Balkong/Terrasse':
            worksheet.write(row,21,"Yes")
        if value == 'Barnevennlig':
            worksheet.write(row,22,"Yes")
        if value == 'Bredbåndstilknytning':
            worksheet.write(row,23,"Yes")    
            
        if value == 'Garasje/P-plass':
            worksheet.write(row,24,"Yes")  
        if value == 'Ingen':
            worksheet.write(row,25,"Yes")  
        if value == 'gjenboere':
            worksheet.write(row,26,"Yes")  
            
        if value == 'Kabel-TV':
            worksheet.write(row,27,"Yes")  
        if value == 'Offentlig':
            worksheet.write(row,28,"Yes")  
        if value == 'vann/kloakk':
            worksheet.write(row,29,"Yes")  
            
        if value == 'Rolig':
            worksheet.write(row,30,"Yes")  
        if value == 'Sentralt':
            worksheet.write(row,31,"Yes")  
        if value == 'Utsikt':
            worksheet.write(row,32,"Yes")  
            
        if value == 'Vaktmester-/vektertjeneste':
            worksheet.write(row,33,"Yes")  
        if value == 'Turterreng':
            worksheet.write(row,34,"Yes")  
    for name in names:
        where_name = names.index(name)
        name  = name.text.strip()
        value =str(values[where_name].text)
        if name == 'Fellesgjeld':
            worksheet.write(row,5,remove(value).replace("kr", ""))
        if name == 'Fellesformue':
            worksheet.write(row,6,remove(value).replace("kr", ""))
        if name == 'Felleskost/mnd.':
            worksheet.write(row,7,remove(value).replace("kr", ""))
        if name == 'Totalpris':
            worksheet.write(row,8,remove(value).replace("kr", ""))
        if name == 'Boligtype':
            worksheet.write(row,9,value)
        if name == 'Eieform bolig':
            worksheet.write(row,10,value)
        if name == 'Soverom':
            worksheet.write(row,11,value)
        if name == 'Primærrom':
            worksheet.write(row,12,remove(value).replace("m²", ""))
        if name == 'Bruksareal':
            worksheet.write(row,13,remove(value).replace("m²", ""))
        if name == 'Byggeår':
            worksheet.write(row,14,value)
        if name == 'Energimerking':
            worksheet.write(row,15,value)
        if name == 'Tomteareal':
            worksheet.write(row,16,remove(value).replace("m²(eiet)", ""))
        if name == 'Bruttoareal':
            worksheet.write(row,17,remove(value).replace("m²", ""))
        if name == 'Boligselgerforsikring':
            worksheet.write(row,18,value)
        if name == 'Fellesformue':
            worksheet.write(row,19,remove(value).replace("kr", ""))        
        if name == 'Formuesverdi':
            worksheet.write(row,20,remove(value).replace("kr", ""))        

    worksheet.write(row,0,url)
    
        
    try:
        worksheet.write(row,1,city)
    except:
        pass
    
    try:
        worksheet.write(row,2,name_of_article[0].text)
    except:
        pass

    
    try:
        worksheet.write(row,2,name_of_article[1].text)
    except:
        worksheet.write(row,2,None)
        pass

    try:
        worksheet.write(row,3,new_address)
    except:
        pass

    try:
        worksheet.write(row,4,price[1].text.strip().replace(u'\xa0', u' ').replace(" ", "").replace("kr", ""))
    except:
        worksheet.write(row,4,price[0].text.strip().replace(u'\xa0', u' ').replace(" ", "").replace("kr", ""))
    finally:
        pass


    row += 1
    return "Done with " + str(url)

pages_to_go = 0
for page in pages:
    r = requests.get("https://www.finn.no/realestate/homes/search.html?page="+ str(page) +"&sort=PUBLISHED_DESC")

    site = BeautifulSoup(r.text, "html.parser")
    articles = site.findAll("a",{"class":"ads__unit__link"})

    for article in articles:
        x = article["href"]
        if x != None:
            try:
                print(get_Info_from_url(str(x)))
            except:
                if "www" not in str(x):
                    try:
                        x = "https://www.finn.no" + str(x)
                        print(get_Info_from_url(x))
                    except:
                        print("\nError at "+ str(x) + "\n")
    pages_to_go += 1
    print("Done with page: "+ str(pages_to_go))
                

try:
    workbook.close()
except:
    print("\nError cant save date into a open excel file\nYou have to close the excel first to make this porgram work\n")

