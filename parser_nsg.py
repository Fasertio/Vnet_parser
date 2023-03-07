##parser info
import json
import xlsxwriter
import re
import glob, os

def main():

    #recupera i documenti json contenuti nella cartella
    os.chdir(".")
    file_name = []
    for file in glob.glob("*.json"):
        file_name.append(file)

    wk = xlsxwriter.Workbook('info_ngs.xlsx')

    ws = wk.add_worksheet('NSG')
    #crea foglio 3
    nsg(ws, file_name)

    wk.close()

def nsg(ws,file_name):
    header(['Name','Tags','Application','Environment','ManagedBy','Security Property Name','Protocol','sourcePortRange','destinationPortRange','sourceAddressPrefix','destinationAddressPrefix','Access','Priority','direction','sourcePortRanges','destinationPortRanges','sourceAddressPrefixes','destinationAddressPrefixes'],ws)
    row = 1
    column = 0

    for f in file_name:
        fp = open(f)
        p = json.load(fp)
        
        if "securityRules" in p["properties"]:
            for c in p["properties"]["defaultSecurityRules"]:
                payload = []
                payload.append(p["name"])
                if "tags" in p and len(p["tags"]) !=0 :
                    payload.append("Yes")
                    if "Application" in p["tags"] and len(p["tags"]["Application"])  != 0:
                        payload.append(p["tags"]["Application"])
                    else:
                        payload.append("-")  
                    if "Environment" in p["tags"] and len(p["tags"]["Environment"]) != 0:
                        payload.append(p["tags"]["Environment"])
                    else:
                        payload.append("-") 
                    if "ManagedBy" in p["tags"] and len(p["tags"]["ManagedBy"]) != 0:
                        payload.append(p["tags"]["ManagedBy"])
                    else:
                        payload.append("-")   
                else:
                    payload.append("No")
                    payload.append("-")
                    payload.append("-")
                    payload.append("-")

                


                '''
                sourcePortRange	
                #destinationPortRange

                #sourceAddressPrefix
                destinationAddressPrefix
                

                sourcePortRanges	
                destinationPortRanges	
                
                sourceAddressPrefixes	
                destinationAddressPrefixes



                '''


                payload.append(c["name"])
                payload.append(c["properties"]["protocol"])
                
                if "sourcePortRange" in c["properties"] and len(c["properties"]["sourcePortRange"]) != 0:
                    payload.append(c["properties"]["sourcePortRange"])
                else:
                    payload.append("-")
                if "destinationPortRange" in c["properties"] and len(c["properties"]["destinationPortRange"]) != 0:
                    payload.append(c["properties"]["destinationPortRange"])
                else:
                    payload.append("-")
                if "sourceAddressPrefix" in c["properties"] and len(c["properties"]["sourceAddressPrefix"]) != 0:
                    payload.append(c["properties"]["sourceAddressPrefix"])
                else:
                    payload.append("-")
                if "destinationAddressPrefix" in c["properties"] and len(c["properties"]["destinationAddressPrefix"]) != 0:
                    payload.append(c["properties"]["destinationAddressPrefix"])
                else:
                    payload.append("-")
                payload.append(c["properties"]["access"])
                payload.append(c["properties"]["priority"])
                payload.append(c["properties"]["direction"])
                
                if "sourcePortRanges" in c["properties"] and len(c["properties"]["sourcePortRanges"]) != 0:
                    payload.append(' \n'.join(c["properties"]["sourcePortRanges"]))
                else:
                    payload.append("-")
                if "destinationPortRanges" in c["properties"] and len(c["properties"]["destinationPortRanges"]) != 0:
                    payload.append(' \n'.join(c["properties"]["destinationPortRanges"]))
                else:
                    payload.append("-")
                if "sourceAddressPrefixes" in c["properties"] and len(c["properties"]["sourceAddressPrefixes"]) != 0:
                    payload.append(' \n'.join(c["properties"]["sourceAddressPrefixes"]))
                else:
                    payload.append("-")
                if "destinationAddressPrefixes" in c["properties"] and len(c["properties"]["destinationAddressPrefixes"]) != 0:
                    payload.append(' \n'.join(c["properties"]["destinationAddressPrefixes"]))
                else:
                    payload.append("-")
                
                '''
                Name
                Application
                Environment
                ManagedBy
                Security Property Name
                Protocol
                sourcePortRange
                destinationAddressPrefix
                Access
                direction
                sourcePortRanges
                destinationPortRanges
                sourceAddressPrefixes
                destinationAddressPrefixes
                '''
                row = devolve(row,column, payload, ws)

#stampa tutta la riga e muove il cursore sulla riga successiva
def devolve(row, column, content,wk):
    for item in content:
        wk.write(row,column,item)
        column += 1
    row +=1
    return row

#scrive l'header dell'excel
def header(hlist, ws):
    row = 0
    column = 0
    for item in hlist:
        ws.write(row,column,item)
        column += 1



if __name__ == "__main__": 
    main()