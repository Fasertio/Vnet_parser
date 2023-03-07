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
    ws = wk.add_worksheet('Info VNet')
    #crea foglio 1
    vnet(ws,file_name)

    ws = wk.add_worksheet('Subnet')
    #crea foglio 2
    subnet(ws, file_name)

    wk.close()

#crea il foglio 1, vnet con nome, subscription, location ecc...
def vnet(ws, file_name):
    header(['Name','Subscription','Resource group','Location'],ws)    
    row = 1
    column = 0

    for f in file_name:
        fp = open(f)
        p = json.load(fp)
        #/subscriptions/d11fa05e-f0d9-401c-9736-dce7e6b3aff7/resourceGroups/EAITPRSG001-NetworkSecurity/providers/Microsoft.Network/virtualNetworks/EAITPVNT001
        subscription = re.search('/subscriptions/(.*)/resourceGroups/',p["id"]).group(1)
        rg = re.search('/resourceGroups/(.*)/providers',p["id"]).group(1)
        payload = []
        payload.append(p["name"])
        payload.append(subscription)
        payload.append(rg)
        payload.append(p["location"])
        row = devolve(row,column, payload, ws)

#crea il foglio 2 con le subnet e la lista delle informazioni attinenti
def subnet(ws, file_name):
    header(['Subnet Name','Vnet Name','Subscription','Resource group','Subnet Ip','Network Security Groups'],ws)
    row = 1
    for f in file_name:
        fp = open(f)
        p = json.load(fp)
        subscription = re.search('/subscriptions/(.*)/resourceGroups/',p["id"]).group(1)
        rg = re.search('/resourceGroups/(.*)/providers',p["id"]).group(1)
        name_vnet = p["name"]

        #crea una nuova riga ad ogni id di ogni subnet
        for r in p["properties"]["subnets"]:
            column = 0
            payload = []
            payload.append(r["name"])
            payload.append(name_vnet)
            payload.append(subscription)
            payload.append(rg)
            payload.append(r["properties"]["addressPrefix"])
            if "networkSecurityGroup" in r["properties"] and len(r["properties"]["networkSecurityGroup"])  != 0:
                payload.append(re.search('networkSecurityGroups/(.*)', r["properties"]["networkSecurityGroup"]["id"]).group(1))
            else:
                payload.append("-")
            row = devolve(row,column, payload, ws)
        '''
        for r in p["properties"]["subnets"]:
                name = r["name"]
                #crea una nuova riga ad ogni id di ogni subnet
                payload = []
                if "ipConfigurations" in r["properties"]:
                    for c in r["properties"]["ipConfigurations"]:
                        payload.append(name)
                        payload.append(subscription)
                        payload.append(rg)
                        payload.append(c["id"])
                        row = devolve(row,column, payload, ws)
        '''

def nsg(ws,file_name):
    header(['Name','Application','Environment','ManagedBy','Security Property Name','Protocol','sourcePortRange','destinationAddressPrefix','Access','direction','sourcePortRanges','destinationPortRanges','sourceAddressPrefixes','destinationAddressPrefixes'],ws)
    row = 1
    column = 0

    for f in file_name:
        fp = open(f)
        p = json.load(fp)
        #/subscriptions/d11fa05e-f0d9-401c-9736-dce7e6b3aff7/resourceGroups/EAITPRSG001-NetworkSecurity/providers/Microsoft.Network/virtualNetworks/EAITPVNT001
        if "securityRules" in p["properties"]:
            for c in p["properties"]["securityRules"]:
                payload = []
                payload.append(p["name"])
                payload.append(p["tags"]["Application"])
                payload.append(p["tags"]["Environment"])
                payload.append(p["tags"]["ManagedBy"])
                payload.append(c["name"])
                payload.append(c["properties"]["protocol"])
                payload.append(c["properties"]["sourcePortRange"])
                payload.append(c["properties"]["destinationAddressPrefix"])
                payload.append(c["properties"]["access"])
                payload.append(c["properties"]["direction"])
                payload.append(c["properties"]["sourcePortRanges"])
                payload.append(c["properties"]["destinationPortRanges"])
                payload.append(c["properties"]["sourceAddressPrefixes"])
                payload.append(c["properties"]["destinationAddressPrefixes"])
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