import os
import xlrd
import xlwt

while True:

	try:
		filename = raw_input("\nInserire il nome essato del file da cui esstrare: ")

		workbook = xlrd.open_workbook(filename+".xlsx")
		
		print "\nIn elaborazione: %s" %(filename+".xlsx")
		
		break
		
	except IOError:
		print "\nQuesto file non esiste. Riprova!"
		
		
	
worksheet = workbook.sheet_by_index(2) # excel page to extract

offset = 1

codice_banca = raw_input("\n\nInsert Codice Banca: ")

print '\n\nInserire il nome prodotto essatamente come in file Excel e Codice Tabella separate tra virgola (,)! \n\nInserire un solo paio a volta! \n\nPremere "Invio" in riga vuota per eseguire lo script.'

print '\nNota importante: I codici tabella non devono contenere estensioni di categoria (Esempio: "_pubblica", "_statale", ecc.), oppure estensioni di assicurazioni (Esempio: "_aviva"). Lo script fa automaticamente questa operazione.'

print '\nEsempio, per il prodotto "CQS - Sigla", digitare "CQS - Sigla, sigla_cqs" e premere "Invio". \nLo script genera automaticamente tutti i prodotti assiocati con questo nome.\n'

codici = {}

def _products (a,b):
	if (a,b) not in codici.items():
		codici[a] = b
	else:
		print "Prodotto gia essistente. Saltato..."

index = 1

while index == 1:
	
	pair = raw_input("--> ")
	
	if len(pair.split(",")) == 2:
	
		codice_tabella = pair.split(",")[1].strip()
		
		prodotto = pair.split(",")[0].strip()
	
		_products (codice_tabella,prodotto)
		
		print "Notato!\n"
			
	elif pair == "":
		index = 0
		print '\nProdotti.txt e Descrizione.txt sono create con successo! \n\nTransferire i contenuti del file "Prodotti.txt" sul file "tabella.php"!'
	
	else:
		print "\nUsare essatamente due variable. Riprova:"
		continue
		
	
rows = []

o_file = {}

for i, row in enumerate(range(worksheet.nrows)):
    if i <= offset:  # skip headers
        continue
    r = []
    for j, col in enumerate(range(worksheet.ncols)):
        r.append(worksheet.cell_value(i, j))
    rows.append(r)
	

def this(a,b):
	if (a,b) not in o_file.items():
		o_file[a] = b
		
	
 #Creating Prodotti File to merge on "tabella.php" file 
prodotto_finale = open(codice_banca.title()+" Prodotti.txt", "w")
	
for x in rows[1:worksheet.nrows]: # range of products on excel file
	if "-" in x[3]:
		imp_min = str(x[3].split("-")[0].strip().replace(".","")).split(",",)[0]
		imp_max = str(x[3].split("-")[1].strip().replace(".","")).split(",",)[0];
		
	elif "PUBBLICA" in x[2] or "STATALE" in x[2] or "PRIVATA" in x[2]:
		imp_max = 0
		imp_min = 0;
	
	else:
		continue;
	
	
	categoria = str(x[2]).lower().strip(" ").replace(" ","_");
	dur = int(x[4])
	eta_min = int(x[5])
	eta_max = int(x[6])
	_aviva = 12048971
	percentuale = x[8]
	aviva = "aviva"
	
	#Looping through the lines and writing the final file
	#New products can be added below
	
	for cod_tab,descrz in codici.items():
		if descrz == x[0].strip() and x[2].strip() == "" and x[7] != _aviva:
			prodotto_finale.write( 'array("%s","%s","%s","%s","%s","%s","%s"),\n' %(cod_tab,imp_min,imp_max,dur,eta_min,eta_max,percentuale))
			
			this(cod_tab,descrz)
			
		if descrz == x[0].strip() and x[2].strip() == "" and x[7] == _aviva:
			prodotto_finale.write( 'array("%s_%s","%s","%s","%s","%s","%s","%s"),\n' %(cod_tab,aviva,imp_min,imp_max,dur,eta_min,eta_max,percentuale))
			
			this(("%s_%s") %(cod_tab,aviva),("%s %s") %(descrz,aviva.title()))
			
		if descrz == x[0].strip() and x[2].strip() != "" and x[7] != _aviva:
			prodotto_finale.write( 'array("%s_%s","%s","%s","%s","%s","%s","%s"),\n' %(cod_tab,categoria,imp_min,imp_max,dur,eta_min,eta_max,percentuale))
			
			this(("%s_%s") %(cod_tab,categoria),("%s %s") %(descrz,categoria.title()))
		
		if descrz == x[0].strip() and x[2].strip()!= "" and x[7] == _aviva:
			prodotto_finale.write( 'array("%s_%s_%s","%s","%s","%s","%s","%s","%s"),\n' %(cod_tab,categoria,aviva,imp_min,imp_max,dur,eta_min,eta_max,percentuale))
			
			this(("%s_%s_%s") %(cod_tab,categoria,aviva),("%s %s %s")%(descrz,categoria.title(),aviva.title()))
			

#Creating Descrizione file for the final client

descri = open(codice_banca.title()+" Descrizione.txt", "w")
for a,b in o_file.items():
	descri.write ('"%s" \t per il prodotto \t "%s" \n' %(a,b))
descri.close()

raw_input("Premi Invio per chiudere. :)")