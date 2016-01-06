import xlrd
import re
import os
import sys
# xl_path needs to be accessed across a number of functions, as does listObj, which contains a list of items so we don't output dupes.
global xl_path
global listObj
listObj = []

print "[-] IOC_Extractor v0.0.6b"

def createWriteQuery():
	tqCnt = 0
	sqCnt = 0
	# We need to open the conf file with the user queries. Ask for input path to def file or use default.
	Clear()
	print "[-] Do you wish to specify a query definitions file?"
	print "[-] If so, enter the full path to it. If not, leave blank and press enter."
	print "[-] If left blank, IOC_Extractior will use the default .conf file provided."
	def_queryPath = raw_input("[+] Enter full path or leave blank: ")
	if len(def_queryPath) == 0:
		confTarg = os.environ['ProgramFiles']+"\IOC_Extractor\queryDefs.conf"
	else:
		if os.path.exists(def_queryPath) and def_quertyPath.endswith(".conf"):
			confTarg = def_queryPath
		else:
			print "[-] Invalid Path, your queries have not been generated. Try again."
			end = raw_input("[+] Press enter to finish: ")
			sys.exit(0)
	confs = open(confTarg).readlines()
	# Take query items, turn into query serialized OR list.
	global md5q
	md5q = '"'+md5q.replace(',', '" OR "')+'"'
	global IPq
	IPq = '"'+IPq.replace(',', '" OR "')+'"'
	global emailq
	emailq = '"'+emailq.replace(',', '" OR "')+'"'
	global URLq
	URLq = '"'+URLq.replace(',', '" OR "')+'"'
	# Dictionary containing all queries to be filled with serialized OR list.
	queryDict = {}
	# Start iterating over configuration file, populating the query dictionary for further processing.
	for line in enumerate(confs):
		if line[1].strip() == "#--- BEGIN QUERY DEFINITION BLOCK ---":
			blkBeg = line[0]
		if line[1].strip() == "#--- END QUERY DEFINITION BLOCK ---":
			blkEnd = line[0]
	for line in enumerate(confs):
		if (line[0] > blkBeg) and (line[0] < blkEnd):
			# Add a line to the dictionary containing user defined TAP query.
			if line[1].startswith("TAP::"):
				qItem = line[1].split("::")[1]
				queryDict["TAP"+str(tqCnt)] = qItem
				tqCnt+=1
			# Add a line to dictionary containing user defined Splunk query.
			if line[1].startswith("SPLUNK::"):
				qItem = line[1].split("::")[1]
				queryDict["SPLUNK"+str(sqCnt)] = qItem
				sqCnt+=1
	# This is where we replace the "place holder" lines in our user defined queries, with the IOC extracted information.
	for key in queryDict:
		if "md5Obj" in queryDict[key]:
			queryDict[key] = queryDict[key].replace('"md5Obj"', md5q[6:len(md5q)])
			queryFileObjMD5.write(queryDict[key]+'\r\n')
		if "urlObj" in queryDict[key]:
			queryDict[key] = queryDict[key].replace('"urlObj"', URLq[6:len(URLq)])
			queryFileObjURL.write(queryDict[key]+'\r\n')
		if "ipObj" in queryDict[key]:
			queryDict[key] = queryDict[key].replace('"ipObj"', IPq[6:len(IPq)])
			queryFileObjIP.write(queryDict[key]+'\r\n')
		if "emailObj" in queryDict[key]:
			queryDict[key] = queryDict[key].replace('"emailObj"', emailq[6:len(emailq)])
			queryFileObjEmail.write(queryDict[key]+'\r\n')
		# Not supported yet.
#		if "fnObj" in queryDict[key]:
			#queryDict[key] = queryDict[key].replace('"fnObj"', FNq[6:len(FNq)])
#			queryFileObjFN.write(queryDict[key]+'\r\n')

def Clear():
	os.system('cls')

def findWriteMD5(cellValue):
	strItems = ""
	md5list = re.findall(r"([a-fA-F\d]{32})", str(cellValue))
	# If the list has items in it, do stuff.
	if len(md5list) > 0:
		# For each md5, write the md5 down to the output textfile.
		for md5 in md5list:
			if md5 not in listObj:
				md5FileObj.write(md5+"\n")
				listObj.append(md5)
			strItems = ","+md5
			strItems = strItems[1:len(strItems)]
	return(strItems)
	
def findWriteEmail(cellValue):
	strItems = ""
	email_list = re.findall(r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", str(cellValue).replace('[', '').replace(']', ''))
	# If the list has items in it, do stuff.
	if len(email_list) > 0:
		# For each email, write the email down to the output textfile.
		for email in email_list:
			if email not in listObj:
				emailFileObj.write(email+"\n")
				listObj.append(email)
			strItems = ","+email
			strItems = strItems[1:len(strItems)]
	return(strItems)
	
def findWriteIP(cellValue):
	strItems = ""
	ip_list = re.findall(r"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", str(cellValue))
	# If we have items, do stuff.
	if len(ip_list) > 0:
		# For each IP, write it down.
		for ip in ip_list:
			if ip not in listObj:
				ipFileObj.write(ip+"\n")
				listObj.append(ip)
			strItems = ","+ip
			strItems = strItems[1:len(strItems)]
	return(strItems)

def findWriteDomain(cellValue):
	# Define list for checking top level domain endings.
	endUrl = [".com",".org",".net",".int",".edu",".gov",".mil",".arpa",".ac",".ad",".ae",".af",".ag",".ai",".al",".am",".an",".ao",".aq",".ar",".as",".at",".au",".aw",".ax",".az",".ba",".bb",".bd",".be",".bf",".bg",".bh",".bi",".bj",".bm",".bn",".bo",".br",".bs",".bt",".bv",".bw",".by",".bz",".ca",".cc",".cd",".cf",".cg",".ch",".ci",".ck",".cl",".cm",".cn",".co",".cr",".cs",".cu",".cv",".cw",".cx",".cy",".cz",".dd",".de",".dj",".dk",".dm",".do",".dz",".ec",".ee",".eg",".eh",".er",".es",".et",".eu",".fi",".fj",".fk",".fm",".fo",".fr",".ga",".gb",".gd",".ge",".gf",".gg",".gh",".gi",".gl",".gm",".gn",".gp",".gq",".gr",".gs",".gt",".gu",".gw",".gy",".hk",".hm",".hn",".hr",".ht",".hu",".id",".ie",".il",".im",".in",".io",".iq",".ir",".is",".it",".je",".jm",".jo",".jp",".ke",".kg",".kh",".ki",".km",".kn",".kp",".kr",".kw",".ky",".kz",".la",".lb",".lc",".li",".lk",".lr",".ls",".lt",".lu",".lv",".ly",".ma",".mc",".md",".me",".mg",".mh",".mk",".ml",".mm",".mn",".mo",".mp",".mq",".mr",".ms",".mt",".mu",".mv",".mw",".mx",".my",".mz",".na",".nc",".ne",".nf",".ng",".ni",".nl",".no",".np",".nr",".nu",".nz",".om",".pa",".pe",".pf",".pg",".ph",".pk",".pl",".pm",".pn",".pr",".ps",".pt",".pw",".py",".qa",".re",".ro",".rs",".ru",".rw",".sa",".sb",".sc",".sd",".se",".sg",".sh",".si",".sj",".sk",".sl",".sm",".sn",".so",".sr",".ss",".st",".su",".sv",".sx",".sy",".sz",".tc",".td",".tf",".tg",".th",".tj",".tk",".tl",".tm",".tn",".to",".tp",".tr",".tt",".tv",".tw",".tz",".ua",".ug",".uk",".us",".uy",".uz",".va",".vc",".ve",".vg",".vi",".vn",".vu",".wf",".ws",".ye",".yt",".yu",".za",".zm",".zr",".zw",".asp",".aspx",".css",".html",".htm",".xhtml",".jhtml",".jsp",".jspx",".js",".php",".php4",".php3",".phtml",".rhtml",".xml",".rss","?id="]
	# Define list for proper URL beginnings.
	begUrl = ["www", "http", "https"]
	# Perform checks on cell to tell if domain/URL.
	# Note the replaces, some URLs appeared to have square brackets around dots, appearing in URLs without DNS names. There is a lot of checking to be done to try and prevent false positives being read in as URLs from the rest of the cell data such as email addresses, etc.
	cellValue = str(cellValue).replace("[", "").replace("]", "")
	strItem = ""
	if "()" not in cellValue and "[@]" not in cellValue and "Backoff" not in cellValue and "/WEB-INF" not in cellValue and "@" not in cellValue:
		if cellValue.startswith(tuple(begUrl)):
			if cellValue not in listObj:
				domainFileObj.write(cellValue+"\n")
				listObj.append(cellValue)
				return(cellValue)
			# Because it checks for URLs based on beginning OR end, if we grab one that would also be a positive below, I want to filter it out so I add bad data to the end of the value after writing and returning it, ensure no duplicate data in our lists.
				cellValue = cellValue+"asdfasdf"
		if cellValue.endswith(tuple(endUrl)):
			if cellValue not in listObj:
				domainFileObj.write(cellValue+"\n")
				listObj.append(cellValue)
				return(cellValue)

def main():
	# Creating a list object containing the names of all the sheets in the Excel file.
	wksList = xlObj.sheet_names()
	# We need to define strings for each query.
	global md5q
	md5q = ""
	global emailq
	emailq = ""
	global URLq
	URLq = ""
	global IPq
	IPq = ""
	# In this FOR loop, I take the list of sheets and work with the row data from each sheet.
	for sheet in wksList:
		# Create the object xlSheet which reads a particular sheet's data into the xlSheet object.
		xlSheet = xlObj.sheet_by_name(sheet)
		# Get the number of rows contained within the xlSheet object, assign it to the xl_rows object.
		xl_rows = xlSheet.nrows - 1
		# Get the number of columns (when used in conjunction with the rows, cells) within the xlSheet object, assign it to the 	xl_cells object.
		xl_cells = xlSheet.ncols -1
		# Set the counter object base_row.
		base_row = -1
		# For the number of rows in the sheet, work with the row data.
		while base_row < xl_rows:
			# Increment counter.
			base_row+=1
			# Grab a row designated by the base_row counter, read it into the xlRow object.
			xlRow = xlSheet.row(base_row)
			# Set the counter object base_cell.
			base_cell = -1
			# For each row being iterated above, we iterate through that row's cells.
			while base_cell < xl_cells:
				# Increment counter.
				base_cell+=1
				# Assign the value of a specified cell using the row number and column number, to the xl_cell_val object.
				xl_cell_val = xlSheet.cell_value(base_row, base_cell)
				# Run functions to parse out data, write lists and create query lists.
				try:
					md5item = findWriteMD5(xl_cell_val)
					if len(md5item) > 1:
						if md5item not in md5q:
							md5q = md5q+","+md5item
					emailitem = findWriteEmail(xl_cell_val)
					if len(emailitem) > 1:
						if emailitem not in emailq:
							emailq = emailq+","+emailitem
					ipitem = findWriteIP(xl_cell_val)
					if len(ipitem) > 1:
						if ipitem not in IPq:
							IPq = IPq+","+ipitem
					urlitem = findWriteDomain(xl_cell_val)
					if len(urlitem) > 1:
						if urlitem not in URLq:
							URLq = URLq+","+urlitem
				except UnicodeEncodeError:
					pass
				except TypeError:
					pass

# Print cosmetic information for the end-user, get the path to file and read it into an object using XLRD.
print "[-] Enter the absolute path to your input IOC Excel file."
# Going to use a while loop for taking input and opening the Excel file, on success, break the loop and move on.
sane = 0
while sane == 0:
	try:
		xl_path = raw_input("[+] Path: ")
		xlObj = xlrd.open_workbook(xl_path)
		sane = 1
		Clear()
		print "[-] Gathering information and writing to files..."
	except IOError:
		Clear()
		print "[-] Invalid path, try again."

# A series of outfiles are created depending on the data being pulled. We need to create these files for later writing.
md5FileObj = open(xl_path[0:xl_path.rfind('\\')+1]+"MD5list.txt", 'w')
emailFileObj = open(xl_path[0:xl_path.rfind('\\')+1]+"EMAILlist.txt", 'w')
ipFileObj = open(xl_path[0:xl_path.rfind('\\')+1]+"IPlist.txt", 'w')
domainFileObj = open(xl_path[0:xl_path.rfind('\\')+1]+"URLlist.txt", 'w')
# Not supported yet.
#filenameFileObj = open(xl_path[0:xl_path.rfind('\\')+1]+"FILENAMEist.txt", 'w')
queryFileObjMD5 = open(xl_path[0:xl_path.rfind('\\')+1]+"MD5querylist.txt", 'w')
queryFileObjIP = open(xl_path[0:xl_path.rfind('\\')+1]+"IPquerylist.txt", 'w')
queryFileObjURL = open(xl_path[0:xl_path.rfind('\\')+1]+"URLquerylist.txt", 'w')
# Not supported yet.
#queryFileObjFN = open(xl_path[0:xl_path.rfind('\\')+1]+"FILENAMEquerylist.txt", 'w')
queryFileObjEmail = open(xl_path[0:xl_path.rfind('\\')+1]+"EMAILquerylist.txt", 'w')

#Call the functions.
main()
createWriteQuery()
Clear()

# Close each file object, writings it's data to disk.
md5FileObj.close()
emailFileObj.close()
ipFileObj.close()
domainFileObj.close()
queryFileObjMD5.close()
queryFileObjIP.close()
queryFileObjURL.close()
# Not supported yet.
#queryFileObjFN.close()
queryFileObjEmail.close()

# Cosmetic
print "[-] Execution complete."
print "[-] Check your input file's directory for output files."
end = raw_input("[+] Press enter to finish: ")
