import re
import sys
import json
import csv
import pandas as pd
import tika
tika.initVM()
from tika import parser

# Purpose:
#   This script should take a CIS Benchmark PDF file and convert it into a usable spreadsheet with all normally useful fields carried in.
#   Additional artifacts are created: a couple text files and a json file, in case those are more useful to use.
#   Those files also make customizing this script fairly easy.
#       cis_text.txt        -> the output from tika to convert the pdf to text
#       text.txt            -> cis_text.txt with all blank lines removed
#       <outputname>.json   -> the .json output
#       <outputname>.xlsx   -> the Excel format output
#
# To run:
#   pip install pandas
#   pip install tika
#   install java (for tika)
#   pip install openpyxl
#   open powershell as administrator and run the following:
#   python script.py <inputfile> <outputfilename>
#   administrator is probably needed to run the tika server
#
# Issues:
#   1)  Formatting is not perfect. Items with typographical changes introduce extra line breaks. Monospace font areas and code blocks particularly exhibit this.
#       The only solution I have is to output this in a nice format, but then the end user will need to remove extra linebreaks if they desire.
#       I much prefer to remove extra line breaks than to smash all the text together and lose paragraph separations. The paragraph separations aid readability too much.
#       My one exception to this is the Remediation section which almost always has a code block in it anyway.
#   2)  Do not blindly trust this output! Make sure to inspect all output for accuracy to your purposes.
#       For example, 18.10.9.1.8 just happens to have a string at a line break to match the start a new item. It's not worth coding around that.
#   3)  This was only tested on CIS_Microsoft_Windows_11_Enterprise_Benchmark_v2.0.0.pdf
#       This will almost certainly require tweaks for other benchmark files.
#
# Acknowledgements:
#       This script was adapted from: https://github.com/refabr1k/CIS-PDF-to-Excel
#

cispdf, outfile = "",""

if len(sys.argv) < 3:
	print("[!] Please provide input and output filename!")
	print("Usage: python {} <input.pdf> <output>\n".format(sys.argv[0]))
	print("Note: For <output>, no need to provide file extension.")
	exit()
else:
    cispdf = sys.argv[1]
    outfile = sys.argv[2]


# json file - converted CIS benchmark to json format with 
cisjson = "{}.json".format(outfile)
cisexcel = "{}.xlsx".format(outfile)

# cis text output
cistext = 'cis_text.txt'

#---------------------------------------------------
print("[+] Converting '{}' to text...".format(cispdf))
# tika write get text from pdf
raw = parser.from_file(cispdf)
data = raw['content']

print("[+] creating temp text file...")
# write pdf to text
f = open(cistext,'w', encoding='utf-8')
f.write(data)

# Remove blank lines

with open(cistext, 'r', encoding='utf-8') as filer:
	with open('temp.txt', 'w', encoding='utf-8') as filew:
		for line in filer:
			if not line.strip():
				continue
			if line:
				# start writing
				filew.write(line)

#-------------------------------------------------------
				

print("[+] Converting to Json...")
flagStart, flagComplete, flagSkipToPageBreak = False, False, False
flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, False, False, False
cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs = "","","","","","","","",""
listObj = []


with open("cis_text.txt", 'r', encoding='utf-8') as filer:
	for line in filer:
		#if not line.strip():  #original author skipped out of all empty lines, but I want to be aware of them
		#	continue
		#if line.strip():

			x = {} #json object
			if flagSkipToPageBreak:
				if "Page " in line:
					#print("we've been skipping, but now we can recover as we found a page break!")
					flagSkipToPageBreak = False
				else:
					#print("boom keep staying out!")
					continue

			if "Page " in line:
				continue

			if re.match(r"^(\d{1}|\d{2})\.(\d{1}|\d{2})\.(\d{1}|\d{2})", line):    # identified CIS item name			
				cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs = "","","","","","","","",""
				flagStart, flagComplete = True, False
				flagName = True
				#print(line)
				flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = True, False, False, False, False, False, False, False, False

			if flagStart:
				# from here we will be handling what section we're in and turning on and off the appropriate flags
				if "Profile Applicability:" in line:	
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, True, False, False, False, False, False, False, False
				if "Description:" in line:	
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, True, False, False, False, False, False, False
				if "Rationale:" in line:
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, True, False, False, False, False, False
				if "Impact:" in line:
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, True, False, False, False, False
                # On a handful of benchmark items, the string Audit: does appear in the description. So, I'll do this one differently.
                # In restrospect, this may have been a better way to do it anyway.
				if re.match(r"^Audit:", line):
				#if "Audit:" in line:
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, True, False, False, False
				if "Remediation:" in line:
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, True, False, False
				if "Default Value:" in line:
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, False, True, False
				if "References:" in line:
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, False, False, True
				if "CIS Controls:" in line:
					# If we see this string, we want to make sure we're dumping out of our item. Sometimes we get here early due to no References section.
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, False, False, False
					flagComplete = True
					flagSkipToPageBreak = True
				if "This section is intentionally blank and exists to ensure" in line:
					# If we see this string, we want to make sure we're dumping out of our item and resetting everything. This is a blank item.
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, False, False, False
					cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs = "","","","","","","","",""
					# flagComplete = True
					flagStart = False
					continue
				if "This section contains" in line:
					# If we see this string, we want to make sure we're dumping out of our item and resetting everything. This is a section header.
					flagName, flagLevel, flagDesc, flagRation, flagImpact, flagAudit, flagRemed, flagDefval, flagRefs = False, False, False, False, False, False, False, False, False
					cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs = "","","","","","","","",""
					# flagComplete = True
					flagStart = False
					continue


				if flagName:    # Here we are stitching together our entries based on the section we're in
					cis_name = cis_name + line.replace(' (Automated)','').replace('(Automated)','')
				if flagLevel:
					cis_level = cis_level + line.replace('Profile Applicability: \n','').replace('•  ','').replace(' \n','\n')
                    # alt way to do levels - Sometimes I just want to label something Level 1 or Leve 2 and not make a deal of it.
					#if "Level 1 (L1) " in line:
					#	cis_level = "Level 1"					
					#elif "Level 2 (L2) " in line:
					#	cis_level = "Level 2"
				if flagDesc:
					cis_desc = cis_desc + line.replace('Description: \n','')
				if flagRation:
					cis_ration = cis_ration + line.replace('Rationale: \n','')
				if flagImpact:
					cis_impact = cis_impact + line.replace('Impact: \n','')
				if flagAudit:
					cis_audit = cis_audit + line.replace('Audit: \n','')
				if flagRemed:
					cis_remed = cis_remed + line.replace('Remediation: \n','')
				if flagDefval:
					cis_defval = cis_defval + line.replace('Default Value: \n','')
				if flagRefs:
					if re.match(r"^http", line):    # identified CIS item name
						cis_refs = cis_refs + line


				if flagComplete:    # we have a completed CIS benchmark item - let's clean up the fields and write it out
					cis_name = cis_name.replace(' \n\n','').replace('\n\n','').rstrip()
                    # for the applicability level note the lacking space before \n\n and only one \n - we just want one carriage return on these
					cis_level = cis_level.replace('\n\n','XMXM').replace('\n','').replace('XMXM','\n').rstrip()
					cis_desc = cis_desc.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
					cis_ration = cis_ration.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
					cis_impact = cis_impact.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
					cis_audit = cis_audit.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
                    #feel free to uncomment the below to play with the remediations formatting. This is really wonky and just leaving line breaks in seems the most readable
					#cis_remed = cis_remed.replace('\n\n','').replace(' \n','').rstrip()
					#cis_remed = cis_remed.replace(' \n','').rstrip()
                    # use this one instead to preserve all paragraph spacing - I found this less useful in the Remediation section
                    # cis_remed = cis_remed.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
					cis_defval = cis_defval.replace(' \n\n','XMXM').replace('\n','').replace('XMXM','\n\n').rstrip()
					cis_refs = cis_refs.rstrip()

					x['name'] = cis_name
					x['level'] = cis_level
					x['description'] = cis_desc
					x['rationale'] = cis_ration
					x['impact'] = cis_impact
					x['audit'] = cis_audit
					x['remediations'] = cis_remed
					x['default value'] = cis_defval
					x['references'] = cis_refs
					# print(x)
					cis_name, cis_level, cis_desc, cis_ration, cis_impact, cis_audit, cis_remed, cis_defval, cis_refs = "","","","","","","","",""
					flagStart = False
					# parsed = json.loads(x)
					# print(json.dumps(x, indent=4))
					listObj.append(x)

print("[+] Writing to '{}' ...".format(cisjson))
# print(listObj)
# print(len(listObj))

with open(cisjson, 'w') as json_file:
    json.dump(listObj, json_file, 
                        indent=4,  
                        separators=(',',': '))
print("[+] Creating '{}' ...".format(cisexcel))
df_json = pd.read_json(cisjson)
df_json.to_excel(cisexcel)
print("[+] Done!")

#print(d)			

#print(record[0])

# with open('test.csv', 'w') as ofile:
# 	for i in record:
# 		ofile.write("%s\n" % i)
# 	print("Done")