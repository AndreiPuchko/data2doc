import sys
import os
import glob
import time

sys.path.insert(1, "../")

from data2doc.data2doc import data2doc

                             
testResultFolder=f"../test-result/"
testSourceFolder="../test-data/test*"

d2d=data2doc()

for testFolderName in glob.glob(testSourceFolder):
	print (testFolderName," ==>> ", end = '')
	ti0=time.time()
	for x in glob.glob(f"{testFolderName}/*"):
		d2d.loadFile(x)
	rez=d2d.data2doc()
	if rez:
		print ("Ok")
		tName=os.path.basename(testFolderName)
		i=0
		while True:
			tFileName=f"{testResultFolder}{tName}-{i}.docx"
			tLock=testResultFolder+".~lock."+tName+str(i)+".docx#"
			print (tLock,glob.glob(tLock))
			if glob.glob(tLock):
				i+=1
			else:
				if glob.glob(tFileName):
					os.remove(tFileName)
				break
		open(tFileName,"wb").write(d2d.docxResultBinary)
	else:
		print ("Fail")
print ("found Excel format strings:")
for y in d2d.usedFormatStrings:
	print (f"{y:20}'{d2d.usedFormatStrings[y]:20}' '{d2d.setNmFmt(d2d.usedFormatStrings[y],y)}'")
if sys.platform=="win32":
	os.startfile(testResultFolder.replace("/","\\"))