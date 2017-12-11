import json
import xlwt
from pprint import pprint
import re
global sum
def readexcel(file):
	datas=[]
	with open(file,'r',encoding="utf-8-sig") as f:
		data=f.readlines()
	for i in range(len(data)):
		datas.append(json.loads(data[i]))
	return datas

def writeM(data,n):
	global sum
	relation=data['relation']
	Orientation=data['tableOrientation']
	tabletype=data['tableType']#表格类型1
	pagename1=data['pageTitle'][0:12].replace(r"/","")#表格名字2/ \ :<> | *?"
	pagename=re.sub('[/\:<>|*?"]','',pagename1)
	if data['hasHeader']==False:
		header='N'#是否有表头5
		header_row=-1#表头在第几行6
	elif data['hasHeader']==True:
		header='Y'
		header_row=data['headerRowIndex']

	if data['hasKeyColumn']==False:
		key='N'#是否有主键7
		keycolumn=-1#主键在第几列8
	elif data['hasKeyColumn']==True:
		key='Y'
		keycolumn=data['keyColumnIndex']
#	length=len(relation)
	if len(relation)>256:
		length=256
		sum+=1
	else:
		length=len(relation)
	if len(relation[0])>256:
		Rlength=256
		sum+=1
	else:
		Rlength=len(relation[0])
	book=xlwt.Workbook()
	sheet=book.add_sheet('sheet1',cell_overwrite_ok=True)
	if Orientation=='HORIZONTAL':
		for i in range(length):
			for j in range(Rlength):
				sheet.write(i,j,relation[i][j][0:3000])
		row=length#行数3
		col=len(relation[0])#列数4
		book.save("table/"+"%s_%s_%s_%s_%s_%s_%s_%s.xls"%(tabletype,pagename,row,col,header,header_row,key,keycolumn))
	elif Orientation=='VERTICAL':
		for i in range(length):
			for j in range(Rlength):
				sheet.write(j,i,relation[i][j][0:3000])
		row=len(relation[0])#行数4
		col=length#列数3
		book.save("table/"+"%s_%s_%s_%s_%s_%s_%s_%s.xls"%(tabletype,pagename,row,col,header,header_row,key,keycolumn))
if __name__ =="__main__":
	global sum
	sum=0
	datass=readexcel("sample")
	for i in range(len(datass)):
		print("第%d个表格"%i)
		writeM(datass[i],i)
	print("有差错的表格数目:",sum)
	