from queue import Queue
from threading import Thread

import xlsxwriter
import requests
import operator
import sys
import re
import os


def extract(data):
  url = data[0]
  index = data[1]
  subIndex = 1

  res = []
  infoEnd = []
  dirname = os.path.dirname(url) + '/'

  source = requests.get(url).text

  titleUrl = re.findall('"col-md-12 content"[\s\S]+?a class="" href="(.*?)""[\s\S]+?<h3>(.*?)<\/h3>', source)

  title = []
  urlPage = []

  if len(titleUrl) > 0:  
    for ut in titleUrl:
      urlPage.append(dirname+ut[0])
      title.append(ut[1])

  companyNameMAT = re.findall('Company Name\s?:\s?<\/span>(?!<br>)([\s\S]+?)<\/div>', source)

  companyNameMA = []
  if len(companyNameMAT) > 0:
    for c in companyNameMAT:
      companyNameMA.append(c.replace("<br>", "").strip())

  contactPersonMAT = re.findall('Contact Person\s?:\s?<\/span>(?!<br>)([\s\S]+?)<\/div>', source)

  contactPersonMA = []
  if len(contactPersonMAT) > 0:
    for c in contactPersonMAT:
      contactPersonMA.append(c.replace("<br>", "").strip())

  tel1MA = re.findall('Tel\s?:\s?<\/span>(?!<br>)([\s\S]+?)<\/div>', source)

  tel1Ult = []
  tel2Ult = []

  if len(tel1MA) > 0:
    for t in range(len(tel1MA)):
      if t % 2 == 0:
        tel1Ult.append(tel1MA[t].strip())
      else:
        tel2Ult.append(tel1MA[t].strip())

  if len(urlPage) == 0:
    contactPerson = ""
    tel1 = ""
    tel2 = ""
    cell = ""
    physicalAddress = ""
    res.append((url, "", contactPerson, tel1, tel2, cell, physicalAddress, companyNameMA, contactPersonMA, tel1Ult, tel2Ult, index, subIndex))

    subIndex += 1

  for i in range(len(urlPage)):
    subUrl = urlPage[i]
    source = requests.get(subUrl).text

    contactPerson = re.search("Contact Person:<\/span>([\s\S]+?)<\/p>", source)
    contactPerson = contactPerson[1].strip()

    tel0 = re.findall(">Tel\s?:\s?<\/span>(.*?)<\/p>", source)

    if "None Provided" not in tel0[0]:
      tel1 = re.search("href='tel:(.*?)'", tel0[0])
      tel1 = tel1[1].strip()
    else:
      tel1 = "None Provided"

    if "None Provided" not in tel0[1]:
      tel2 = re.search("href='tel:(.*?)'", tel0[1])
      tel2 = tel2[1].strip()
    else:
      tel2 = "None Provided"

    cell = re.search(">Cell\s?:\s*<\/span>(.*?)<\/p>", source)

    if "None Provided" not in cell[1]:
      cell = re.search("href='tel:(.*?)'", cell[0])
      cell = cell[1].strip()
    else:
      cell = "None Provided"

    physicalAddress = re.search("Physical Address:<\/span>([\s\S]+?)<\/p>", source)

    physicalAddress = physicalAddress[1].replace("<br>", "").replace("\n", "").replace("\r", " ").strip()

    if "," == physicalAddress[-1]:
      physicalAddress = physicalAddress[:-1]

    res.append((url, title[i], contactPerson, tel1, tel2, cell, physicalAddress, companyNameMA, contactPersonMA, tel1Ult, tel2Ult, index, subIndex))

    subIndex += 1

  return res

def runThreads(q, out):
  while not q.empty():
    try:
      u = q.get()
      out.put(extract(u))
      q.task_done()
    except:
      exc_type, exc_obj, exc_tb = sys.exc_info()
      fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
      print(exc_type, fname, exc_tb.tb_lineno)
      q.put(u)
      q.task_done()


if __name__ == "__main__":
  f = open("urls.txt", "r")
  urls = f.read().splitlines()
  f.close()

  q, out = Queue(), Queue()

  index = 1
  for url in urls:
    q.put((url, index))
    index += 1

  num_threads = min(200, q.qsize())
  for i in range(num_threads):
    worker = Thread(target=runThreads, args=(q, out))
    worker.setDaemon(True)
    worker.start()

  q.join()


  workbook = xlsxwriter.Workbook("contractorfind.xlsx")
  worksheet = workbook.add_worksheet("info")

  wrap = workbook.add_format({'text_wrap': True})

  worksheet.write(0, 0, "URL")
  worksheet.write(0, 1, "Name")
  worksheet.write(0, 2, "Contact Person")
  worksheet.write(0, 3, "Tel")
  worksheet.write(0, 4, "Tel")
  worksheet.write(0, 5, "Cell")
  worksheet.write(0, 6, "Physical Address")

  row = 1

  listF = []
  while not out.empty():
    for i in out.get():
      listF.append(i)

  listF = sorted(listF, key = operator.itemgetter(11, 12))

  for r in listF:
    col = 7
    url = r[0]
    title = r[1]
    contactPerson = r[2]
    tel1 = r[3]
    tel2 = r[4]
    cell = r[5]
    physicalAddress = r[6]
    companyNameMA = r[7]
    contactPersonMA = r[8]
    tel1Ult = r[9]
    tel2Ult = r[10]

    worksheet.write(row, 0, url)
    worksheet.write(row, 1, title)
    worksheet.write(row, 2, contactPerson)
    worksheet.write(row, 3, tel1)
    worksheet.write(row, 4, tel2)
    worksheet.write(row, 5, cell)
    worksheet.write(row, 6, physicalAddress)

    for d in range(len(companyNameMA)):
      infoComplete = "Company Name: "+companyNameMA[d]+"\nContact Person: "+contactPersonMA[d]+"\nTel: "+tel1Ult[d]+"\nTel: "+tel2Ult[d]

      worksheet.write(0, col, "More Additions & Alterations Contractors in West Rand")
      worksheet.write(row, col, infoComplete, wrap)

      col += 1

    row += 1

  workbook.close()


