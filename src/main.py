import argparse
import os
import sys
import win32com.client
import pythoncom
import cv2
import editdistance
import numpy as np
import tensorflow as tf
import pyodbc
import datetime
import pandas as pd
import re
from pdf2jpg import pdf2jpg
import pytesseract
import PIL
tf.reset_default_graph()
import glob,os,shutil
from DataLoader import Batch, DataLoader, FilePaths
from SamplePreprocessor import preprocessor, wer
from Model import DecoderType, Model
from SpellChecker import correct_sentence
from pytesseract import Output

class Handler_Class():
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        decoderType = DecoderType.BestPath
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            subject = mail.Subject
            print(mail.Attachments)
            try:
                match = re.search(r'\D\D\D\D\d\d\d\d\d\d\d.\d\d' +"|"+ r"\D\D\D\D\d\d\d\d\d\d\d\d\d", subject)
                match.group()
                print("yes")
                attachments = mail.Attachments
                attachment = attachments.Item(1)
                print(attachment.FileName)
                attachfile = attachment.FileName
                attachment.SaveAsFile(r'C:\Users\MANOMAY\Desktop\Handwritten\Handwritten_mail_monitor\data\\' + attachment.FileName)
                #cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                print("yes")
                cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                cursor = cnxn.cursor()
                if attachfile in ["Howard, renewal.pdf","DOC006.pdf","20200519143614592.pdf","20200624114243.pdf","2020 APPLICATION TO SAFEPOINT PER THEIR REQUIREMENT.pdf","[Untitled].pdf","[Untitled](1).pdf"]:
                    docname = attachfile
                    main(docname,cursor)
                    
                    
                else:
                    print("else")
                    docname = attachfile
                    dateint = datetime.datetime.now()
                    dt_string = dateint.strftime("%d/%m/%Y %H:%M:%S")
                    dateint1= datetime.datetime.strptime(dt_string, '%d/%m/%Y %H:%M:%S')
                    folderpath = r'C:\Users\MANOMAY\Desktop\Handwritten\Handwritten_mail_monitor\data'
                    filepath = folderpath+"\\"+attachfile       
                    outputpath = os.path.dirname(filepath)+"\Output1"
                    pdffilename = os.path.basename(filepath)
                    print(filepath,outputpath)
                    pdf2jpg.convert_pdf2jpg(filepath, outputpath, dpi=350,pages = "ALL")
                    print("hi")
                    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'    
                    path = os.path.dirname(filepath) +"\\Output1\\"   
                    path = path+ "\\"+pdffilename+"_dir"
                    t= os.path.split(os.path.dirname(filepath))
                    pathout1 = r"C:\Users\MANOMAY\Desktop\Handwritten\Handwritten-Line-Text-Extraction\data\static\\"+pdffilename+"_dir"
                    print(pathout1)
                    pathout = path  
                    files = []
                    #pdf_pages=[]
                    for r, d, f in os.walk(path):
                        for file in f:
                            if '.jpg' in file:
                                files.append(os.path.join(r, file))
                
                
                    t=""
                    #openfilename = filepath.replace(".pdf",".txt")
                    #newfile = open(openfilename,"w+")
                    for f in files:
                        
                        #print(f,"f")                                     					
                        value=PIL.Image.open(f)
                        text = pytesseract.image_to_string(value, config='',lang='eng')
                        t=t+text
                        print(t)
                    try:
                        match = re.search(r'\D\D\D\D\d\d\d\d\d\d\d.\d\d'  +"|"+ r"\D\D\D\D\d\d\d\d\d\d\d\d\d"+"|"+ r"\D\D\D\D\d\d\d\d\d\d\d", t)
                        pol = match.group()
                        policy = pol
                        print(policy)
                    except:
                        policy = "--"
                        print("oooppp")
                    
                    #print("Extracted Policy number For PDF"+ i+" is : " + match.group() )
                    try:
                        m = re.search('Name of Applicant:(.+?)\n'+"|"+"Insured:(.+?)\n"+"|"+"Named Insured:(.+?)\n"+"|"+"Named Insured and Mailing Address:(.+?)\n", t)
                        nameit = m.group()
                        lst1=['Name of Applicant:',"Insured:","Named Insured:",'Named Insured and Mailing Address:']
                        for j in lst1:
                            if j in nameit:
                                    resfile = nameit.replace(j,"")
                    except:
                        resfile = "--"
                    lst = ["Renewal Supplemental Application","Citizens Assumption Policies.","Citizens Assumption Policies","COMMON POLICY CHANGE ENDROSEMENT","ACKNOWLEDGEMENT OF CONSENT TO RATE","INSURANCE COVERAGE NOTIFICATION(S) ","COMMERCIAL PROPERTY POLICY DECLARATIONS"]
                    try:
                        for i in lst:
                            if i in t:
            
                                typedoc = i
                                if i=="Citizens Assumption Policies." or i=="Citizens Assumption Policies":
                                    i= "Renewal Supplemental Application " + i
                                    print(i)
                                    typedoc = i
                                if i=="Renewal Supplemental Application " +"Citizens Assumption Policies." or i=="Renewal Supplemental Application " +"Citizens Assumption Policies":
                                    sf ="RSACAP"
                                elif i== "Renewal Supplemental Application":
                                    sf = "RSA"
                                elif i=="COMMON POLICY CHANGE ENDROSEMENT":
                                    sf="CPCE"
                                elif i=="ACKNOWLEDGEMENT OF CONSENT TO RATE":
                                    sf="ACR"
                                elif i=="INSURANCE COVERAGE NOTIFICATION(S) ":
                                    sf="ICN"
                                elif i=="COMMERCIAL PROPERTY POLICY DECLARATIONS":
                                    sf="CPPD"
                                else:
                                    sf="--"
                    except:
                        typedoc = "--"
                    try:
                        lst1=['Name of Applicant:',"Insured:","Named Insured:",'Named Insured and Mailing Address:']
                        for j in lst1:
                            if j in nameit:
                                resfile = nameit.replace(j,"")                    
                        resfile1 = resfile.replace("\n","")
                        resfile1 = resfile1.replace("/","")
                        resfile1 = resfile1.replace("|","")
                        resfile1 = resfile1.replace("}","")
                        resfile1 = resfile1.replace("[","")
                        
                        #cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                        #sql = "UPDATE SP_PolicyExtraction SET Meta_Data_Extracted = ?,Status = ? WHERE ID = (SELECT max(ID) FROM SP_PolicyExtraction)"
                       # metadata = policy+"_"+resfile1+"_"+typedoc
                    except:
                        resfile1 = "--"
            
                    dateend1 = datetime.datetime.now()
                    dt_strings = dateend1.strftime("%d/%m/%Y %H:%M:%S")
                    dateend=datetime.datetime.strptime(dt_strings, '%d/%m/%Y %H:%M:%S')
                    renamepath = outputpath+"\\"+policy +"_"+resfile1+"_"+sf+".pdf"
                    metdata = policy +"_"+resfile1+"_"+sf+".pdf"
                    if metdata.count("--")==1:
                        rate = "70%"
                    elif metdata.count("--")==2:
                        rate = "50%"
                    elif metdata.count("--")==3:
                        rate = "0%"
                    else:
                        rate ="100%"
                    sql = "INSERT INTO SP_PolicyExtraction(Date,Doc_Name,Origin,File_path,PolicyNumber,Name_Insured,TypeOf_Doc,Meta_Data_Extracted,Date_Completed ,Status,Success_Rate ,file_rename) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
                    val = (dateint1,docname,"Mail",pathout1,policy,resfile1,typedoc,metdata,dateend,"Success",rate,renamepath)
                    print(renamepath,"oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo")
                    cursor.execute(sql, val)
                    cursor.commit()
                    shutil.copy(filepath,outputpath+"\\"+policy +"_"+resfile1+"_"+sf+".pdf")
                    shutil.move(pathout, pathout1)
            except Exception as e: print(e)
def train(self,model, loader):
    """ Train the neural network """
    epoch = 0  # Number of training epochs since start
    bestCharErrorRate = float('inf')  # Best valdiation character error rate
    noImprovementSince = 0  # Number of epochs no improvement of character error rate occured
    earlyStopping = 25  # Stop training after this number of epochs without improvement
    batchNum = 0

    totalEpoch = len(loader.trainSamples)//Model.batchSize 

    while True:
        epoch += 1
        print('Epoch:', epoch, '/', totalEpoch)

        # Train
        print('Train neural network')
        loader.trainSet()
        while loader.hasNext():
            batchNum += 1
            iterInfo = loader.getIteratorInfo()
            batch = loader.getNext()
            loss = model.trainBatch(batch, batchNum)
            print('Batch:', iterInfo[0], '/', iterInfo[1], 'Loss:', loss)

        # Validate
        charErrorRate, addressAccuracy, wordErrorRate = self.validate(model, loader)
        cer_summary = tf.Summary(value=[tf.Summary.Value(
            tag='charErrorRate', simple_value=charErrorRate)])  # Tensorboard: Track charErrorRate
        # Tensorboard: Add cer_summary to writer
        model.writer.add_summary(cer_summary, epoch)
        address_summary = tf.Summary(value=[tf.Summary.Value(
            tag='addressAccuracy', simple_value=addressAccuracy)])  # Tensorboard: Track addressAccuracy
        # Tensorboard: Add address_summary to writer
        model.writer.add_summary(address_summary, epoch)
        wer_summary = tf.Summary(value=[tf.Summary.Value(
            tag='wordErrorRate', simple_value=wordErrorRate)])  # Tensorboard: Track wordErrorRate
        # Tensorboard: Add wer_summary to writer
        model.writer.add_summary(wer_summary, epoch)

        # If best validation accuracy so far, save model parameters
        if charErrorRate < bestCharErrorRate:
            print('Character error rate improved, save model')
            bestCharErrorRate = charErrorRate
            noImprovementSince = 0
            model.save()
            open(FilePaths.fnAccuracy, 'w').write(
                'Validation character error rate of saved model: %f%%' % (charErrorRate*100.0))
        else:
            print('Character error rate not improved')
            noImprovementSince += 1

        # Stop training if no more improvement in the last x epochs
        if noImprovementSince >= earlyStopping:
            print('No more improvement since %d epochs. Training stopped.' %
                  earlyStopping)
            break


def validate(model, loader):
    """ Validate neural network """
    print('Validate neural network')
    loader.validationSet()
    numCharErr = 0
    numCharTotal = 0
    numWordOK = 0
    numWordTotal = 0

    totalCER = []
    totalWER = []
    while loader.hasNext():
        iterInfo = loader.getIteratorInfo()
        print('Batch:', iterInfo[0], '/', iterInfo[1])
        batch = loader.getNext()
        recognized = model.inferBatch(batch)

        print('Ground truth -> Recognized')
        for i in range(len(recognized)):
            numWordOK += 1 if batch.gtTexts[i] == recognized[i] else 0
            numWordTotal += 1
            dist = editdistance.eval(recognized[i], batch.gtTexts[i])
            ## editdistance
            currCER = dist/max(len(recognized[i]), len(batch.gtTexts[i]))
            totalCER.append(currCER)

            currWER = wer(recognized[i].split(), batch.gtTexts[i].split())
            totalWER.append(currWER)

            numCharErr += dist
            numCharTotal += len(batch.gtTexts[i])
            print('[OK]' if dist == 0 else '[ERR:%d]' % dist, '"' +
                  batch.gtTexts[i] + '"', '->', '"' + recognized[i] + '"')

    # Print validation result
    charErrorRate = sum(totalCER)/len(totalCER)
    addressAccuracy = numWordOK / numWordTotal
    wordErrorRate = sum(totalWER)/len(totalWER)
    print('Character error rate: %f%%. Address accuracy: %f%%. Word error rate: %f%%' %
          (charErrorRate*100.0, addressAccuracy*100.0, wordErrorRate*100.0))
    return charErrorRate, addressAccuracy, wordErrorRate


def load_different_image():
    imgs = []
    for i in range(1, Model.batchSize):
       imgs.append(preprocessor(cv2.imread("../data/check_image/a ({}).png".format(i), cv2.IMREAD_GRAYSCALE), Model.imgSize, enhance=False))
    return imgs


def generate_random_images():
    return np.random.random((Model.batchSize, Model.imgSize[0], Model.imgSize[1]))


def infer(model, fnImg,dateint1,docname,typedoc,filepath1,sf,pdffilename,pathout1,path,cursor):
    """ Recognize text in image provided by file path """
    print(fnImg,"ophfjdsgfhjgdfhjgdshjfgvhsdjgfhjdsghjfghjsdgfhjsdgfhjgdd")
    img = preprocessor(cv2.imread(fnImg, cv2.IMREAD_GRAYSCALE), imgSize=Model.imgSize)
    if img is None:
        print("Image not found")

    imgs = load_different_image()
    imgs = [img] + imgs
    batch = Batch(None, imgs)
    recognized = model.inferBatch(batch)  # recognize text

    print("Without Correction", recognized[0])
    print("With Correction", correct_sentence(recognized[0]))
    rec=correct_sentence(recognized[0])
    print(typedoc)
    print(filepath1)
    print(rec)
    print(path)
    #path = path+ "\\"+pdffilename+"_dir"
    t= os.path.split(os.path.dirname(filepath1))
    pathout1 = r"C:\Users\MANOMAY\Desktop\Handwritten\Handwritten-Line-Text-Extraction\data\static\\"+pdffilename+"_dir"
    print(pathout1)
    pathout = path  
   # t= os.path.split(os.path.dirname(filepath1))
    #pathout1 = t[0]+"\\"+"static\\"+pdffilename+"_dir"
    outputpath = os.path.dirname(filepath1)+"\Output1"
    shutil.copy(filepath1,outputpath+"\\" +"--_"+rec+"_"+sf+".pdf")
    shutil.move(pathout, pathout1)
    policy = "--"
    metdata = "--" +"_"+rec+"_"+sf+".pdf"
    dateend2 = datetime.datetime.now()
    dt_strings1 = dateend2.strftime("%d/%m/%Y %H:%M:%S")
    dateend=datetime.datetime.strptime(dt_strings1, '%d/%m/%Y %H:%M:%S')
    renamepath = outputpath+"\\"+policy +"_"+rec+"_"+sf+".pdf"
    sql = "INSERT INTO SP_PolicyExtraction(Date,Doc_Name,Origin,File_path,PolicyNumber,Name_Insured,TypeOf_Doc,Meta_Data_Extracted,Date_Completed ,Status,Success_Rate ,file_rename) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
    val = (dateint1,docname,"Mail",pathout1,policy,rec,typedoc,metdata,dateend,"Success","70%",renamepath)
    print(renamepath,"oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo")
    cursor.execute(sql, val)
    cursor.commit()
    return recognized[0]


def main(docname,cursor):
    print("hiuouououo")
    """ Main function """
    # Opptional command line args
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--train", help="train the neural network", action="store_true")
    parser.add_argument(
        "--validate", help="validate the neural network", action="store_true")
    parser.add_argument(
        "--wordbeamsearch", help="use word beam search instead of best path decoding", action="store_true")
    args = parser.parse_args()

    decoderType = DecoderType.BestPath
    if args.wordbeamsearch:
        decoderType = DecoderType.WordBeamSearch

    # Train or validate on Cinnamon dataset
    if args.train or args.validate:
        # Load training data, create TF model
        loader = DataLoader(FilePaths.fnTrain, Model.batchSize,
                            Model.imgSize, Model.maxTextLen, load_aug=True)

        # Execute training or validation
        if args.train:
            model = Model(loader.charList, decoderType)
            train(model, loader)
        elif args.validate:
            model = Model(loader.charList, decoderType, mustRestore=False)
            validate(model, loader)

    # Infer text on test image
    else:
        print(docname)
        dateint = datetime.datetime.now()
        dt_string = dateint.strftime("%d/%m/%Y %H:%M:%S")
        dateint1= datetime.datetime.strptime(dt_string, '%d/%m/%Y %H:%M:%S')
        folderpath = r'C:\Users\MANOMAY\Desktop\Handwritten\Handwritten_mail_monitor\data'
        filepath1 = folderpath+"\\"+docname    
        outputpath = os.path.dirname(filepath1)+"\Output1"
        pdffilename = os.path.basename(filepath1)
        pdf2jpg.convert_pdf2jpg(filepath1, outputpath, dpi=350,pages = "ALL")
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'    
        path = os.path.dirname(filepath1) +"\\Output1\\"   
        path = path+ "\\"+pdffilename+"_dir"
        t= os.path.split(os.path.dirname(filepath1))
        pathout1 = t[0]+"\\"+"static\\"+pdffilename+"_dir"
        pathout = path  
        files = []
        #pdf_pages=[]
        for r, d, f in os.walk(path):
            for file in f:
                if '.jpg' in file:
                    files.append(os.path.join(r, file))
    
    
        t=""
        #openfilename = filepath.replace(".pdf",".txt")
        #newfile = open(openfilename,"w+")
        for f in files:
            
            #print(f,"f")                                     					
            value=PIL.Image.open(f)
            text = pytesseract.image_to_string(value, config='',lang='eng')
            t=t+text
        lst = ["Renewal Supplemental Application","Citizens Assumption Policies.","Citizens Assumption Policies","COMMON POLICY CHANGE ENDROSEMENT","ACKNOWLEDGEMENT OF CONSENT TO RATE","INSURANCE COVERAGE NOTIFICATION(S) ","COMMERCIAL PROPERTY POLICY DECLARATIONS"]
        try:
            for i in lst:
                if i in t:

                    typedoc = i
                    if i=="Citizens Assumption Policies." or i=="Citizens Assumption Policies":
                        i= "Renewal Supplemental Application " + i
                        print(i)
                        typedoc = i
                    if i=="Renewal Supplemental Application " +"Citizens Assumption Policies." or i=="Renewal Supplemental Application " +"Citizens Assumption Policies":
                        sf ="RSACAP"
                    elif i== "Renewal Supplemental Application":
                        sf = "RSA"
                    elif i=="COMMON POLICY CHANGE ENDROSEMENT":
                        sf="CPCE"
                    elif i=="ACKNOWLEDGEMENT OF CONSENT TO RATE":
                        sf="ACR"
                    elif i=="INSURANCE COVERAGE NOTIFICATION(S) ":
                        sf="ICN"
                    elif i=="COMMERCIAL PROPERTY POLICY DECLARATIONS":
                        sf="CPPD"
                    else:
                        sf="--"
        except:
            typedoc = "--"
        #img = pathout+"//"+"0_"+docname+".jpg"
        #print(img)
        #image = cv2.imread(img)

        #gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        #blur = cv2.GaussianBlur(gray, (3,3), 0)
        #thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        
        # Detect horizontal lines
        #horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (350,1))
        #horizontal_mask = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
        #cv2.imwrite('hori.jpg',horizontal_mask)
        
        #white_pixels = np.array(np.where(horizontal_mask == 255))
        
        #first_white_pixel = white_pixels[:,0]
        #print(first_white_pixel)
        #x = first_white_pixel[0]
        #y= first_white_pixel[1]
         
        #im_crop = image[y:y+160,y:y+1549]
        #scale_percent = 100 # percent of original size
        #width = int(im_crop.shape[1] * scale_percent / 100)
        #height = int(im_crop.shape[0] * scale_percent / 100)
        #dim = (width, height)
        # resize image
        #resized = cv2.resize(im_crop, dim, interpolation = cv2.INTER_AREA)
        #cv2.imwrite("../data/"+docname+"_crop.png",resized)
        #print("../data/"+docname+"_crop.jpg")
        #opimg = path+"//"+docname+"_crop.jpg"
        #opimg = r"C:\Users\MANOMAY\Desktop\Handwritten\Handwritten-Line-Text-Extraction\data\DOC006.png"
        #print(FilePaths.fnAccuracy+"\\"+i.split(".pdf")[0]+".png","ghsgdfsdhjgfhgsc gf sesc ruhawifghug hjghfughgfsigfhjsdghjfghjafghjgf")
        tf.reset_default_graph()
        print(open(FilePaths.fnAccuracy).read())
        model = Model(open(FilePaths.fnCharList).read(),
                      mustRestore=False)
        infer(model, "../data/"+docname.split(".pdf")[0]+".png",dateint1,docname,typedoc,filepath1,sf,pdffilename,pathout1,path,cursor)
        print(open(FilePaths.fnAccuracy).read())
        #model = Model(open(FilePaths.fnCharList).read(),
         #             decoderType, mustRestore=False)
        #infer(model, FilePaths.fnInfer)


def infer_by_web(path, option):
    decoderType = DecoderType.BestPath
    print(open(FilePaths.fnAccuracy).read())
    model = Model(open(FilePaths.fnCharList).read(),
                  decoderType, mustRestore=False)
    recognized = infer(model, path)

    return recognized


if __name__ == '__main__':
    outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)

#and then an infinit loop that waits from events.
    pythoncom.PumpMessages()
