from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox
import json
import requests
import win32api
import win32print


#Checks if all require fields have informations in.
def confirmItems():

    if dlg.RecName.text() == "":
        msg = QMessageBox()
        msg.setWindowTitle("No Name")
        msg.setText("A Recipient Name is Mandatory")
        msg.exec() 
        dlg.RecName.setFocus()
    elif dlg.Add1.text() == "":
        msg = QMessageBox()
        msg.setWindowTitle("Address Line 1")
        msg.setText("Address Line 1 is Mandatory")
        msg.exec()
        dlg.Add1.setFocus()
    elif dlg.Town.text() == "":
        msg = QMessageBox()
        msg.setWindowTitle("No Town")
        msg.setText("Town is Mandatory")
        msg.exec()
        dlg.Town.setFocus()
    elif dlg.PostCode.text() == "":
        msg = QMessageBox()
        msg.setWindowTitle("No Postcode")
        msg.setText("PostCode is Mandatory")
        msg.exec()
        dlg.PostCode.setFocus()

    elif dlg.serviceSelector.currentText() == "Next Day":
        msg = QMessageBox()
        msg.setWindowTitle("No Postcode")
        msg.setText(dlg.serviceSelector.currentText())
        msg.exec()
    else:
        createJSON()

    #creates the JSON file that is submitted in the API call.
def createJSON():
    #list.json is a template file which as the authorisation and the itemdetails already set up   
    with open('list.json','r') as file:
        data = json.load(file)

    data["ItemDetails"].append({'RecipientName':dlg.RecName.text(),'RecipientCompany':dlg.Company.text(),'CarrierName':'CMS Network','ClientShipmentReference':'PlaceHolder','ClientName':'CMS TEST POSTER 1',
                                'RecipientAddress':{'AddressLine1':dlg.Add1.text(),'AddressLine2':dlg.Addr2.text(),'AddressLine3':dlg.Addr3.text(),'AddressLine4':dlg.Town.text(),'PostCode':dlg.PostCode.text()},
                                'Weight':20,'ItemHeight':40,'ItemWidth':40,'ItemLength':40,'CarrierServiceCode':'CMSDNDCMS','DeliveryTime':'Next Day by 9AM'})

    #create a new json file with the information entered in the form
    with open('list2.json','w') as file:
        json.dump(data,file)

    #calls the function to create the label
    submitToDocketHub()

#Uses the API details and the JSON file created about to enter the shipment into the docket hub test site.
def submitToDocketHub():

    contents = open('list2.json', 'rb').read()
    headers = {'Content-Type' : 'application/json', 'Accept' : 'application/json'}
    url = 'https://dhuat1services.dockethubtest.com/Mosaic.DocketHub.ItemAdvice/ItemAdviceServiceV3.svc/restService/SubmitItemAdvice'
    response = requests.post(url , data=contents, headers=headers)

    #if the status code is 200 then a connection was accepted.
    if response.status_code == 200:

        jsonresp = response.json()

        errCode = json.dumps(jsonresp[0]['ErrorCode'])

        #If the Json response gives an error code of two then there was an error creating the booking.
        if errCode == "2" or errCode == "1":

            errMessage = json.dumps(jsonresp[0]['ErrorMessage'])
            msg = QMessageBox()
            msg.setWindowTitle("Error With Booking")
            msg.setText(errMessage)
            msg.exec()

        else:
            #gets the reference from the JSON response and displays the reference.
            label = json.dumps(jsonresp[0]['CarrierItemReference'])
            newLabel = label.strip('"')
            '''
            msg = QMessageBox()
            msg.setWindowTitle("Label")
            msg.setText(newLabel)
            msg.exec()
            '''
            #uses the reponse to access dockethub and get the label for the shipment just created.
            url2 = 'https://dhuat1services.dockethubtest.com/Mosaic.DocketHub.ItemAdvice/ItemAdviceServiceV3.svc/restService/ShipmentLabel/'+newLabel
            r = requests.get(url2, auth=('12263labelwebservice','12263lab3ls'))

            #Downloads the label to the PC
            open(newLabel+'.pdf', 'wb').write(r.content)

            #Prints the label that was just downloaded to the default printer.
            filename = newLabel+".pdf"
            currentprinter = win32print.GetDefaultPrinter()

            win32print.SetDefaultPrinter(currentprinter)
            win32api.ShellExecute(0, "print", filename, None,  ".",  0)
            win32print.SetDefaultPrinter(currentprinter)

    #Displays the error code if there was an error connecting to the docket hub API
    else:
        msg = QMessageBox()
        msg.setWindowTitle("Status Code")
        msg.setText(response.status_code)
        msg.exec()
 

app = QtWidgets.QApplication([])
dlg = uic.loadUi("APIBooking.ui")

dlg.pushButton.clicked.connect(confirmItems)

dlg.show()
app.exec()
