from MyModules import SAP_Class
from MyModules import WebAutomation
from MyModules import utils

SAP = SAP_Class.VladSAP()


# Get Ship-to from SAP and check if there is some info in the Excel file
def getNameAddress():
    conCountry = SAP.getconCountry()  # this done imiddiatly from first view in del in SAP
    ShipToNr = SAP.getShipTo()
    ConsName, ConsAddress = utils.COO_automation().GetSHipTOinfo(ShipToNr)
    print("Consigne ", ConsName)
    if ConsName == "":
        utils.COO_automation().InpoutBox(str(ShipToNr))
        ConsName, ConsAddress = utils.COO_automation().GetSHipTOinfo(ShipToNr)
    COO.fillConsignee(ConsName, ConsAddress, conCountry)


if __name__ == '__main__':
    # open Delivery
    # DelNr = utils.COO_automation().getDelivery()
    DelNr = ""
    SAP.open_del_03(DelNr)
    # get plant and COO
    plant = SAP.getPlant()
    plantCOO = utils.COO_automation().get_plant_coo(plant)
    # get all goods info from SAP
    goods = SAP.getGoodsInfo()
    # get Sales Order nr
    sonr = SAP.getSOnr()
    # go to WebSite
    COO = WebAutomation.COOsite()
    COO.logingtoSIte()
    # get Consignee details
    getNameAddress()
    #COO.fillCOO(plantCOO)
    # Select SEA transport mode
    COO.fillTranportMode()
    for items in range(0, len(goods), 4):
        COO.fillGoodsDiscrp(goods[items], goods[items + 1], goods[items + 2], goods[items + 3])
    COO.fillRemarks(text=(sonr + " / " + DelNr))
    COO.restofclicks()
    utils.COO_automation().uploadInvoice()
