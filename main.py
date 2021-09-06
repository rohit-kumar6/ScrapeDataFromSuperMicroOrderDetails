import rpa as r
import pandas as pd

import tkinter as tk
from tkinter import ttk

import constants

df_list = list()
customerIdList = set()


def listbox_used(event):
    curselection = listbox.curselection()
    global customerIdList
    customerIdList = set()
    for index in curselection:
        customerIdList.add(listbox.get(index))


window = tk.Tk()
window.title("SuperMicro Automation")
window.geometry("450x450")

greeting = tk.Label(text="Hello, User")
greeting.pack()

listbox = tk.Listbox(window, height=8, selectmode='multiple', exportselection=0)
customerIds = ["QT00010N00_3888", "QT00010T00_2888", "QT00010U00_1889", "QT0001BN00_3888", "QT0001BT00_2888",
               "QT0001BU00_1889", "QT0001CT00_2888", "QT0001CU00_1889", "QT0001DU00_1889", "QT0001GU00_1889",
               "QT0001HU00_1889", "QT0001IU00_1889", "QT0001JU00_1889", "QT0001KU00_1889", "QT0001LU00_1889"]
for item in customerIds:
    listbox.insert(customerIds.index(item), item)
listbox.bind("<<ListboxSelect>>", listbox_used)
listbox.pack(padx=10, pady=10)

fromDate = tk.Label(text="Enter From Date")
fromDate.pack()
fromDateEntry = tk.Entry()
fromDateEntry.pack()


def getSoldToAddress():
    soldToAddress = r.read(constants.soldToAddress.format(1))
    soldToAddress += " " + r.read(constants.soldToAddress.format(2))
    return soldToAddress


def getShipToAddress():
    shipToAddress = r.read(constants.shipToAddress.format(1))
    shipToAddress += " " + r.read(constants.shipToAddress.format(2))
    return shipToAddress


def getOrderItems(tempList, orderDetailsList):
    r.click(constants.orderItem)
    r.wait(2)
    headersCount = r.count(constants.orderItemHeaders)
    if headersCount == 0:
        orderDetailsList.append(tempList.copy())
        return

    headersDict = {}
    for k in range(1, headersCount + 1):
        headersDict[r.read(constants.orderItemHeadersValue.format(k))] = k

    rowsCount = r.count(constants.orderItemRows)
    lineNoCol = headersDict.get("Line No.", -1)
    itemNoCol = headersDict.get("Item Number", -1)
    orderDescriptionCol = headersDict.get("Description", -1)
    qtyOrderedCol = headersDict.get("QTY Ordered", -1)
    qtyShippedCol = headersDict.get("QTY Shipped", -1)
    boQtyCol = headersDict.get("B/O QTY", -1)
    unitPriceCol = headersDict.get("Unit Price", -1)
    extendedPriceCol = headersDict.get("Extended Price", -1)

    for row in range(2, rowsCount + 1):
        tempListCopy = tempList.copy()
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, lineNoCol))) if lineNoCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, itemNoCol))) if itemNoCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, orderDescriptionCol))) if orderDescriptionCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, qtyOrderedCol))) if qtyOrderedCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, qtyShippedCol))) if qtyShippedCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, boQtyCol))) if boQtyCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, unitPriceCol))) if unitPriceCol != -1 else tempListCopy.append("")
        tempListCopy.append(r.read(constants.orderItemRowsValue.format(row, extendedPriceCol))) if extendedPriceCol != -1 else tempListCopy.append("")
        orderDetailsList.append(tempListCopy)


def GetClosedOrderDetails(closeOrder, customerId):
    rowCount = r.count(constants.closeOrderTableRowCount)
    if r.present(constants.pageCount):
        rowCount -= 2

    for i in range(2, rowCount + 1):
        tempList = list()
        tempList.append(customerId)
        tempList.append(r.read(constants.closeOrderSalesOrder.format(i)))
        tempList.append(r.read(constants.closeOrderOrderData.format(i)))
        tempList.append(r.read(constants.closeOrderCustomerPO.format(i)))
        tempList.append(r.read(constants.closeOrderAssemblyType.format(i)))
        r.click(constants.closeOrderOrderDetailsButton.format(i))
        r.wait(2)
        tempList.append(r.read(constants.closeOrderOrderStatus))
        tempList.append(getSoldToAddress())
        tempList.append(getShipToAddress())
        getOrderItems(tempList, closeOrder)


def GetOpenOrderDetails(openOrder):
    rowCount = r.count(constants.openOrderTableRowCount)
    if r.present(constants.pageCount):
        rowCount -= 2

    for i in range(2, rowCount + 1):
        tempList = list()
        tempList.append(r.read(constants.openOrderSoldToId.format(i)))
        tempList.append(r.read(constants.openOrderSalesOrder.format(i)))
        tempList.append(r.read(constants.openOrderCustomerPO.format(i)))
        tempList.append(r.read(constants.openOrderShipToParty.format(i)))
        tempList.append(r.read(constants.openOrderShipToCountry.format(i)))
        tempList.append(r.read(constants.openOrderCreatedTime.format(i)))
        r.click(constants.openOrderOrderDetailsButton.format(i))
        r.wait(2)
        tempList.append(r.read(constants.openOrderAssemblyType))
        tempList.append(r.read(constants.openOrderOrderStatus))
        tempList.append(getSoldToAddress())
        tempList.append(getShipToAddress())
        tempList.append(r.read(constants.esd))
        tempList.append(r.read(constants.message))
        getOrderItems(tempList, openOrder)


def GetDataOfCurrentPage(orderType, openOrder, closeOrder, customerId):
    if orderType.__eq__("CustID"):
        GetOpenOrderDetails(openOrder)
    else:
        GetClosedOrderDetails(closeOrder, customerId)


def runAutomation():
    r.init()
    r.url(constants.url)

    while not r.exist(constants.orderType):
        r.wait(1)

    openOrder = list()
    closeOrder = list()
    for orderType in ["CLOSE", "CustID"]:
        r.select(constants.orderType, orderType)
        r.wait(2)
        r.type(constants.fromDate, "[clear]")
        r.wait(2)
        r.type(constants.fromDate, fromDateEntry.get())

        for customerId in customerIdList:
            r.select(constants.customerId, customerId)
            r.wait(2)
            r.click(constants.searchIcon)
            r.wait(2)
            if r.exist(constants.orderTable):
                GetDataOfCurrentPage(orderType, openOrder, closeOrder, customerId)

                curPageNo = 2
                dotClicked = False
                endPageReached = False
                while True:
                    if not r.present(constants.pageCount):
                        break
                    if endPageReached:
                        break
                    r.wait(2)
                    pageCount = r.count(constants.pageCount)
                    for page in range(2, pageCount + 1):
                        pageText = r.read(constants.selectPage.format(page))
                        if page == pageCount and pageText.__eq__("..."):
                            r.click(constants.selectPage.format(page))
                            dotClicked = True
                            break
                        if pageText.__eq__(str(curPageNo)):
                            if not dotClicked:
                                r.click(constants.selectPage.format(page))

                            dotClicked = False
                            curPageNo += 1
                            r.wait(2)
                            GetDataOfCurrentPage(orderType, openOrder, closeOrder, customerId)
                            if page == pageCount:
                                endPageReached = True

    openOrderDf = pd.DataFrame(openOrder, columns=['Sold To ID', 'Sales Order', 'Customer PO', 'Ship To Party',
                                                   'Ship To Country', 'Created Time', 'Assembly Type',
                                                   'Order Status', 'Sold-To', 'Ship-To', 'ESD', 'Message',
                                                   'Line No.', 'Item Number', 'Description', 'QTY Ordered',
                                                   'QTY Shipped', 'B/O QTY', 'Unit Price', 'Extended Price'])
    closeOrderDf = pd.DataFrame(closeOrder, columns=['Sold To ID', 'Sales Order', 'Order Date', 'Customer PO', 'Assembly Type',
                                                     'Order Status', 'Sold-To', 'Ship-To', 'Line No.',
                                                     'Item Number', 'Description', 'QTY Ordered', 'QTY Shipped',
                                                     'B/O QTY', 'Unit Price', 'Extended Price'])

    with pd.ExcelWriter("OrderExport.xlsx") as writer:
        openOrderDf.to_excel(writer, sheet_name="openOrder", index=False)
        closeOrderDf.to_excel(writer, sheet_name="closeOrder", index=False)

    r.close()


button = ttk.Button(text='Start Automation!', command=runAutomation)
button.pack(padx=10, pady=10)
window.mainloop()
