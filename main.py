import xlwings as xw
import pandas as pd
import re
from tabulate import tabulate

#Parse through the sales order form table and print the result
def parse_sales_order_form_table(salesOrderFormTable):
    print("--------------------- Sales Order Form table ---------------------")
    listIterator = 1
    for types in salesOrderFormTable:
        if listIterator < len(salesOrderFormTable):
            print("\n", salesOrderFormTable[listIterator][0], ": ", salesOrderFormTable[listIterator][1])
            if(isinstance(salesOrderFormTable[listIterator][1], str)):
                characters = re.compile(r'\s|-|/')     #Find the special characters (if any) in order to divide the string into subtypes

                if characters.findall(salesOrderFormTable[listIterator][1]):
                    for subtype in re.split('\s|-|/',salesOrderFormTable[listIterator][1]):
                        print("Subtype: ", subtype)
        listIterator += 1
    print("---------------------------------------------------------\n")

#Parse through the contact table and print the result
def parse_contact_table(contactTable):
    print("--------------------- Contact table ---------------------")
    indexIterator = 0
    listIterator = 1
    for types in contactTable[0]:
        print("\n", types, ": ")
        while listIterator < len(contactTable):
            print(contactTable[listIterator][indexIterator])
            if(isinstance(contactTable[listIterator][indexIterator], str)):
                characters = re.compile(r'\s|-|/') 

                if characters.findall(contactTable[listIterator][indexIterator]):
                    for subtype in re.split('\s|-|/',contactTable[listIterator][indexIterator]):
                        print("Subtype: ", subtype)

            listIterator  = listIterator + 1
        indexIterator = indexIterator + 1
        listIterator = 1
    print("---------------------------------------------------------\n")

#Parse through the data table and print the result
def parse_data_table(dataTable):
    print("--------------------- data table ---------------------")
    listIterator = 1
    for types in dataTable:
        if listIterator < len(dataTable):
            print("\n", dataTable[listIterator][0], ": ", dataTable[listIterator][1])
        listIterator += 1
    print("---------------------------------------------------------\n")

#Print the features table
def parse_features_table(featuresTable):
    print("--------------------- features table ---------------------")
    print(featuresTable[0], ": ", featuresTable[1])
    print("---------------------------------------------------------\n")

#Parse through the supplier table and print the result
def parse_supplier_table(supplierTable):
    print("--------------------- Supplier table ---------------------")
    indexIterator = 0
    listIterator = 2
    for types in supplierTable[1]:
        print("\n", types, ": ")
        while listIterator < len(supplierTable):
            print(supplierTable[listIterator][indexIterator])
            if(isinstance(supplierTable[listIterator][indexIterator], str)):
                characters = re.compile(r'\s|-|/')
                if characters.findall(supplierTable[listIterator][indexIterator]):
                    for subtype in re.split('\s|-|/',supplierTable[listIterator][indexIterator]):
                        print("Subtype: ", subtype)

            listIterator  = listIterator + 1
        indexIterator = indexIterator + 1
        listIterator = 2
    print("---------------------------------------------------------")

#Parse through the file table and print the result
def parse_files_table(filesTable):
    print("--------------------- File table ---------------------")
    indexIterator = 0
    listIterator = 2
    for types in filesTable[1]:
        print("\n", types, ": ")
        while listIterator < len(filesTable):
            print(filesTable[listIterator][indexIterator])
            if(isinstance(filesTable[listIterator][indexIterator], str)):
                characters = re.compile(r'\s|-|/')
                if characters.findall(filesTable[listIterator][indexIterator]):
                    for subtype in re.split('\s|-|/',filesTable[listIterator][indexIterator]):
                        print("Subtype: ", subtype)
            listIterator  = listIterator + 1
        indexIterator = indexIterator + 1
        listIterator = 2
    print("---------------------------------------------------------")

#Parse through the quality table and print the result
def parse_quality_table(qualityTable):
    print("--------------------- Quality table ---------------------")
    indexIterator = 0
    listIterator = 2
    for types in qualityTable[1]:
        print("\n", types, ": ")
        while listIterator < len(qualityTable):
            print(qualityTable[listIterator][indexIterator])
            if(isinstance(qualityTable[listIterator][indexIterator], str)):
                characters = re.compile(r'\s|-|/')
                if characters.findall(qualityTable[listIterator][indexIterator]):
                    for subtype in re.split('\s|-|/',qualityTable[listIterator][indexIterator]):
                        print("Subtype: ", subtype)
            listIterator  = listIterator + 1
        indexIterator = indexIterator + 1
        listIterator = 2
    print("---------------------------------------------------------")

if __name__ == "__main__":

    #Open excel from project folder 
    ws = xw.Book("Test_data.xlsx").sheets['Test Data']
    
    #Fetch tables inside the excel file
    salesOrderFormTable     = ws.range("A1:B6").value
    contactTable            = ws.range("A8:C10").value
    dataTable               = ws.range("D2:E5").value
    featuresTable           = ws.range("F1:F2").value
    supplierTable           = ws.range("A12:I14").value
    filesTable              = ws.range("A16:F21").value
    qualityTable            = ws.range("A23:E26").value

    #Parsing functions in order to fetch the wanted data
    parse_sales_order_form_table(salesOrderFormTable)
    parse_contact_table(contactTable)
    parse_data_table(dataTable)
    parse_features_table(featuresTable)
    parse_supplier_table(supplierTable)
    parse_files_table(filesTable)
    parse_quality_table(qualityTable)
    parse_quality_table(qualityTable)
