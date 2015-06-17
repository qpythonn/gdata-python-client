#!/usr/bin/python

# This script is editing an already existing google spreadsheet.
# It loops over a file and parses each line to get a dataset and its corresponding number of events
# At each use you will be prompt the select the google spread sheet you wouldlike to edit and the sheet number of it.  



try:
  from xml.etree import ElementTree
except ImportError:
  from elementtree import ElementTree
import gdata.spreadsheet.service
import gdata.service
import atom.service
import gdata.spreadsheet
import time
import atom
import getopt
import sys
import string


class SimpleCRUD:

  def __init__(self):
    self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
    self.gd_client.email = "qpython@vub.ac.be"
    self.gd_client.password = "Testfaco"
    self.gd_client.source = 'Spreadsheets GData Sample'
    self.gd_client.ProgrammaticLogin()
    self.curr_key = ''
    self.curr_wksht_id = ''
    self.list_feed = None
    
  def _PromptForSpreadsheet(self):
    # Get the list of spreadsheets
    feed = self.gd_client.GetSpreadsheetsFeed()
    self._PrintFeed(feed)
    print  "among the list above, choose the desired document you want to work with by choosing the corresponding number. Be very carefull as the file will be overwritten by the upcoming changes!!!"
#    input = raw_input('\nSelection: ')
    input = raw_input('\nSelection: ')
#    input = str(0) # hardcoded
    id_parts = feed.entry[string.atoi(input)].id.text.split('/')
    self.curr_key = id_parts[len(id_parts) - 1]
  
  def _PromptForWorksheet(self):
    # Get the list of worksheets
    feed = self.gd_client.GetWorksheetsFeed(self.curr_key)
    "among the list above, choose the desired sheet you want to work with by choosing the corresponding number. Be very carefull as the file will be overwritten by the upcoming changes!!!"
    self._PrintFeed(feed)
    input = raw_input('\nSelection: ')
#    input = str(0) # hardcoded  
    id_parts = feed.entry[string.atoi(input)].id.text.split('/')
    self.curr_wksht_id = id_parts[len(id_parts) - 1]
    
#  for it in range (0,10)
  def _PromptForCellsAction(self, row, col, input_value):
    self._CellsUpdateAction(row, col, input_value)
  
  def _PromptForListAction(self):
    print ('dump\n'
           'insert {row_data} (example: insert label=content)\n'
           'update {row_index} {row_data}\n'
           'delete {row_index}\n'
           'Note: No uppercase letters in column names!\n'
           '\n')
    input = raw_input('Command: ')
    command = input.split(' ' , 1)
    if command[0] == 'dump':
      self._ListGetAction()
    elif command[0] == 'insert':
      self._ListInsertAction(command[1])
    elif command[0] == 'update':
      parsed = command[1].split(' ', 1)
      self._ListUpdateAction(parsed[0], parsed[1])
    elif command[0] == 'delete':
      self._ListDeleteAction(command[1])
    else:
      self._InvalidCommandError(input)
  
  def _CellsGetAction(self):
    # Get the feed of cells
    feed = self.gd_client.GetCellsFeed(self.curr_key, self.curr_wksht_id)
    self._PrintFeed(feed)
    
  def _CellsUpdateAction(self, row, col, inputValue):
    entry = self.gd_client.UpdateCell(row=row, col=col, inputValue=inputValue, 
        key=self.curr_key, wksht_id=self.curr_wksht_id)
    if isinstance(entry, gdata.spreadsheet.SpreadsheetsCell):
      print 'Updated!'
        
  def _ListGetAction(self):
    # Get the list feed
    self.list_feed = self.gd_client.GetListFeed(self.curr_key, self.curr_wksht_id)
    self._PrintFeed(self.list_feed)
    
  def _ListInsertAction(self, row_data):
    entry = self.gd_client.InsertRow(self._StringToDictionary(row_data), 
        self.curr_key, self.curr_wksht_id)
    if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
      print 'Inserted!'
        
  def _ListUpdateAction(self, index, row_data):
    self.list_feed = self.gd_client.GetListFeed(self.curr_key, self.curr_wksht_id)
    entry = self.gd_client.UpdateRow(
        self.list_feed.entry[string.atoi(index)], 
        self._StringToDictionary(row_data))
    if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
      print 'Updated!'
  
  def _ListDeleteAction(self, index):
    self.list_feed = self.gd_client.GetListFeed(self.curr_key, self.curr_wksht_id)
    self.gd_client.DeleteRow(self.list_feed.entry[string.atoi(index)])
    print 'Deleted!'
    
  def _StringToDictionary(self, row_data):
    dict = {}
    for param in row_data.split():
      temp = param.split('=')
      dict[temp[0]] = temp[1]
    return dict
  
  def _PrintFeed(self, feed):
    for i, entry in enumerate(feed.entry):
      if isinstance(feed, gdata.spreadsheet.SpreadsheetsCellsFeed):
        print '%s %s\n' % (entry.title.text, entry.content.text)
      elif isinstance(feed, gdata.spreadsheet.SpreadsheetsListFeed):
        print '%s %s %s' % (i, entry.title.text, entry.content.text)
        # Print this row's value for each column (the custom dictionary is
        # built using the gsx: elements in the entry.)
        print 'Contents:'
        for key in entry.custom:  
          print '  %s: %s' % (key, entry.custom[key].text) 
        print '\n',
      else:
        print '%s %s\n' % (i, entry.title.text)
        
  def _InvalidCommandError(self, input):
    print 'Invalid input: %s\n' % (input)
    
  def Run(self):
    f = open('output.txt', 'r')
    self._PromptForSpreadsheet()
    self._PromptForWorksheet()
    #input = raw_input('cells or list? ')
    print "replace cell by cell for now on"
    input='cells'
    if input == 'cells':
      
      # getting time info at the 
      start_date= time.strftime('%d/%m/%Y')
      start_time= time.strftime('%H:%M:%S')

      # erase previous date information
      self._PromptForCellsAction("1","1","This document is BEING edited by a script right now! Please do not modify it!")
      self._PromptForCellsAction("2","1","")

      # edit the "title" of the column
      self._PromptForCellsAction("4","2", "dataset")
      self._PromptForCellsAction("4","11", "nEvents")

      # loop over the list of dataset
      iterator=1
      for line in f:
        # do the parsing
        p1=line.split("dataset ",1)
        p2=p1[1].split(" = ",1)
        p3=p2[0].split("/", 3)
        dataset=p3[1]+"/"+p3[2]
        nEvents=p2[1]
        
        # edit the dataset name column
        self._PromptForCellsAction(str(iterator+4),"2", dataset)
        # edit the nEvents column
        self._PromptForCellsAction(str(iterator+4),"11", nEvents)
        if (iterator % 5 == 0):
          self._PromptForCellsAction("2","1","last change on the "+time.strftime('%d/%m/%Y')+" at "+time.strftime('%H:%M:%S'))
        iterator=iterator+1
        # end of loop

        
      # get time information
      end_date= time.strftime('%d/%m/%Y')
      end_time= time.strftime('%H:%M:%S')

      # print date and time information on the first cell of the document
      self._PromptForCellsAction("1","1","This document has been edited by a script on the "+start_date+" at "+start_time)
      self._PromptForCellsAction("2","1","The last change has been done on the "+end_date+" at "+end_time)

    elif input == 'list':
      while True:
        self._PromptForListAction()


def main():
  # parse command line options  
  key = ''
  # Process options
        
  sample = SimpleCRUD()
  sample.Run()


if __name__ == '__main__':
  main()
