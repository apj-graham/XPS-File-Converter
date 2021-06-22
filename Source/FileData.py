import os
import pandas as pd
from pyexcelerate import Workbook

class FileData:
    """Class to process and store data in .txt files"""

    def __init__(self, file_path):
        """Initialise key values need to create excel fileS"""
        # Removes file extension
        self.filename = str.split(file_path, "\\")[-1][:-4]
        self.df = pd.read_csv(file_path, sep="\t", header=None)
        self.all_scans = []
        self._get_scan_data()

    def _get_scan_data(self):
        """Get the data from each scan. Discards empty scansS"""
        start_indexes = self._get_start_indexes()
        end_indexes = [idx-1 for idx in start_indexes[1:]]
        end_indexes.append(len(self.df.index)) #set endpoint of last set as the last row in the dataframe
        for start, end in zip(start_indexes, end_indexes):
            if self.df.at[start+3, 0] == "1": #data sets with this value != 1 are empty
                scan = self.df[start: end]#
                scan.fillna(" ", inplace=True)
                self.all_scans.append(scan)

    def _get_start_indexes(self):
        """Finds indexes of the start of datsetS"""
        start_indexes = self.df.index[self.df[0].str.contains("Region")] #each new data set starts with "Region"
        return start_indexes.tolist()

    def save_data_as_xlsx(self):
        count = 1
        wb = Workbook()
        for scan in self.all_scans:
            values = scan.values.tolist()
            wb.new_sheet("Sheet " + str(count), data = values)
            count += 1
        wb.save(self.filename + ".xlsx")
