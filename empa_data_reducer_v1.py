# -*- coding: utf-8 -*-
"""Import and sort EMPA data with measurements in rows.

@author: Christopher Beyer
"""
import pandas as pd
import os

pd.options.mode.chained_assignment = None  # default='warn'

src = input("Paste source file address with \ replaced by /: ")
dest = input("Specify source folder for output: ")
sample = input("Please specify the sample you want to analyze: ")

raw_data = pd.read_excel(src)

# Select rows of the specific sample
sample_o = raw_data[raw_data["Comment"].str.contains(sample, na=False)]

# Ask user to specify Totals range in absolute deviation
dev = float(input("Specify maximum deviation"
                 " from 100% totals in x.x%: "))

# Remove data that does not match the asked Totals
sample_o.drop(sample_o[(sample_o.Total < (100 - dev)) |
        (sample_o.Total > (100 + dev))].index, inplace=True)

# Creates general statisitcs for each mineral dataframe.
select_min = []

def min_stat(select_min):
    select_min.loc["mean"] = select_min.mean()
    select_min.loc["stdev"] = select_min.std()
    select_min.loc["mean", "Total"] = select_min.loc[
                   "mean", "SiO2":"Na2O"].sum()

class Minerals():
    """Uses user set compostitional boundaries to classify and
       sort minerals.

       Parameters:
           Si_max: Maximum conc. of SiO2 in wt.%
           Si_min: Minimum conc. of SiO2 in wt.%
           MgFe: Minimum sum of FeO and MgO in wt.%
           Al: Minimum conc. of Al2O3 in wt.%
           Ca: Minimum conc. of CaO in wt.

       Methods:
           get_min: Selects the mineral based on the user set
           parameters and creates a new dataframe.

       Returns:
           oxides:
    """
    def __init__(self, Si_max, Si_min, MgFe, Al, Ca):
        self.Si_max = Si_max
        self.Si_min = Si_min
        self.MgFe = MgFe
        self.Al = Al
        self.Ca = Ca
        self.dival = MgFe + Ca

    @property  # decorator sets method as attribute
    def get_min(self):
        oxides = sample_o
        oxides = oxides.loc[(oxides.SiO2 < self.Si_max) & (
                (oxides.FeO + oxides.MgO) > self.MgFe) & (
                (oxides.CaO + oxides.FeO + oxides.MgO)
                > self.dival) & (oxides.CaO > self.Ca) & (oxides.Al2O3
                > self.Al) & (oxides.SiO2 > self.Si_min)]
        min_stat(oxides)

        return oxides

# Defines minerals by compositonal boundaries, see class Minerals.
ol = Minerals(41, 38, 55, 0, 0)
grt = Minerals(54, 39, 10, 5, 1)
opx = Minerals(60, 50, 40, 0, 0)
qtz = Minerals(100, 95, 0, 0, 0)

unknown = sample_o.loc[~sample_o.index.isin(ol.get_min.index
                                           | grt.get_min.index
                                           | opx.get_min.index
                                           | qtz.get_min.index)]       


# get number of copied data sets for each sheet
ol_len = len(ol.get_min.index) - 2
grt_len = len(grt.get_min.index) - 2
opx_len = len(opx.get_min.index) - 2
qtz_len = len(qtz.get_min.index) - 2
remainder = len(unknown)

# save to excel in individual sheets within a given workbook.
i = 0
while os.path.exists("%s-%s.xlsx" % (sample, i)):
    i += 1

writer = pd.ExcelWriter('%s-%s.xlsx' % (sample, i), engine='xlsxwriter')
ol.get_min.to_excel(writer, "ol %s" % ol_len)
grt.get_min.to_excel(writer, "grt %s" % grt_len)
opx.get_min.to_excel(writer, "opx %s" % opx_len)
qtz.get_min.to_excel(writer, "qtz %s" % qtz_len)
unknown.to_excel(writer, "unknown")
writer.save()

print("%s measurements were not categorized" % remainder)
