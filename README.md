# Kickstarter backers export for Stamps.com shipping labels

This is a work in progress, but has been used with one Kickstarter

## Getting Started
Export your backers as as .csv file in USPS format and run MakeSlips.py
Export the backers in Kickstarter format and run CountryCodeToName.py

Both of these are Python3 scripts. Open them up to see the configuration options at the top.

Each shipping weight/package type imported into Stamps.com currently needs to re-run CountryCodeToName.py 
with the relevant LxWxH and weight paramaters.


