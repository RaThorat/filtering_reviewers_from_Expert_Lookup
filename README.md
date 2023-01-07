# Proposal Reviewer Selector
Funders use Elsevier Expert Lookup for a list of reviewers for submitted research proposals. It still takes a lot of time to check whether the selected reviewers from Expert Lookup are already in contact with the funder or whether they have already assessed a proposal from the funder.

With the script a funder can easily filter list of reviewers from Expert Lookup from funder´s database.

## Preparation 

You must have a list of reviewers for the relevant proposal as a csv file, downloaded from Expert Lookup. In addition, you use the excel files of proposal information, reviewers suggested and blocked by the applicant and reviewers from the funder's database. These are in total five input files.

Empty examples of files (applicant info, reviewers from Expert Lookup, suggested reviewers, blocked reviewers, database) with their column name is given in the github repository.



## Usage

Run the code in IDE to launch the user interface.

Click the "get proposal info" button to import proposal information.

Click the "get suggested reviewers" button to import suggested reviewers.

Click the "get datawarehouse reviewers" button to import reviewers from the datawarehouse.

Click the "get Expert Lookup reviewers" button to import Expert Lookup reviewers.

Click the "get non reviewers" button to import non reviewers.

Then close the window.

The code will automatically process and merge the imported data to generate a list of suggested reviewers. You will get an excel file of the filtered rreviewers in the same folder where you ran the file.

## Notes

The code expects the imported files to be in a specific format. Please refer to the documentation for details on the expected file formats.

The code is currently set up to handle Excel and CSV files. If you need to import data from a different file type, you may need to modify the code to support that file type.

# Code description

This code is a tool for selecting reviewers for proposals. It allows the user to import various data sources and then process and merge the data to generate a list of suggested reviewers.

## Features
Import proposal information, suggested reviewers, reviewers from a datawarehouse, Expert Lookup reviewers, and non reviewers from Excel or CSV files
Merge and process data to generate a list of suggested reviewers

## Requirements

pandas

numpy

tkinter

nameparser

## Steps processing a list of reviewers from Expert Lookup

This code is for processing a list of reviewers for a grant proposal. The input is a dataframe containing the reviewers' names, organizations, and countries (df_RL). The code first replaces any empty values in the dataframe with "nan" and resets the index. It then filters a second dataframe (df_pv3) by the grant number and resets the index. The name of the grant applicant is extracted from the filtered df_pv3 and added to df_RL. The columns in df_RL are then renamed.

The code then uses a library called "HumanName" to extract the title, first name, middle name, and last name for each reviewer and adds these values to df_RL as new columns. The dataframe is then modified to include empty columns for "Rangorde," "URL," and "m/v." Any empty values in df_RL are again replaced with "nan."

A third dataframe (df_SR) containing suggested reviewers for the grant proposal is then filtered by the grant number and the columns are renamed. The code then compares the last names of the reviewers in df_RL to those in df_SR and adds a new column, "Bron," to df_RL indicating whether the reviewer was selected by the applicant or was found through "Expert Lookup." Finally, the columns in df_RL are rearranged and the index is reset.

## Steps processing suggested reviewers list from sources other than Expert Lookup
This code processes a suggested reviewers list and adds it to a reviewers list. It first filters the suggested reviewers list to only include certain columns: 'Dossiernummer', 'Achternaam', 'Voornaam', 'ScopusLink', 'Email', 'Inst', 'Land'.

It then adds several new columns to the suggested reviewers list with constant values or values from the reviewers list, such as 'Bron', 'Rangorde', 'Referent', 'ProposalTitle', 'ProposalLink', 'Applicants', and 'Hoofdaanvrager'. The columns are then rearranged to be in a specific order.

The suggested reviewers list is then added to the reviewers list and any duplicate records are removed. The resulting list is sorted by 'Dossiernummer' and the index is reset.

## Steps cleaning 
The given code is used to compare a reference list with pivot table references from a data warehouse. It performs the following tasks:

1. Changes the column names of the data warehouse table and removes unnecessary columns.

2. Filters out non-reviewers from the data warehouse table.

3. Finds the salutation (Dr. or Prof) of the reviewers in the data warehouse table (assuming only the first salutation is selected).

4. Adds new columns to the data warehouse table for Scopus link, web search link, proposal title, and proposal link.

5. Merges the reference list and data warehouse table based on email.

6. Replaces null values in the 'Eindtotaal' column with 0 for further calculation.

7. Changes the data type of the 'Eindtotaal' column for comparison.

8. Writes in the 'Opmerking' column whether a reviewer is already in the data warehouse or is new.

9. Writes in the 'Opmerking' column whether a reviewer in the data warehouse has worked for the funder in 2021.

10. Removes unnecessary columns from the merged table.

## Steps to identify applicants from previous years

This code is used to identify applicants who have previously applied for a subsidy in 2018, 2019, or 2020. It does this by searching the funders database for common strings of application numbers associated with those years. The resulting data is then merged with another data set and various columns are added or renamed. The resulting data set is used to identify reapplications and other relevant information.

## Steps filtering non reviewers or blocked reviewers
The code is used to extract the names, emails, and affiliations of non-reviewers for a talent program. The data is stored in an excel file and is read into a Pandas dataframe called df_nonref. The names of the non-reviewers are separated from other information such as university name and email address by splitting on commas. The first part of the split is assumed to be the name, and this is then further split into first and last name using the HumanName library. Email addresses are extracted using a regular expression search, and the university name is extracted by searching for keywords in the remaining text after the name and email have been removed. The extracted information is then stored in a new dataframe called df_nRef, which includes columns for proposal number, grant number, last and first name, Scopus Author ID, OrcId, email, affiliation, and country. The grant number is extracted by taking the last digits of the grant number, and the proposal number is extracted by splitting on periods and taking the last element. The dataframe is then exported to an excel file for further use.

## Steps arranging all dataframes into a single dataframe

This code appears to be taking a number of dataframes and combining them into one new dataframe called df_RL3. It then replaces any empty values in the dataframe with NaN, fills any missing values in the 'Bron' column with the string 'Expert Lookup', removes any rows with missing values in the 'Achternaam' column, replaces certain values in the 'Opmerking' column with the string 'New', removes duplicate rows based on the 'Email' column and keeps the last occurrence, and sorts the dataframe by the 'Applicants' column. Finally, it creates a new dataframe called df_RL4 which is a copy of df_RL3 with the index reset.



