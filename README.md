# Proposal Reviewer Selector
Funders use Elsevier Expert Lookup for a list of reviewers for submitted research proposals. It still takes a lot of time to check whether the selected reviewers from Expert Lookup are already in contact with the funder or whether they have already assessed a proposal from the funder.

With the script a funder can easily filter list of reviewers from Expert Lookup from funder´s database.

You must have a list of reviewers for the relevant proposal as a csv file, downloaded from Expert Lookup. In addition, you use the excel files of proposal information, reviewers suggested and blocked by the applicant and reviewers from the funder's database. These are in total five input files.

Empty examples of files (applicant info, reviewers from Expert Lookup, suggested reviewers, blocked reviewers, database) with their column name is given in the github repository.

## Code

This code is a tool for selecting reviewers for proposals. It allows the user to import various data sources and then process and merge the data to generate a list of suggested reviewers.

## Features
Import proposal information, suggested reviewers, reviewers from a datawarehouse, Expert Lookup reviewers, and non reviewers from Excel or CSV files
Merge and process data to generate a list of suggested reviewers
## Requirements

pandas

numpy

tkinter

nameparser

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
