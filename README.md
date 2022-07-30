# filtering_reviewers_from_Expert_Lookup
Funders use Elsevier Expert Lookup for a list of reviewers for submitted research proposals. It still takes a lot of time to check whether the selected reviewers from Expert Lookup are already in contact with the funder or whether they have already assessed a proposal from the funder.

With the script a funder can easily filter list of reviewers from Expert Lookup from funder´s database.

You must have a list of reviewers for the relevant proposal as a csv file, downloaded from Expert Lookup. In addition, you use the excel files of proposal information, reviewers suggested and blocked by the applicant and reviewers from the funder's database. These are in total five input files.

Empty examples of files (applicant info, reviewers from Expert Lookup, suggested reviewers, blocked reviewers, database) with their column name is given in the github repository:

If you run the code in IDE, a graphic user interface (GUI) appears:
Click on 'get proposal info' to upload the file applicant info.
Click on 'get suggested reviewers' to upload file suggested reviewers by applicant.
Click on 'get data warehouse reviewers' to upload a reviewer database from funder (This may take one or two minutes.)
Click on 'get Expert Lookup reviewers' to upload the csv file from the Expert Lookup.
Click on 'get non reviewers' to upload the non reviewers.
Then close the window.

The script will run itself and then you will get an excel file of the filtered rreviewers in the same folder where you ran the file.
