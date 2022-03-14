# filtering_reviewers_from_Expert_Lookup
Funders use Expert Lookup for a list of reviewers for submitted research proposals. It still takes a lot of time to check whether the selected reviewers from Expert Lookup have already in contact with the funder or whether they have already assessed a proposal.  With this code a funder can easily filter list of reviewers. 

You must have a list of reviewers for the relevant proposal, csv file, downloadable from Expert Lookup. In addition, you use the excel files about proposal info, reviewers suggested and blocked by the applicant and reviewers from the funder datawarehouse-database. So it's a total of five input files.

Empty example of files (applicant info, reviewers from Expert Lookup, suggested reviewers, blocked reviewers, dwh database) with their column name is given. 

If you run the code in IDE, a graphic user interface (GUI) appears:
Click on 'get proposal info' to upload the file applicant info.
Click on 'get suggested reviewers' to upload file suggested reviewers by applicant.
Click on 'get data warehouse reviewers' to upload a reviewer database from funder (This may take one or two minutes.)
Click on 'get Expert Lookup reviewers' to upload the csv file from the Expert Lookup.
Click on 'get non reviewers' to upload the non reviewers.
Then close the window.

The script will run itself and then you will get an excel file of the filtered rreviewers in the same folder where you ran the file.
