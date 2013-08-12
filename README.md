HackerSchoolDemo
================



Short program to read excel data on BNI Member performance/scorecard, do some computations to check if performance is to baseline, and create new excel file with new data.
 

Data input is done by another BNI member on a weekly basis and stored on Google Docs.

The program takes data collected on a weekly basis for performance measurement by person name, by week they were a member, and gives us a scorecard for that person.
A negative score means they are behind and a positive score means they are ahead by that amount.


~~~~~~~~~~~~
TO RUN IN PYTHON:

Please install modules xlrd and xlwt from http://www.python-excel.org/

In the python interpreter type:
$ python bnireport.py 	# and hit enter to run

The program will access the excel file from the same directory this README file is in and where the code is as well.

The program will also create a new excel file called "Performance" and save it in the same directory.
~~~~~~~~~

Thanks!
Nune