# Nova-Web-App

### This web application processes payroll for Nova Home Support. 

Inputs: shift record (.xlsx) from Therap and old tracker (.xlsx) maintained by the managers.

Outputs:

A. Full Cycle: a payroll output, a new tracker, an invoice (all .xlsx), and a payroll for machine (.csv).

B. Off Cycle: a payroll output (.xlsx) for selected staff and a payroll for machine (.csv).

Please follow the instructions on the webpage. 
Messages and warnings are displayed as alerts. 
The app saves your progress automatically.
To clear all files uploaded or generated, click "refresh" at the bottom.

This application is built using Flask under Python 3. 
AWS App Runner automatically deploys the program by pulling this GitHub repo. 
The server is set to use 0.5 Virtual CPU and 1 GB of storage. 
Finishing one process takes about 10-20 seconds.

The web app is running at https://wuabvtbqcm.us-west-2.awsapprunner.com/
