# Nova-Web-App

### This web application processes payroll for Nova Home Support. 

Inputs: shift record (.xlsx) from Therap and old tracker (.xlsx) maintained by the managers.

Outputs: a payroll output, a new tracker, and an invoice (all .xlsx).

Please follow the instructions on the webpage. 
Messages and warnings are displayed on a new page. 
Click the back button on your browser to return to the main page. 
The app saves your progress automatically.
To clear all files uploaded or generated, click "refresh" at the bottom.

This application is built using Flask under Python 3. 
AWS App Runner automatically deploys the program by pulling this GitHub repo. 
The server is set to use 0.5 Virtual CPU and 1 GB of storage. 
Finishing one process takes about 10-20 seconds.

The web app is running at https://sm3vj8qpj6.us-east-1.awsapprunner.com/
