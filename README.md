# msc-schedule-generator
A program to automate tutor scheduling for the Math Skills Center at college.

The program works in three stages: 

First, the student workers submit the Tutor Availability Form with their specified work availability. All these files are placed in the specified folder.

Second, the script runs. It uses simulated annealing to develop a reasonable solution that tries to schedule all according to their first and second priorities. It uses third priorities if nobody else is available. MSC Tutor Schedule.xlsx is the final schedule from the annealing process. It saves a report file, describing the resulting schedule, as well as a collection of all availability constraints, for the professors use next. 

Third, the professor refines the resulting draft, before making it official. 


The program uses a configuration spreadsheet: 
MSC Hours of Operation.xlsx holds the hours that the MSC is open, as well as how many students work each shift. It stores seperate hours for the seperate locations.

Other files within the repository were for testing and redacting purposes. 

