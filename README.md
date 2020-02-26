# CognosCleanup
In the event of Cognos Reporting Breakdown, run this to validate Cognos numbers and clean up file


# Steps

1. Create a Conda environment by running conda create --name <env> --file requirements.txt
2. Activate Environment <env>
3. Run the batch file createTodayFolder.
4. Put the 13/14 input files from IT Team into the folder created.
5. Run batch file Run Code.
6. Check the log file to validate the numbers.
