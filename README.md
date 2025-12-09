# Solution overview
1. We are reading data from preview and storing it in an excel(can be viewed under output folder)
2. Reading data from prod and comparing it against the stored data in excel(under output) 
3. If any mismatches found we can see them under output->Mismatches

# Naming conventions 
1. Each folder under output holds data that is read from preview named with testdata 
2. Any file name which has Database indicates that is the data pulled after reload 

# To run locally 

1. Run `npm i` command to install all the dependencies
2. Run `npx playwright test` to run the script