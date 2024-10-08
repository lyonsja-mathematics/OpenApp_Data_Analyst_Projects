# In version 8 of the DataSHIELD tutorial we will login with a user with restricted access to the tables and run through some of the functions in the dsBaseClient, DSI and DSOpal packages using DataSHIELD's training servers and with tables which I will upload, these tables will have TrolleyGAR data for Beaumont and Navan. We will test out the basic DataSHIELD functions, custom functions added on the Opal servers and safeguards DataSHIED places on the confidentiality of individual level data.

# Clean up
rm(list=ls())

# Load libraries
#library(opal)
library(dsBaseClient)
library(DSI)
library(DSOpal)

## Build the login details, keep separate login data for each project
newDSLoginBuilder() -> builder
builder$append(server = "PROVIDER.A", url = "http://192.168.56.100:8080/", user = "Tester1", password = "OpenApp2020", driver = "OpalDriver", table = c("PET.PROVIDER_A"))
builder$append(server = "PROVIDER.B", url = "http://192.168.56.101:8080/", user = "Tester1", password = "OpenApp2020", driver = "OpalDriver", table = c('PET.PROVIDER_B'))
builder$build() -> pet.logindata; pet.logindata

# Log in for project PET
DSI::datashield.login(logins = pet.logindata, assign = TRUE, symbol = "D", variables = c('REGISTRATION','DEPART','AGE','DISCHARGE_CODE')) -> PET
datashield.connections_default('PET')

# We need to check if is_in are in the list of server side assign methods
subset(ds.listServersideFunctions()$serverside.assign.functions, is.na(package))

# Create a new variable, ADM, with value 1 if the discharge code is either 20, 30, 80 or 100
ds.assign(toAssign = 'is_in(D$DISCHARGE_CODE, c(20,30,80,100))', newobj = 'ADM')

# Take a subset of the data, records with discharge code 20, 30, 80, 100, i.e. the admissions
ds.dataFrameSubset(df.name = 'D', V1.name = 'ADM', Boolean.operator = '==', V2.name = '1', newobj = 'ADMISSIONS')
ds.dim(x = 'ADMISSIONS')

# Create another variable, ADM100, ADM multiplied by 100
ds.assign(toAssign = '100*ADM', newobj = 'ADM100')

# The mean of ADM100 is the admission rate
ds.mean(x = 'ADM100')

# We can demonstrate DS's ability to protect confidentiality of individual level data. In the BEAUMONT_PET table there are 17 rows with DISCHARGE_CODE = 40 but only 8 rows in NAVAN_PET; the admin user for the server can set the value of default.nfilter.subset from the default value of 3 to 10, these options prevent the analysis of small subsets of the data.
ds.dataFrameSubset(df.name = 'D', V1.name = 'D$DISCHARGE_CODE', Boolean.operator = '==', V2.name = '40', newobj = 'S1')
ds.dim(x = 'S1')

# Log out of project PET
DSI::datashield.logout(conns = PET)
  