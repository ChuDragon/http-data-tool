# CBOE Put-Call Data Tool
### VBA on Excel To Automate Repetitive HTTP Requests With Parameters

## Functionality
It was developed for a client in VBA to download the CBOE put-call ratio historical data. 
This simple tool raches to the exchange web site and downloads the CBOE Put-Call Ratio for the S&P 500 for a specific date. 
The above is repeated in a loop to construct the put-call Ratio historical data series.
Versions of the tool can be useful to automate repetitively obtaining online data with multiple parameters (dates, multiple URLs, etc.), such as in finance or banking.

## Code Notes
Each HTTP "GET" request is syncronous and blocking (unfortunately, VBA doesn't support async), so the app might freeze for a moment while waiting for server response.  
The .xlsx workbook contains all the VBA code - just click the button to run, or modify the VBA macro. 
The app is also saved separately as .bas file. For custom modifications, contact me at the email in my profile.

## Setup/Dependencies
Microsoft XML, v6.0.
