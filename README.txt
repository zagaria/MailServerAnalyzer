How it work

1. Run script with parametr database.txt
Example: Mail.ps1 Data.txt

2. You can set bulk and sleep (in seconds) option (Default: 20 bulk and 1 second sleep)
Example: Mail.ps1 Data.txt 10 1

3. Result
You have 3 file:
Base.csv - result in format: email,mx,server
cache.txt - caching detect
config.txt - last position
log.txt - log file with detail work
stat.txt - file with Statictics
Data.txt - sample for test