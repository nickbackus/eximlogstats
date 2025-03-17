import re
import subprocess
import datetime
import os.path

FULL_PATH = "" # Path of the folder where you have this code, e.g. /home/[username]/EximLogStatsFolder

with open("emails.txt") as file:
    emails = [line.strip() for line in file]
    
today = str(datetime.datetime.now())[0:10]
finalData = ""

for email in emails:
    cmd = ["/sbin/exigrep", email, FULL_PATH + "/eximlog"]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    
    o, e = proc.communicate()
    
    if proc.returncode:
        print('Error: '  + e.decode('ascii'))
        print('code: ' + str(proc.returncode))
    
    
    parsed_results = re.split("\n\n", o.decode('ascii'))
    
    sent_count = 0
    received_count = 0
    
    for entry in parsed_results:
        if re.search(today, entry):
            if re.search("S=" + email, entry):
                sent_count += 1
            if re.search("for\s" + email, entry):
                received_count += 1
    
    finalData += today + "," + email + "," + str(sent_count) + "," + str(received_count) + "\n"

filepath = FULL_PATH + "/email_stats/email_stats_" + today[5:7] + "-" + today[0:4] + ".csv"

if os.path.exists(filepath):
    out = open (filepath, "a")   
else:
    out = open (filepath, "w")
    finalData = "Date,Email,Emails Sent,Emails Received\n" + finalData
out.write(finalData)
out.close()
