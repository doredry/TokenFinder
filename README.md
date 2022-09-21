# TokenFinder

 ```
 _______      _                    _______  _             _               
(_______)    | |                  (_______)(_)           | |              
    _   ___  | |  _  _____  ____   _____    _  ____    __| | _____   ____ 
   | | / _ \ | |_/ )| ___ ||  _ \ |  ___)  | ||  _ \  / _  || ___ | / ___)
   | || |_| ||  _ ( | ____|| | | || |      | || | | |( (_| || ____|| |    
   |_| \___/ |_| \_)|_____)|_| |_||_|      |_||_| |_| \____||_____)|_|  
```

Tool to extract powerful tokens from Office desktop apps memory. 
The tool will iterate over all the running processes, find the relevant Office desktop apps processes and will extract only the useable tokens into a .txt file. 

## Install 
```
pip install pywin32
pip install psutil
```
## Usage
default mode, iterate over all running processes and search for relevant ones
```
python3 TokenFinder.py 
```
provide specific PID's to extract tokens from
```
python3 TokenFinder -p 1543 5454 1050 
```


credit to minidump implementation: [link](https://stackoverflow.com/questions/21401663/python-windows-api-dump-process-in-buffer-then-regex-search)
