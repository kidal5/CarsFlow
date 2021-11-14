# CarsFlow

## First time setup
1. Download and install latest python. Be sure to check ```Add Python to PATH``` option. [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Download and unpack [this repository](https://github.com/kidal5/CarsFlow/archive/refs/heads/main.zip).
3. Install visual studio build tools. https://stackoverflow.com/questions/40504552/how-to-install-visual-c-build-tools
4. Open Command Prompt (cmd) in unpacked folder. (Right click on folder while holding **shift**, select open Command Prompt (Powershell) here.)
7. Type ```pip install -r requirements.txt``` into cmd and hit enter.

## Run program
1. Edit ```parameters.yaml``` file to match your needs. This yaml file is just ordinary text file, so it can be opened with every text processor. Just be aware, that yaml datatype uses indentation for data validation and wouldn't work if indentation is broken! I recommend using special text editor like [Notepadd++](https://notepad-plus-plus.org/downloads/) or [VS Code](https://code.visualstudio.com/Download) to handle correct indentation for you.  
3. Open cmd in working folder.
4. Run program by typing ```python .\main.py``` into cmd and hit enter.
5. Profit

## More yaml files
You can create more ```xxx.yaml``` files and switch between them, no need to have just one. To use this feature create new parameters file, ex. ```params.yaml```. Then run script with one optional parametr specified  ```python .\main.py -p params.yaml```.
