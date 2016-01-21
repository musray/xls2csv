from distutils.core import setup
import py2exe
includes = ["encodings", "encodings.*"]
options = {"py2exe":
            {   "compressed": 1,
                "optimize": 2,
                "includes": includes,
                "bundle_files": 1
            }
          }
setup(   
    version = "0.2.0",
    description = "convert xls files to csv",
    name = "xls2csv robot",
    options = options,
    zipfile=None,
    #console = [{"script": "simpleList.py", "icon_resources": [(1, "list.ico")] }],  
    windows = [{"script": "convert_gui.py", "icon_resources": [(1, "./icon/csv.ico")] }]
    #windows = [{"script": "convert_gui.py"}],  
    
    )
